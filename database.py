"""
数据库：SQLite + FTS5 全文搜索。
表：rows（每行数据 + 公式等）、images（图片路径）、rows_fts（全文索引）。
"""
import json
import sqlite3
from pathlib import Path
from typing import Any, Optional

import config


def get_connection():
    conn = sqlite3.connect(config.DB_PATH)
    # 读多场景下加速
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA cache_size=-64000")  # 64MB 缓存
    conn.execute("PRAGMA temp_store=MEMORY")
    return conn


def init_db():
    """创建表及 FTS5 全文索引。"""
    conn = get_connection()
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS rows (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sheet_name TEXT NOT NULL,
            row_index INTEGER NOT NULL,
            column_data TEXT NOT NULL,
            formula_text TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS images (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            row_id INTEGER NOT NULL,
            col_ref TEXT,
            image_path TEXT NOT NULL,
            alt_text TEXT,
            FOREIGN KEY (row_id) REFERENCES rows(id)
        )
    """)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_rows_sheet ON rows(sheet_name)")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_images_row ON images(row_id)")

    # 搜索列表用轻量预览列，避免读大字段 column_data
    try:
        conn.execute("ALTER TABLE rows ADD COLUMN preview TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass
    # LIKE 备用搜索用：前若干列拼成一段文本，便于按商品型号等检索
    try:
        conn.execute("ALTER TABLE rows ADD COLUMN search_text TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass

    # FTS5 全文搜索（若已存在则跳过）
    conn.execute("""
        CREATE VIRTUAL TABLE IF NOT EXISTS rows_fts USING fts5(
            sheet_name,
            column_data,
            formula_text,
            content='rows',
            content_rowid='id'
        )
    """)
    try:
        conn.execute("""
            CREATE TRIGGER IF NOT EXISTS rows_ai AFTER INSERT ON rows BEGIN
                INSERT INTO rows_fts(rowid, sheet_name, column_data, formula_text)
                VALUES (new.id, new.sheet_name, new.column_data, new.formula_text);
            END
        """)
        conn.execute("""
            CREATE TRIGGER IF NOT EXISTS rows_ad AFTER DELETE ON rows BEGIN
                INSERT INTO rows_fts(rows_fts, rowid, sheet_name, column_data, formula_text)
                VALUES ('delete', old.id, old.sheet_name, old.column_data, old.formula_text);
            END
        """)
        conn.execute("""
            CREATE TRIGGER IF NOT EXISTS rows_au AFTER UPDATE ON rows BEGIN
                INSERT INTO rows_fts(rows_fts, rowid, sheet_name, column_data, formula_text)
                VALUES ('delete', old.id, old.sheet_name, old.column_data, old.formula_text);
                INSERT INTO rows_fts(rowid, sheet_name, column_data, formula_text)
                VALUES (new.id, new.sheet_name, new.column_data, new.formula_text);
            END
        """)
    except sqlite3.OperationalError:
        pass  # 可能 FTS 表已存在但 trigger 不同
    conn.commit()
    conn.close()


def insert_row(cursor, sheet_name: str, row_index: int, column_data: list, formula_text: Optional[str] = None, preview: Optional[str] = None, search_text: Optional[str] = None):
    col_json = json.dumps(column_data, ensure_ascii=False)
    if preview is None:
        preview = json.dumps(column_data[:5] if column_data else [], ensure_ascii=False)[:2000]
    if search_text is None:
        search_text = " ".join(str(c) for c in (column_data[:15] if column_data else []))[:5000]
    try:
        cursor.execute(
            "INSERT INTO rows (sheet_name, row_index, column_data, formula_text, preview, search_text) VALUES (?,?,?,?,?,?)",
            (sheet_name, row_index, col_json, formula_text or "", preview, search_text),
        )
    except sqlite3.OperationalError:
        try:
            cursor.execute(
                "INSERT INTO rows (sheet_name, row_index, column_data, formula_text, preview) VALUES (?,?,?,?,?)",
                (sheet_name, row_index, col_json, formula_text or "", preview),
            )
        except sqlite3.OperationalError:
            cursor.execute(
                "INSERT INTO rows (sheet_name, row_index, column_data, formula_text) VALUES (?,?,?,?)",
                (sheet_name, row_index, col_json, formula_text or ""),
            )
    return cursor.lastrowid


def insert_image(cursor, row_id: int, image_path: str, col_ref: Optional[str] = None, alt_text: Optional[str] = None):
    cursor.execute(
        "INSERT INTO images (row_id, col_ref, image_path, alt_text) VALUES (?,?,?,?)",
        (row_id, col_ref or "", image_path, alt_text or ""),
    )


def _fts_escape(q: str) -> str:
    """FTS5 短语查询转义：用双引号包起来，内部双引号改为两个双引号。"""
    q = (q or "").strip()
    if not q:
        return ""
    return '"' + q.replace('"', '""') + '"'


def _search_ids_like(conn, query: str, limit: int) -> list[int]:
    """用 LIKE 在 search_text / preview / formula_text / column_data 里搜，适合商品型号、中文等。"""
    q = (query or "").strip()
    if not q:
        return []
    like = "%" + q + "%"
    try:
        cur = conn.execute(
            "SELECT id FROM rows WHERE COALESCE(search_text,'') LIKE ? OR COALESCE(preview,'') LIKE ? OR COALESCE(formula_text,'') LIKE ? OR COALESCE(column_data,'') LIKE ? LIMIT ?",
            (like, like, like, like, limit),
        )
    except sqlite3.OperationalError:
        try:
            cur = conn.execute(
                "SELECT id FROM rows WHERE COALESCE(preview,'') LIKE ? OR COALESCE(formula_text,'') LIKE ? OR COALESCE(column_data,'') LIKE ? LIMIT ?",
                (like, like, like, limit),
            )
        except sqlite3.OperationalError:
            cur = conn.execute(
                "SELECT id FROM rows WHERE COALESCE(preview,'') LIKE ? OR COALESCE(formula_text,'') LIKE ? LIMIT ?",
                (like, like, limit),
            )
    return [r[0] for r in cur.fetchall()]


def search_fts(conn, query: str, limit: int = 100) -> list[dict]:
    """全文搜索，返回 row id 列表（可再查 rows 表取详情）。"""
    if not query.strip():
        return []
    cur = conn.execute(
        "SELECT rowid FROM rows_fts WHERE rows_fts MATCH ? LIMIT ?",
        (query.strip(), limit),
    )
    return [{"id": r[0]} for r in cur.fetchall()]


def _build_preview_results(conn, ids: list[int]) -> list[dict]:
    """根据 id 列表批量拼出列表项（与 search_with_preview 返回格式一致）。"""
    if not ids:
        return []
    placeholders = ",".join("?" * len(ids))
    cur = conn.execute(
        f"SELECT id, sheet_name, row_index, COALESCE(preview,'[]'), substr(COALESCE(formula_text,''),1,100) FROM rows WHERE id IN ({placeholders})",
        ids,
    )
    rows_by_id = {}
    for r in cur.fetchall():
        try:
            col_preview = json.loads(r[3]) if r[3] else []
            if not isinstance(col_preview, list):
                col_preview = [str(col_preview)]
        except (TypeError, json.JSONDecodeError):
            col_preview = [str(r[3])[:200]] if r[3] else []
        rows_by_id[r[0]] = {
            "id": r[0],
            "sheet_name": r[1],
            "row_index": r[2],
            "column_preview": col_preview[:10],
            "formula_preview": (r[4] or "").strip(),
        }
    cur = conn.execute(
        f"SELECT row_id, COUNT(*) FROM images WHERE row_id IN ({placeholders}) GROUP BY row_id",
        ids,
    )
    image_count_by_id = {r[0]: r[1] for r in cur.fetchall()}
    results = []
    for row_id in ids:
        if row_id not in rows_by_id:
            continue
        row = rows_by_id[row_id]
        row["image_count"] = image_count_by_id.get(row_id, 0)
        results.append(row)
    return results


def search_with_preview(conn, query: str, limit: int = 100) -> list[dict]:
    """
    支持多关键词：query 内换行或逗号分隔的多个词，任意一个匹配即命中（OR）。
    FTS 与 LIKE 结果合并去重，避免只返回一条。
    """
    raw = (query or "").replace(",", "\n").replace("，", "\n")
    terms = [t.strip() for t in raw.splitlines() if t.strip()]
    if not terms:
        return []
    seen = set()
    all_ids = []
    for q in terms:
        if len(all_ids) >= limit:
            break
        ids_fts = []
        fts_query = _fts_escape(q)
        if fts_query:
            try:
                cur = conn.execute(
                    "SELECT rowid FROM rows_fts WHERE rows_fts MATCH ? LIMIT ?",
                    (fts_query, limit),
                )
                ids_fts = [r[0] for r in cur.fetchall()]
            except (sqlite3.OperationalError, sqlite3.ProgrammingError):
                pass
        ids_like = _search_ids_like(conn, q, limit)
        for i in ids_fts:
            if i not in seen:
                seen.add(i)
                all_ids.append(i)
                if len(all_ids) >= limit:
                    break
        if len(all_ids) >= limit:
            break
        for i in ids_like:
            if i not in seen:
                seen.add(i)
                all_ids.append(i)
                if len(all_ids) >= limit:
                    break
    return _build_preview_results(conn, all_ids[:limit])


def get_row_ids_by_models(conn, models: list[str], limit_per_term: int = 500) -> list[int]:
    """按多个型号/关键词查 row id，任意一个匹配即命中（OR），去重后按出现顺序返回。会尝试原词、小写、大写以兼容大小写差异。"""
    if not models:
        return []
    seen = set()
    result = []
    for q in models:
        term = (q or "").strip()
        if not term:
            continue
        # 先试原词，再试小写/大写（避免数据里大小写不一致导致查不到）
        variants = [term]
        low, up = term.lower(), term.upper()
        if low != term:
            variants.append(low)
        if up != term and up not in variants:
            variants.append(up)
        for t in variants:
            ids = []
            fts_query = _fts_escape(t)
            if fts_query:
                try:
                    cur = conn.execute(
                        "SELECT rowid FROM rows_fts WHERE rows_fts MATCH ? LIMIT ?",
                        (fts_query, limit_per_term),
                    )
                    ids = [r[0] for r in cur.fetchall()]
                except (sqlite3.OperationalError, sqlite3.ProgrammingError):
                    pass
            if not ids:
                ids = _search_ids_like(conn, t, limit_per_term)
            for i in ids:
                if i not in seen:
                    seen.add(i)
                    result.append(i)
            if ids:
                break
    return result


def get_full_rows_by_ids(conn, ids: list[int]) -> list[dict]:
    """根据 id 列表返回完整行（id, sheet_name, row_index, column_data），保持 ids 顺序。"""
    if not ids:
        return []
    placeholders = ",".join("?" * len(ids))
    cur = conn.execute(
        f"SELECT id, sheet_name, row_index, column_data FROM rows WHERE id IN ({placeholders})",
        ids,
    )
    by_id = {}
    for r in cur.fetchall():
        try:
            col_data = json.loads(r[3]) if r[3] else []
        except (TypeError, json.JSONDecodeError):
            col_data = []
        by_id[r[0]] = {"id": r[0], "sheet_name": r[1], "row_index": r[2], "column_data": col_data}
    return [by_id[i] for i in ids if i in by_id]


def get_images_by_row_ids(conn, ids: list[int]) -> dict:
    """根据行 id 列表返回每行对应的图片列表。返回 {row_id: [{"col_ref": "col_0", "image_path": "..."}, ...]}。"""
    if not ids:
        return {}
    placeholders = ",".join("?" * len(ids))
    cur = conn.execute(
        f"SELECT row_id, col_ref, image_path FROM images WHERE row_id IN ({placeholders})",
        ids,
    )
    by_row = {}
    for r in cur.fetchall():
        rid = int(r[0]) if r[0] is not None else None
        if rid is None:
            continue
        if rid not in by_row:
            by_row[rid] = []
        by_row[rid].append({"col_ref": r[1] or "", "image_path": r[2] or ""})
    return by_row


def get_sheet_header_labels(conn, sheet_name: str) -> list:
    """取指定工作表第一行（row_index=1）的 column_data，用作列名。"""
    cur = conn.execute(
        "SELECT column_data FROM rows WHERE sheet_name = ? AND row_index = 1 LIMIT 1",
        (sheet_name,),
    )
    r = cur.fetchone()
    if not r or not r[0]:
        return []
    try:
        data = json.loads(r[0])
        return data if isinstance(data, list) else []
    except (TypeError, json.JSONDecodeError):
        return []


def get_row_by_id(conn, row_id: int) -> Optional[dict]:
    cur = conn.execute("SELECT id, sheet_name, row_index, column_data, formula_text FROM rows WHERE id = ?", (row_id,))
    r = cur.fetchone()
    if not r:
        return None
    cur2 = conn.execute("SELECT id, col_ref, image_path, alt_text FROM images WHERE row_id = ?", (row_id,))
    images = [{"id": i[0], "col_ref": i[1], "path": i[2], "alt": i[3]} for i in cur2.fetchall()]
    return {
        "id": r[0],
        "sheet_name": r[1],
        "row_index": r[2],
        "column_data": json.loads(r[3]) if r[3] else [],
        "formula_text": r[4] or "",
        "images": images,
    }


def clear_all_rows(conn):
    """清空所有行与图片，用于重新导入时避免重复。"""
    conn.execute("DELETE FROM images")
    conn.execute("DELETE FROM rows")
    conn.commit()


def get_sheets(conn) -> list[str]:
    cur = conn.execute("SELECT DISTINCT sheet_name FROM rows ORDER BY sheet_name")
    return [r[0] for r in cur.fetchall()]


def get_sheet_counts(conn) -> list[tuple[str, int]]:
    """返回 (sheet_name, row_count) 列表，用于校验数据。"""
    cur = conn.execute(
        "SELECT sheet_name, COUNT(*) FROM rows GROUP BY sheet_name ORDER BY sheet_name"
    )
    return [(r[0], r[1]) for r in cur.fetchall()]


def get_sheet_total(conn, sheet_name: str) -> int:
    cur = conn.execute("SELECT COUNT(*) FROM rows WHERE sheet_name = ?", (sheet_name,))
    return cur.fetchone()[0]


def get_sheet_rows(conn, sheet_name: str, offset: int = 0, limit: int = 100, order: str = "desc") -> list[dict]:
    """按工作表分页取行。order=desc 时：表头(row_index=1)置顶，其余按 row_index 倒序。"""
    order = (order or "desc").strip().lower()
    if order not in ("asc", "desc"):
        order = "desc"

    def _load(ids_and_rows):
        rows = []
        for r in ids_and_rows:
            try:
                col_data = json.loads(r[2]) if r[2] else []
            except (TypeError, json.JSONDecodeError):
                col_data = []
            row_id = int(r[0]) if r[0] is not None else None
            rows.append({"id": row_id, "row_index": r[1], "column_data": col_data, "images": []})
        if not rows:
            return rows
        ids = [row["id"] for row in rows if row["id"] is not None]
        if not ids:
            return rows
        placeholders = ",".join("?" * len(ids))
        cur = conn.execute(
            f"SELECT row_id, col_ref, image_path FROM images WHERE row_id IN ({placeholders})",
            ids,
        )
        row_images = {}
        for r in cur.fetchall():
            rid, col_ref, path = int(r[0]) if r[0] is not None else None, r[1], r[2]
            if rid is not None:
                if rid not in row_images:
                    row_images[rid] = []
                row_images[rid].append({"col_ref": col_ref or "", "path": path or ""})
        for row in rows:
            row["images"] = row_images.get(int(row["id"]) if row["id"] is not None else None, []) or []
        return rows

    if order == "desc":
        # 第 0 页：表头( row_index=1 ) + 倒序数据前 (limit-1) 条
        if offset == 0:
            cur = conn.execute(
                "SELECT id, row_index, column_data FROM rows WHERE sheet_name = ? AND row_index = 1",
                (sheet_name,),
            )
            header = cur.fetchone()
            cur = conn.execute(
                "SELECT id, row_index, column_data FROM rows WHERE sheet_name = ? AND row_index >= 2 ORDER BY CAST(row_index AS INTEGER) DESC LIMIT ? OFFSET ?",
                (sheet_name, limit - 1, 0),
            )
            data_rows = cur.fetchall()
            if header:
                rows = [header] + list(data_rows)
            else:
                rows = list(data_rows)
            return _load(rows)
        # 第 1 页起：倒序数据的后续页，每页 limit 条
        data_offset = (limit - 1) + (offset - 1) * limit
        cur = conn.execute(
            "SELECT id, row_index, column_data FROM rows WHERE sheet_name = ? AND row_index >= 2 ORDER BY CAST(row_index AS INTEGER) DESC LIMIT ? OFFSET ?",
            (sheet_name, limit, data_offset),
        )
        return _load(cur.fetchall())

    cur = conn.execute(
        "SELECT id, row_index, column_data FROM rows WHERE sheet_name = ? ORDER BY row_index LIMIT ? OFFSET ?",
        (sheet_name, limit, offset),
    )
    return _load(cur.fetchall())
