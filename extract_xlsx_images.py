"""
从 xlsx（zip）内提取 xl/media 图片，并解析 drawing 得到 (sheet_name, row, col) -> 图片路径。
用于导入时把图片写入 DB，网页即可显示。
"""
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import config


# 常用命名空间（xlsx drawing）
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}


def _ns(tag, prefix="xdr"):
    return "{%s}%s" % (NS.get(prefix, NS["xdr"]), tag)


def _text(el, default=""):
    if el is None:
        return default
    return (el.text or "").strip() or default


def _int(el, default=0):
    try:
        return int(_text(el, "0"))
    except ValueError:
        return default


def _itertext(el):
    if el.text:
        yield el.text
    for c in el:
        yield from _itertext(c)
        if c.tail:
            yield c.tail


def _cell_ref_to_row_col(r: str) -> tuple[int, int] | None:
    """将 A1、B2、AA10 等转为 (row_0based, col_0based)。"""
    if not r or not r.strip():
        return None
    r = r.strip().upper()
    col_s = ""
    row_s = ""
    for c in r:
        if c.isalpha():
            col_s += c
        elif c.isdigit():
            row_s += c
        else:
            break
    if not row_s:
        return None
    try:
        row_1based = int(row_s)
        row_0 = max(0, row_1based - 1)
    except ValueError:
        return None
    col_0 = 0
    for ch in col_s:
        col_0 = col_0 * 26 + (ord(ch) - ord("A") + 1)
    col_0 = max(0, col_0 - 1)
    return (row_0, col_0)


def _extract_cell_picture_cells(z, namelist, ws_path_to_name, sheet_order: list[str]) -> list[tuple[str, int, int]]:
    """收集所有含 DISPIMG/IMAGE 公式的单元格 (sheet_name, row_0, col_0)，按 sheet_order 与行列排序。"""
    out = []
    # 按 workbook 中 sheet 顺序遍历，以便与 xl/media 顺序更一致
    name_to_path = {}
    for n in namelist:
        if n.startswith("xl/worksheets/") and n.endswith(".xml") and "._rels" not in n:
            key = n[3:] if n.startswith("xl/") else n
            sn = ws_path_to_name.get(key) or ws_path_to_name.get(n)
            if sn:
                name_to_path[sn] = n
    for sheet_name in sheet_order:
        path = name_to_path.get(sheet_name)
        if not path:
            continue
        try:
            sheet_root = ET.fromstring(z.read(path))
        except Exception:
            continue
        for c in sheet_root.iter():
            if (c.tag or "").split("}")[-1] != "c":
                continue
            r = c.get("r")
            if not r:
                continue
            rc = _cell_ref_to_row_col(r)
            if rc is None:
                continue
            row_0, col_0 = rc
            has_img_formula = False
            for child in c:
                if (child.tag or "").split("}")[-1] == "f":
                    text = (child.text or "") + "".join(_itertext(child))
                    if "DISPIMG" in text.upper() or ("IMAGE" in text.upper() and "IMAGE(" in text.upper().replace(" ", "")):
                        has_img_formula = True
                        break
            if has_img_formula:
                out.append((sheet_name, row_0, col_0))
    out.sort(key=lambda x: (x[0], x[1], x[2]))
    return out


def _extract_all_media_to_folder(z, namelist) -> list[str]:
    """把 xl/media 下所有文件解出到 static/images，返回 [路径, ...] 按文件名排序。"""
    media_files = sorted(n for n in namelist if "xl/media" in n and not n.rstrip("/").endswith("media"))
    paths = []
    for i, mpath in enumerate(media_files):
        try:
            data = z.read(mpath)
        except Exception:
            continue
        ext = Path(mpath).suffix or ".png"
        out_name = f"cellimg_{i}{ext}"
        out_path = config.IMAGES_DIR / out_name
        out_path.write_bytes(data)
        paths.append(f"/static/images/{out_name}")
    return paths


def extract_images_from_xlsx(excel_path: str) -> dict[tuple[str, int, int], str]:
    """
    返回 map: (sheet_name, row_0based, col_0based) -> "/static/images/xxx.png"
    """
    result = {}
    excel_path = Path(excel_path)
    if not excel_path.is_file() or excel_path.suffix.lower() not in (".xlsx", ".xlsm"):
        return result

    config.IMAGES_DIR.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(excel_path, "r") as z:
        namelist = z.namelist()

        # 1) workbook.xml.rels: rId -> Target (path relative to xl/)
        rels_path = "xl/_rels/workbook.xml.rels"
        if rels_path not in namelist:
            return result
        rels_root = ET.fromstring(z.read(rels_path))
        rid_to_target = {}
        for rel in rels_root.iter():
            if (rel.tag or "").split("}")[-1] == "Relationship":
                rid = rel.get("Id")
                target = rel.get("Target", "")
                if rid:
                    rid_to_target[rid] = target

        # 2) workbook.xml: sheet rId -> name (order = sheet index)
        wb_path = "xl/workbook.xml"
        if wb_path not in namelist:
            return result
        wb_root = ET.fromstring(z.read(wb_path))
        sheet_order = []
        for s in wb_root.iter():
            if (s.tag or "").split("}")[-1] != "sheet":
                continue
            name = s.get("name") or ""
            rid = ""
            for k, v in (s.attrib or {}).items():
                if k.endswith("}id") or k == "r:id" or (("id" in k.lower()) and ("relationship" in k.lower() or "r:" in k)):
                    rid = v
                    break
            if not rid:
                rid = s.get("r:id") or s.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") or ""
            if name and rid:
                sheet_order.append((rid, name))

        # worksheet path (e.g. worksheets/sheet1.xml) -> sheet name
        ws_path_to_name = {}
        for rid, name in sheet_order:
            t = rid_to_target.get(rid, "")
            if t:
                ws_path_to_name[t.lstrip("/")] = name
                if t.startswith("xl/"):
                    ws_path_to_name[t[3:]] = name  # worksheets/sheet1.xml

        # 3) 遍历 xl/worksheets/sheet*.xml，找 drawing
        for path in namelist:
            if not path.startswith("xl/worksheets/") or not path.endswith(".xml") or "._rels" in path:
                continue
            # path = xl/worksheets/sheet1.xml -> key worksheets/sheet1.xml
            key = path[3:] if path.startswith("xl/") else path
            sheet_name = ws_path_to_name.get(key) or ws_path_to_name.get(path)
            if not sheet_name:
                continue
            try:
                sheet_root = ET.fromstring(z.read(path))
            except Exception:
                continue
            # <drawing r:id="rId1"/>
            drawing = None
            for el in sheet_root.iter():
                if (el.tag or "").split("}")[-1] == "drawing":
                    drawing = el
                    break
            if drawing is None:
                continue
            draw_rid = None
            for k, v in (drawing.attrib or {}).items():
                if "embed" in k.lower() or k.endswith("}id"):
                    draw_rid = v
                    break
            draw_rid = draw_rid or drawing.get("r:id") or drawing.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if not draw_rid:
                continue
            # sheet rels: rId -> drawings/drawing1.xml
            sheet_rels_path = "xl/worksheets/_rels/" + Path(path).name + ".rels"
            if sheet_rels_path not in namelist:
                continue
            sheet_rels_root = ET.fromstring(z.read(sheet_rels_path))
            draw_target = None
            for rel in sheet_rels_root.iter():
                if (rel.tag or "").split("}")[-1] == "Relationship" and rel.get("Id") == draw_rid:
                    draw_target = rel.get("Target", "")
                    break
            if not draw_target:
                continue
            # drawing path: drawings/drawing1.xml 或 ../drawings/drawing1.xml -> xl/drawings/drawing1.xml
            if draw_target.startswith("xl/"):
                draw_path = draw_target
            elif draw_target.startswith("../"):
                draw_path = "xl/" + draw_target[3:]
            else:
                draw_path = "xl/" + draw_target
            if draw_path not in namelist:
                for n in namelist:
                    if "drawings" in n and Path(draw_target).name in n:
                        draw_path = n
                        break
                else:
                    continue
            if draw_path not in namelist:
                continue
            try:
                draw_root = ET.fromstring(z.read(draw_path))
            except Exception:
                continue
            # drawing rels: blip rId -> ../media/image1.png
            draw_rels_path = "xl/drawings/_rels/" + Path(draw_path).name + ".rels"
            if draw_rels_path not in namelist:
                continue
            draw_rels_root = ET.fromstring(z.read(draw_rels_path))
            blip_rid_to_media = {}
            for rel in draw_rels_root.iter():
                if (rel.tag or "").split("}")[-1] != "Relationship":
                    continue
                rid = rel.get("Id")
                target = rel.get("Target", "")
                if rid and target and "media" in target:
                    mp = "xl/" + target.lstrip("./") if not target.startswith("xl/") else target
                    if not mp.startswith("xl/"):
                        mp = "xl/media/" + Path(target).name
                    blip_rid_to_media[rid] = mp
            # oneCellAnchor / twoCellAnchor: from -> row, col; blip -> rId（用 local name 兼容各种前缀）
            def find_attr(el, attr):
                for k, v in (el.attrib or {}).items():
                    if "embed" in k.lower() or k.endswith("}embed"):
                        return v
                return None

            def find_child(parent, local_name):
                if parent is None:
                    return None
                for c in parent:
                    if (getattr(c, "tag", "") or "").split("}")[-1] == local_name:
                        return c
                return None

            def find_children(parent, local_name):
                out = []
                for c in parent.iter():
                    if (getattr(c, "tag", "") or "").split("}")[-1] == local_name:
                        out.append(c)
                return out

            for anchor in find_children(draw_root, "oneCellAnchor") + find_children(draw_root, "twoCellAnchor"):
                from_ = find_child(anchor, "from")
                if from_ is None:
                    continue
                row_el = find_child(from_, "row")
                col_el = find_child(from_, "col")
                row_0 = _int(row_el)
                col_0 = _int(col_el)
                blip = find_child(anchor, "blip")
                embed = find_attr(blip, "embed") if blip is not None else None
                if not embed:
                    continue
                media_rel = blip_rid_to_media.get(embed)
                if not media_rel:
                    continue
                media_path = media_rel if media_rel.startswith("xl/") else "xl/media/" + Path(media_rel).name
                if media_path not in namelist:
                    for n in namelist:
                        if "media" in n and (Path(media_path).name in n or Path(media_rel).name in n):
                            media_path = n
                            break
                    else:
                        continue
                try:
                    data = z.read(media_path)
                except Exception:
                    continue
                ext = Path(media_path).suffix or ".png"
                safe_name = re.sub(r"[^\w\-.]", "_", sheet_name)[:30]
                out_name = f"import_{safe_name}_r{row_0}_c{col_0}_{len(result)}{ext}"
                out_path = config.IMAGES_DIR / out_name
                out_path.write_bytes(data)
                result[(sheet_name, row_0, col_0)] = f"/static/images/{out_name}"

        # 4) 单元格内图片（DISPIMG/IMAGE 公式）：收集公式单元格，与 xl/media 按顺序对应
        sheet_names_in_order = [name for _, name in sheet_order]
        cell_picture_cells = _extract_cell_picture_cells(z, namelist, ws_path_to_name, sheet_names_in_order)
        if cell_picture_cells:
            media_paths = _extract_all_media_to_folder(z, namelist)
            n = min(len(cell_picture_cells), len(media_paths))
            for i in range(n):
                key = cell_picture_cells[i]
                if key not in result:
                    result[key] = media_paths[i]
    return result
