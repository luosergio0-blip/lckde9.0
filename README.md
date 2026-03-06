# Excel 查询站

针对**超大 Excel 文件**（约 3GB）的查询网站：先将数据导入 SQLite，再通过网页全文搜索、查看详情（含公式、图片），并支持添加新内容。

## 功能

- **全文搜索**：对表格内容、公式等做关键词搜索（SQLite FTS5）
- **详情展示**：查看每行的列数据、公式文本；支持 LaTeX 公式渲染（KaTeX）、图片展示
- **添加内容**：在网页上新增一行到指定 Sheet，无需改 Excel

## 架构说明

- **大文件策略**：不把 3GB Excel 全部读入内存。使用 `openpyxl` 的 **read_only 模式**流式读取，按行写入 SQLite，内存占用可控。
- **存储**：SQLite 数据库（`excel_data.db`）+ 图片导出到 `static/images/`
- **后端**：FastAPI，提供搜索、详情、新增接口
- **前端**：单页 HTML（`static/index.html`），含搜索框、结果列表、详情页、添加表单

## 环境要求

- Python 3.10+
- 磁盘空间：至少能放下 Excel + 数据库 + 导出图片

## 快速开始

### 1. 安装依赖

```bash
cd "c:\Users\Administrator\Desktop\9.0网页计划"
pip install -r requirements.txt
```

### 2. 配置 Excel 路径

任选其一：

- 把 Excel 放到项目目录下，命名为 `data.xlsx`；或  
- 设置环境变量（推荐大文件用绝对路径）：

```powershell
$env:EXCEL_PATH = "D:\你的路径\你的大文件.xlsx"
```

也可直接改 `config.py` 里的 `EXCEL_PATH`。

### 3. 导入 Excel 到数据库（首次必须执行）

```bash
python import_excel.py
```

- 会流式读取 Excel，按批写入 SQLite，并尝试导出图片到 `static/images/`。  
- 3GB 文件可能需要较长时间，请耐心等待；控制台会按批打印进度。

### 4. 启动网站

```bash
uvicorn main:app --host 0.0.0.0 --port 8000
```

浏览器访问：**http://localhost:8000**

- **搜索**：在搜索框输入关键词，回车或点「搜索」  
- **查看详情**：点击某条结果，可看该行全部列、公式、图片  
- **添加内容**：切到「添加内容」标签，填写 Sheet 名、列数据（可 JSON 数组或每行一列）、公式文本后提交

## 项目结构

```
9.0网页计划/
├── config.py          # 配置：EXCEL_PATH、DB_PATH、图片目录等
├── database.py        # SQLite 表结构、FTS5、增删查
├── import_excel.py    # 流式导入 Excel → SQLite + 图片
├── main.py            # FastAPI 应用与 API
├── requirements.txt
├── README.md
├── static/
│   ├── index.html     # 前端页面
│   └── images/        # 从 Excel 导出的图片（导入时生成）
├── excel_data.db      # 运行后生成的数据库（可不在版本控制）
└── data.xlsx          # 你的 Excel（或通过 EXCEL_PATH 指定）
```

## API 简要说明

| 接口 | 说明 |
|------|------|
| `GET /api/sheets` | 获取所有 Sheet 名称 |
| `GET /api/search?q=关键词&limit=100` | 全文搜索，返回匹配行摘要 |
| `GET /api/row/{id}` | 获取单行详情（列数据、公式、图片） |
| `POST /api/row` | 新增一行（body: `sheet_name`, `column_data`, `formula_text`） |
| `POST /api/import` | 手动触发重新导入（同步，大文件慎用） |

## 公式与图片说明

- **公式**：导入时保留 Excel 公式字符串（`data_only=False`），详情页会原样显示；若公式文本中含 LaTeX（如 `$E=mc^2$`），会用 KaTeX 渲染。
- **图片**：openpyxl 在 read_only 模式下对图片支持有限，能解析的会导出到 `static/images/` 并在详情中显示；无法解析的不会报错，仅该行无图。

## 注意事项

1. **首次使用**：必须先运行 `import_excel.py` 生成数据库，再启动网站。
2. **Excel 更新后**：重新执行 `python import_excel.py` 会**覆盖**现有数据库（当前为全量重建）。若需“增量”或“追加”，需自行改 `import_excel.py` 逻辑。
3. **新增内容**：通过网页「添加内容」只会写入 SQLite，不会回写 Excel；Excel 仍可作为原始数据源，定期重新导入以覆盖或合并策略需自行设计。

按上述步骤即可在本地搭建并查询你的大 Excel，并支持后续添加新内容。
