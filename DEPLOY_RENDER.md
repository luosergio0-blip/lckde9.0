# 在 Render 上部署本项目的步骤（你需要配合的部分）

本文说明如何把本项目部署到 [Render](https://render.com) 的 Web Service，以及你需要做的操作。

---

## 一、部署前准备（你需要做的）

### 1. 代码放到 Git 仓库

- 本仓库地址：**https://github.com/luosergio0-blip/lckde9.0.git**（已配置为 `origin`，可直接推送）。
- 在 Render 连接 GitHub 后，选择仓库 **luosergio0-blip/lckde9.0** 即可。
- 若项目不在仓库根目录，记下**子目录路径**，后面在 Render 里要填 **Root Directory**（本项目在根目录则留空）。

### 2. 是否要“带数据”上线（大文件才做网站，部署也要有数据）

- **当前方案使用 SQLite**，Render 免费实例**没有持久磁盘**，每次重新部署或实例重启后，数据库会清空。
- **若需要别人访问全部数据**（不是一部分），用 **Git LFS** 推送完整 `excel_data.db`（约 115MB）：
  - **步骤一**：安装 [Git LFS](https://git-lfs.com/)（Windows 可下载安装包或 `winget install GitHub.GitLFS`），安装后在项目目录执行一次 **`git lfs install`**。
  - **步骤二**：本地**完整导入**（不要设 `IMPORT_MAX_ROWS`）：**`python import_excel.py`**，得到完整的 `excel_data.db` 和 `static/images`。
  - **步骤三**：提交并推送（项目里已配好 `.gitattributes`，`excel_data.db` 会走 LFS，不会触发 GitHub 100MB 限制）：
    ```bash
    git add .gitattributes
    git add -f excel_data.db
    git add static/
    git add .
    git commit -m "完整数据部署：excel_data.db 通过 LFS"
    git push -u origin main
    ```
  - Render 部署时会自动拉取 LFS 文件，**部署后别人即可访问全部数据**。
- **若只希望带“部分数据”**（例如前几千行、库小于 100MB）：可用 **`IMPORT_MAX_ROWS=3000 python import_excel.py`**，再用 `git add -f excel_data.db` 推送，无需 LFS。
- **注意**：**不要**把原始大 Excel（如 9.9911.xlsx）提交到 Git，已写在 .gitignore 中；只提交 `excel_data.db` 与 `static/images`。

---

## 二、在 Render 创建 Web Service（你需要做的）

1. **注册/登录**  
   打开 [https://render.com](https://render.com)，用 GitHub 账号登录。

2. **新建 Web Service**  
   - 点击 **New +** → **Web Service**。  
   - 连接你的 GitHub 账号（若未连接），选择**存放本项目的仓库**。  
   - 若项目在子目录，在 **Root Directory** 里填该子目录（例如 `9.0网页计划` 或你实际目录名）。

3. **填写配置（必须与下面一致）**

   | 配置项 | 值 |
   |--------|-----|
   | **Name** | 随意，例如 `excel-query` |
   | **Region** | 选离你近的 |
   | **Runtime** | **Python 3** |
   | **Build Command** | `pip install -r requirements.txt` |
   | **Start Command** | `uvicorn main:app --host 0.0.0.0 --port $PORT` |

   - Render 会自动注入环境变量 **`PORT`**，Start Command 里的 `$PORT` 必须保留，不能改成固定数字。

4. **环境变量（可选）**

   - **EXCEL_PATH**：一般不用在 Render 上设置（线上不做导入时不需要）。  
   - **DB_PATH**：不填则使用项目内的 `excel_data.db`（注意：免费实例重启后数据会清空）。

5. **创建并部署**  
   点击 **Create Web Service**。Render 会执行 Build Command 再执行 Start Command；若日志无报错，部署完成后会给你一个地址，例如：  
   `https://excel-query-xxxx.onrender.com`

---

## 三、部署完成后（你需要做的）

1. **访问站点**  
   浏览器打开 Render 提供的 URL（如上面的 `https://excel-query-xxxx.onrender.com`），应能看到前端的查询页面。

2. **给他人访问**  
   把该 URL 发给他人即可；前端已有 CORS，他人浏览器可直接访问该地址。

3. **若使用“带数据”部署**  
   若你按“方式 A”把 `excel_data.db` 和 `static/images` 提交后部署，首页和搜索会显示你导入的数据；否则一开始是空库，需要数据时再按“方式 B”在本地导入并重新提交部署。

---

## 四、可选：用 Gunicorn 启动（更适合作生产）

若希望用 Gunicorn 多进程，可把 **Start Command** 改为：

```bash
gunicorn main:app -k uvicorn.workers.UvicornWorker -b 0.0.0.0:$PORT
```

项目已包含 `gunicorn` 依赖，无需改代码。

---

## 五、小结：你需要配合的清单

- [ ] 把项目推到 GitHub（或 GitLab），记下仓库与根目录。  
- [ ] 在 Render 创建 Web Service，连到该仓库，填好 **Root Directory**（若在子目录）。  
- [ ] **Build Command** 设为：`pip install -r requirements.txt`。  
- [ ] **Start Command** 设为：`uvicorn main:app --host 0.0.0.0 --port $PORT`（或上面的 gunicorn 命令）。  
- [ ] 不设 **DB_PATH** 即用项目内 SQLite；接受“免费实例重启后数据清空”，或按上文“带数据”方式提交 `excel_data.db` 与 `static/images` 再部署。  
- [ ] 部署完成后用 Render 给的 URL 访问，并把该 URL 发给需要访问的人。

如有报错，在 Render 的 **Logs** 里查看 Build 或 Start 阶段的错误信息，再根据提示排查。
