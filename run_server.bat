@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动 Excel 查询站...
echo 浏览器访问: http://localhost:8000
uvicorn main:app --host 0.0.0.0 --port 8000
