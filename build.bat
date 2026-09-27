@echo off
chcp 65001 >nul
REM 一键打包 ResumeAssistant.exe（在独立 venv 中构建，规避 Anaconda 的 pathlib 冲突）
cd /d "%~dp0"

if not exist .venv\Scripts\python.exe (
    echo 创建虚拟环境...
    python -m venv .venv
)

echo 安装依赖...
.venv\Scripts\python.exe -m pip install --quiet --upgrade pip
.venv\Scripts\python.exe -m pip install --quiet python-docx pyinstaller

echo 打包中...
.venv\Scripts\python.exe -m PyInstaller --noconsole --onefile --name ResumeAssistant main.py

echo.
echo 完成：dist\ResumeAssistant.exe
pause
