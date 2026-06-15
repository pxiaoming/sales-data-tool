@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" scripts\run.py
) else (
    echo 错误：未找到 Python 环境，请联系维护人员安装依赖。
    echo 维护人员请在项目目录执行：pip install -r requirements.txt
)

echo.
pause
