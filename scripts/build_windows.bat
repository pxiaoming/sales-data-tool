@echo off
chcp 65001 >nul
cd /d "%~dp0\.."

echo ========================================
echo   KA 促销费用报表 — Windows 打包
echo ========================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到 Python，请先安装 Python 3.10 或以上版本。
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo 正在创建虚拟环境...
    python -m venv .venv
    if errorlevel 1 (
        echo 错误：创建虚拟环境失败。
        pause
        exit /b 1
    )
)

call .venv\Scripts\activate.bat

echo 正在安装依赖...
pip install -r requirements.txt -q
pip install -r requirements-build.txt -q
pip install -e . -q
if errorlevel 1 (
    echo 错误：依赖安装失败。
    pause
    exit /b 1
)

echo 正在打包 exe（可能需要几分钟）...
pyinstaller finance-data.spec --noconfirm --clean
if errorlevel 1 (
    echo 错误：打包失败。
    pause
    exit /b 1
)

echo.
echo 打包完成！
echo 可执行文件：dist\KA促销费用报表.exe
echo.
echo 分发时请一并提供：
echo   - dist\KA促销费用报表.exe
echo   - inputs\ 文件夹（可空）
echo   - 使用说明.txt
echo.
pause
