@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
python --version >nul 2>&1
if errorlevel 1 (
    echo 未找到 Python，请先安装 Python 3.10 或更高版本，并勾选 Add python.exe to PATH。
    pause
    exit /b 1
)
python -m pip install --upgrade pip
if errorlevel 1 goto failed
python -m pip install -r requirements.txt
if errorlevel 1 goto failed
python -m PyInstaller --clean --noconfirm "合同生成工具.spec"
if errorlevel 1 goto failed
echo.
echo 打包完成：dist\合同生成工具.exe
echo 使用时请把真实 公司数据.xlsx 放到 exe 同目录。
echo 运行示例：dist\合同生成工具.exe 产品明细表.xls
pause
exit /b 0
:failed
echo.
echo 打包失败，请检查上方错误信息。
pause
exit /b 1
