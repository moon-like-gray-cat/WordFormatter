@echo off
REM ========================================
REM Build WordFormatter EXE using PyInstaller
REM ========================================

REM 进入项目根目录（可选，根据你的路径调整）
cd /d %~dp0

REM 调用 PyInstaller 打包
pyinstaller ^
    --onefile ^
    --windowed ^
    --icon=src\wordtool\resources\icon.ico ^
    --add-data "src\wordtool\resources;resources" ^
    run.py

REM 打包完成提示
echo.
echo ============================
echo Build finished!
echo EXE is located in the "dist" folder.
echo ============================
pause
