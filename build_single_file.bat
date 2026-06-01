@echo off
echo ========================================
echo   Excel 筛选工具 瘦身打包脚本
echo ========================================
echo.

echo 1. 激活虚拟环境...
call .venv\Scripts\activate

echo.
echo 2. 清理旧打包文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del *.spec

echo.
echo 3. 开始精简打包...
pyinstaller -F -w --strip --upx-dir . ^
    --exclude-module=tkinter ^
    --exclude-module=test ^
    --exclude-module=unittest ^
    --exclude-module=lib2to3 ^
    --add-data "assets/*;assets/" ^
    --icon=assets/icons/ExcelTool.ico ^
    --name=ExcelTool ^
    run.py

echo.
echo 4. 打包完成！
echo 输出：dist\ExcelTool.exe
echo.
pause