@echo off
echo ========================================
echo   Excel 筛选工具打包脚本
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
echo 3. 开始打包...
pyinstaller -D -w --name="ExcelTool" ^
    --hidden-import=openpyxl ^
    --hidden-import=pandas ^
    --hidden-import=PySide6.QtXml ^
    --hidden-import=PySide6.QtCore ^
    --hidden-import=PySide6.QtGui ^
    --hidden-import=PySide6.QtWidgets ^
    run.py

echo.
echo 4. 打包完成！
echo 输出文件位置: dist\ExcelFilterTool.exe
echo.
pause