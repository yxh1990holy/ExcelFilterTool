# Excel 数据处理工具

一个基于 PySide6 的 Excel 数据处理工具，支持多工作表查看、数据筛选和导出功能。

## ✨ 功能特性

- 📊 支持多工作表同时显示
- 🔍 灵活的数据筛选功能（支持 =, !=, >, <, contains 等操作）
- 💾 导出当前工作表或所有工作表
- 🚀 懒加载机制，大文件快速启动
- 🎨 现代化的界面设计


# 打包命令

pyinstaller -D -w --clean --name="ExcelTool" run.py

--clean：清理临时文件
--noconfirm：覆盖输出目录时不询问
-F,--onefile:打包成单个可执行文件（启动较慢，因为需要解压到临时目录）
-D,--onedir:打包成一个文件夹（默认模式，启动快，推荐使用）
--exclude-module 排除未使用的库（如pyinstaller --exclude-module=tkinter script.py）
-w, --windowed, --noconsole:不显示控制台窗口（GUI程序使用）
-c, --console, --nowindowed:显示控制台窗口（默认，命令行程序使用）
-i, --icon:指定程序图标，如 -i icon.ico
