@echo off
echo 正在安装PyInstaller...
pip install pyinstaller
echo.
echo 正在打包...
pyinstaller --noconfirm --onefile --windowed --name "完税证明提取工具" --add-data "app.py;." app.py
echo.
echo 打包完成！可执行文件在 dist 目录下
pause
