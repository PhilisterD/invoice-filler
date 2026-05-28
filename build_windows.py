"""
Windows 打包脚本
在 Windows 上运行：python build_windows.py
"""
import PyInstaller.__main__
import os
import sys

base_dir = os.path.dirname(os.path.abspath(__file__))

PyInstaller.__main__.run([
    'app.py',
    '--name=早鸟天筹明细生成器',
    '--onefile',
    '--windowed',
    '--add-data=templates;templates',
    '--icon=NONE',
    '--clean',
    '--noconfirm',
])

print("Build complete: dist/invoice-filler.exe")
