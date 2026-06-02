"""
Windows build script
Run on Windows: python build_windows.py
"""
import PyInstaller.__main__
import os

base_dir = os.path.dirname(os.path.abspath(__file__))

# Use --collect-all to bundle entire packages including C-extensions
# This is more reliable than --hidden-import for numpy/pandas
hidden_imports = [
    '--collect-all=numpy',
    '--collect-all=pandas',
    '--collect-all=openpyxl',
    '--collect-all=docx',
    '--hidden-import=pkg_resources',
]

add_data = []
if os.path.exists(os.path.join(base_dir, 'templates')):
    add_data.append('--add-data=templates;templates')
if os.path.exists(os.path.join(base_dir, 'static')):
    add_data.append('--add-data=static;static')
if os.path.exists(os.path.join(base_dir, 'template.docx')):
    add_data.append('--add-data=template.docx;.')
if os.path.exists(os.path.join(base_dir, 'uploads')):
    add_data.append('--add-data=uploads;uploads')


args = [
    'app.py',
    '--name=invoice-filler',
    '--onefile',
    '--windowed',
    '--clean',
    '--noconfirm',
] + add_data + hidden_imports

PyInstaller.__main__.run(args)
