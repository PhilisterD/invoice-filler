"""
Windows build script
Run on Windows: python build_windows.py
"""
import PyInstaller.__main__
import os

base_dir = os.path.dirname(os.path.abspath(__file__))

hidden_imports = [
    '--hidden-import=numpy.core._dtype_ctypes',
    '--hidden-import=numpy.core._multiarray_umath',
    '--hidden-import=numpy.core.multiarray',
    '--hidden-import=numpy.random.common',
    '--hidden-import=numpy.random.bounded_integers',
    '--hidden-import=numpy.random.entropy',
    '--hidden-import=pandas._libs.tslibs.timedeltas',
    '--hidden-import=pandas._libs.tslibs.nattype',
    '--hidden-import=pandas._libs.tslibs.np_datetime',
    '--hidden-import=pandas._libs.tslibs.base',
    '--hidden-import=pandas._libs.tslibs.timezones',
    '--hidden-import=pandas._libs.tslibs.conversion',
    '--hidden-import=pandas._libs.tslib',
    '--hidden-import=pandas._libs.hashtable',
    '--hidden-import=pandas._libs.interval',
    '--hidden-import=pandas._libs.missing',
    '--hidden-import=pandas._libs.tslibs.offsets',
    '--hidden-import=pandas._libs.tslibs.parsing',
    '--hidden-import=pandas._libs.tslibs.period',
    '--hidden-import=pandas._libs.tslibs.timestamps',
    '--hidden-import=pandas._libs.tslibs.ccalendar',
    '--hidden-import=pandas._libs.tslibs.tzconversion',
    '--hidden-import=pandas._libs.testing',
    '--hidden-import=pandas._libs.hashing',
    '--hidden-import=pandas._libs.ops',
    '--hidden-import=pandas._libs.parsers',
    '--hidden-import=pandas._libs.reduction',
    '--hidden-import=pandas._libs.index',
    '--hidden-import=pandas._libs.indexing',
    '--hidden-import=pandas._libs.internals',
    '--hidden-import=pandas._libs.join',
    '--hidden-import=pandas._libs.lib',
    '--hidden-import=pandas._libs.properties',
    '--hidden-import=pandas._libs.reshape',
    '--hidden-import=pandas._libs.skiplist',
    '--hidden-import=pandas._libs.sparse',
    '--hidden-import=pandas._libs.writers',
    '--hidden-import=openpyxl',
    '--hidden-import=openpyxl.cell._writer',
    '--hidden-import=pkg_resources',
    '--collect-submodules=docx',
    '--collect-submodules=docx.oxml',
]

add_data = []
if os.path.exists(os.path.join(base_dir, 'templates')):
    add_data.append('--add-data=templates;templates')
if os.path.exists(os.path.join(base_dir, 'static')):
    add_data.append('--add-data=static;static')

args = [
    'app.py',
    '--name=invoice-filler',
    '--onefile',
    '--windowed',
    '--clean',
    '--noconfirm',
] + add_data + hidden_imports

PyInstaller.__main__.run(args)
