# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules, collect_data_files
pkgs = ['pandas','numpy','xlrd','xlwt','xlutils','xlsxwriter','openpyxl','tkinter']
hiddenimports = []
for p in pkgs:
    hiddenimports += collect_submodules(p)
datas = []
for p in pkgs:
    datas += collect_data_files(p)
block_cipher = None
a = Analysis(
    ['invup.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    noarchive=False
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='Inventory Updater',
    console=False,
    icon=None
)
