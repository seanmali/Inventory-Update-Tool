# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

pkgs = ['pandas','numpy','xlrd','xlwt','xlutils','xlsxwriter','openpyxl']
datas, binaries, hiddenimports = [], [], []
for p in pkgs:
    d, b, h = collect_all(p)
    datas += d
    binaries += b
    hiddenimports += h

a = Analysis(
    ['invup.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    noarchive=False
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)
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
