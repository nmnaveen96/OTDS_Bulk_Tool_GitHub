# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec – OTDS Bulk Provisioning Tool
# Build: pyinstaller otds_bulk_tool.spec --clean --noconfirm

block_cipher = None

a = Analysis(
    ['otds_bulk_tool.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'requests','requests.adapters','requests.auth','requests.models',
        'requests.sessions','requests.utils',
        'urllib3','urllib3.util','urllib3.util.retry','urllib3.poolmanager',
        'charset_normalizer','certifi','idna',
        'pandas','pandas._libs','pandas._libs.tslibs','pandas.core',
        'pandas.io','pandas.io.excel','pandas.io.parsers',
        'openpyxl','openpyxl.styles','openpyxl.styles.fills',
        'openpyxl.styles.fonts','openpyxl.styles.alignment',
        'openpyxl.styles.borders','openpyxl.utils','openpyxl.utils.cell',
        'openpyxl.reader','openpyxl.writer','openpyxl.workbook','openpyxl.worksheet',
        'et_xmlfile','pkg_resources','pkg_resources.extern',
        'dateutil','dateutil.parser','pytz','six',
    ],
    excludes=[
        'tkinter','matplotlib','scipy','numpy','PIL',
        'PyQt5','PyQt6','PySide2','PySide6','wx','gi','gtk',
        'botocore','boto3','sqlalchemy','django','flask','tornado','twisted',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name='otds_bulk_tool',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='assets/icon.ico',
)
