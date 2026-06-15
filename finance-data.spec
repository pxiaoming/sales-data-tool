# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置，请在 Windows 上执行 scripts/build_windows.bat。"""

block_cipher = None

a = Analysis(
    ["scripts/run.py"],
    pathex=[".", "src"],
    binaries=[],
    datas=[
        ("configs/template-example.yml", "configs"),
        ("templates/ka_promo_and_expense_template.xlsx", "templates"),
        ("使用说明.txt", "."),
    ],
    hiddenimports=[
        "finance_excel_flow",
        "finance_excel_flow.config",
        "finance_excel_flow.engine",
        "finance_excel_flow.paths",
        "finance_excel_flow.console",
        "openpyxl",
        "openpyxl.cell._writer",
        "yaml",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="KA促销费用报表",
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
    uac_admin=False,
)
