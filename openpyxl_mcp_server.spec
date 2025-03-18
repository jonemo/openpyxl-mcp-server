# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['openpyxl_mcp_server.py'],
    hiddenimports=['openpyxl', 'mcp'],
    # Windows-specific settings (ignored on non-Windows platforms)
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name='openpyxl_mcp_server',
    console=True,
    argv_emulation=True,  # Ensures proper command-line argument handling on all platforms
    upx=True,  # Enable UPX compression for smaller executables
) 