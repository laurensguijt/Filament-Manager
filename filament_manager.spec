# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
import sys
import os

# Get the customtkinter, tkcalendar, etc. files
datas = []
binaries = []
hiddenimports = []

# Add the customtkinter and tkcalendar packages
packages = ['customtkinter', 'tkcalendar', 'openpyxl', 'PIL', 'reportlab', 'qrcode', 'barcode']
for package in packages:
    tmp_datas, tmp_binaries, tmp_hiddenimports = collect_all(package)
    datas += tmp_datas
    binaries += tmp_binaries
    hiddenimports += tmp_hiddenimports

# We laten de Excel-database niet meenemen in de executable
# Deze wordt automatisch aangemaakt op eerste gebruik door init_excel() in data_operations.py

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Include the entire Filament_Manager directory
a.datas += Tree('Filament_Manager', prefix='Filament_Manager')

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# Executable metadata om virusdetecties te verminderen
exe_options = {
    'name': 'Filament Manager',
    'debug': False,
    'bootloader_ignore_signals': False,
    'strip': False,
    'upx': False,  # UPX compressie uitschakelen om virusdetecties te verminderen
    'console': False,
    'disable_windowed_traceback': False,
    'argv_emulation': False,
    'target_arch': None,
    'codesign_identity': None,
    'entitlements_file': None,
    'version': '1.0.0',
    'file_description': 'Filament Manager - 3D Printing Inventory Management',
    'product_name': 'Filament Manager',
    'legal_copyright': '© 2025 laurensguijt',
    'company_name': 'laurensguijt',
    'uac_admin': False,
}

# Gebruik icon als het bestaat
if os.path.exists('icon.ico'):
    exe_options['icon'] = 'icon.ico'

# Single-file executable configuratie
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=exe_options['name'],
    debug=exe_options['debug'],
    bootloader_ignore_signals=exe_options['bootloader_ignore_signals'],
    strip=exe_options['strip'],
    upx=exe_options['upx'],
    runtime_tmpdir=None,  # Tijdelijke bestanden in-memory houden
    console=exe_options['console'],
    disable_windowed_traceback=exe_options['disable_windowed_traceback'],
    argv_emulation=exe_options['argv_emulation'],
    target_arch=exe_options['target_arch'],
    codesign_identity=exe_options['codesign_identity'],
    entitlements_file=exe_options['entitlements_file'],
    icon=exe_options.get('icon', None),
    version='file_version_info.txt'  # Versie-informatie uit apart bestand
) 