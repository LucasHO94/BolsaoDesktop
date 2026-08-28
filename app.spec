# -*- mode: python ; coding: utf-8 -*-
import os
from PyInstaller.utils.hooks import collect_data_files

spec_dir = os.path.abspath(".")
block_cipher = None

datas_list = [
    (os.path.join(spec_dir, 'carta.html'), '.'),
    (os.path.join(spec_dir, 'style.css'), '.'),
    (os.path.join(spec_dir, 'images'), 'images'),
    (os.path.join(spec_dir, 'dist', 'updater.exe'), '.')
]
datas_list += collect_data_files('ttkbootstrap')

# CORREÇÃO: Adicionados hiddenimports para o Google Sheets funcionar
hidden_imports_list = [
    'gspread',
    'google.auth',
    'google.oauth2',
    'google.auth.transport.requests',
    'requests',
    'json',
    'ttkbootstrap'
]

a = Analysis(
    [os.path.join(spec_dir, 'app.py')],
    pathex=[spec_dir],
    binaries=[('C:\\Program Files\\GTK3-Runtime Win64\\bin\\*.dll', '.')],
    datas=datas_list,
    hiddenimports=hidden_imports_list,  # <-- AQUI ESTÁ A CORREÇÃO
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'torch', 'tensorflow', 'keras', 'sklearn', 'scipy', 'matplotlib', 
        'cv2', 'torchaudio', 'torchvision', 'numpy.core._multiarray_tests'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='GestorBolsao',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False, # Mantenha False para não abrir janela preta
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(spec_dir, 'images', 'matriz.ico')
)