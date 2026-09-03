import os

conda_lib_bin = r'D:\miniconda3\Library\bin'
required_binaries = []
for dll_name in ['libexpat.dll', 'libcrypto-3-x64.dll', 'libssl-3-x64.dll', 'liblzma.dll', 'LIBBZ2.dll']:
    dll_path = os.path.join(conda_lib_bin, dll_name)
    if os.path.exists(dll_path):
        required_binaries.append((dll_path, '.'))

a = Analysis(
    ['server.py'],
    pathex=[],
    binaries=required_binaries,
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'numpy',
        'pandas',
        'scipy',
        'matplotlib',
        'PIL',
        'Pillow',
        'tkinter',
        '_tkinter',
        'tcl',
        'tk',
        'sqlite3',
        'unittest',
        'pytest',
        'IPython',
        'jupyter',
        'tornado',
        'zmq',
        'pygments',
        'scipy',
        'statsmodels',
        'numba',
        'llvmlite',
        'botocore',
        'boto3',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='server',
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
)
