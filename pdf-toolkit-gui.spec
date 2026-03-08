# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

datas = []
binaries = []
hiddenimports = []

vendor_dir = Path("vendor")
if vendor_dir.exists():
    for path in vendor_dir.rglob("*"):
        if path.is_file():
            relative_parent = path.parent.relative_to(vendor_dir)
            target_dir = str(Path("vendor") / relative_parent)
            if path.suffix.lower() in {".exe", ".dll", ".pyd", ".so", ".dylib"}:
                binaries.append((str(path), target_dir))
            else:
                datas.append((str(path), target_dir))


a = Analysis(
    ['src\\pdf_toolkit\\gui.py'],
    pathex=['src'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'aiohttp',
        'boto3',
        'botocore',
        'cv2',
        'faiss',
        'fastapi',
        'frontend',
        'gradio',
        'huggingface_hub',
        'jinja2',
        'matplotlib',
        'nltk',
        'numba',
        'numpy.testing',
        'pandas',
        'scipy',
        'sklearn',
        'starlette',
        'tensorflow',
        'torch',
        'torchaudio',
        'torchvision',
        'transformers',
        'uvicorn',
    ],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='pdf-toolkit-gui',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='pdf-toolkit-gui',
)
