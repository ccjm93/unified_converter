# -*- mode: python ; coding: utf-8 -*-
"""통합 파일형식 변환기 PyInstaller 스펙.

빌드:  pyinstaller build/unified_converter.spec   (unified_converter 폴더에서 실행)

배포 안정성 주의:
  - 한컴/Excel COM 접근은 win32com.client.Dispatch(late-binding)만 사용한다.
    gencache.EnsureDispatch를 쓰지 않으므로 gen_py 캐시를 번들할 필요가 없다.
  - 만약 향후 early-binding이 필요해지면, sys.frozen 시 win32com.__gen_path__를
    쓰기 가능한 임시 경로로 리디렉션하는 런타임 훅을 추가해야 한다.
"""
import os

from PyInstaller.utils.hooks import collect_data_files

block_cipher = None
ROOT = os.path.abspath(os.path.join(os.getcwd()))

# tkinterdnd2의 네이티브 tkdnd 라이브러리(.dll/.tcl)를 번들해야 frozen exe에서 DnD가 동작한다.
# 미설치여도 빌드가 깨지지 않도록 실패는 무시(런타임은 graceful degradation).
try:
    _dnd_datas = collect_data_files('tkinterdnd2')
except Exception:
    _dnd_datas = []

a = Analysis(
    [os.path.join(ROOT, 'main.py')],   # spec이 build/ 하위에 있으므로 ROOT 기준 절대경로로 지정
    pathex=[ROOT],
    binaries=[],
    datas=_dnd_datas,
    hiddenimports=[
        'tkinterdnd2',
        'win32com',
        'win32com.client',
        'win32com.shell',
        'pythoncom',
        'pywintypes',
        'win32api',
        'win32process',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pyhwpx'],   # 의도적으로 미사용
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
    name='통합변환기',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
