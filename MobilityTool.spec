# -*- mode: python ; coding: utf-8 -*-
#
# PyInstaller spec file for building the MobilityTool executable.
#
# Why this file is needed:
# - PyInstaller uses a spec file to know which script(s), data files, and
#   hidden/dynamically imported modules to include when building a standalone
#   executable. Without this spec the default build may miss required files or
#   imports (especially modules imported dynamically by libraries like
#   pandas/openpyxl).
# - This spec explicitly bundles the helper file `ugent_mobility_backend_fixed.py`
#   into the distribution (datas), and lists `openpyxl` and `pandas` as
#   hiddenimports so PyInstaller includes them even if they are imported
#   dynamically at runtime.
# - It also configures the build options (e.g., `console=False` to create a
#   windowed app, `upx=True` to enable binary compression). Keeping these
#   settings in version control ensures reproducible builds for other
#   developers and CI.


a = Analysis(
    ['MobilityApp2026.py'],
    pathex=[],
    binaries=[],
    datas=[('ugent_mobility_backend_fixed.py', '.')],
    hiddenimports=['openpyxl', 'pandas'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='MobilityTool',
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
