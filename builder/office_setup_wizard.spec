# -*- mode: python ; coding: utf-8 -*-

import os

builder = os.path.abspath(os.getcwd())
root = os.path.dirname(builder)

a = Analysis(
    [os.path.join(builder, 'office_setup_wizard.py')],
    pathex=[],
    binaries=[],
    datas=[
        (os.path.join(builder, 'office_setup_wizard.ico'), '.'),
        (os.path.join(builder, 'icons'), 'icons'),
        (os.path.join(root, 'resources', 'setup.exe'), '.'),
        (os.path.join(root, 'resources', 'officedeploymenttool_20131-20090.exe'), '.'),
        (os.path.join(root, 'resources', 'configuration-Office365-x64.xml'), '.'),
        (os.path.join(root, 'resources', 'Configuration.xml'), 'Configuration'),
    ],
    hiddenimports=[],
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
    name='OfficeSetupWizard',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=os.path.join(builder, 'version_info.txt'),
    uac_admin=True,
    icon=[os.path.join(builder, 'office_setup_wizard.ico')],
)
