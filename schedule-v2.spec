# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for schedule-v2
"""
block_cipher = None

a = Analysis(
    ['call_ai.py', 'json_to_xlsx.py'],
    path=[],
    binaries=[],
    datas=[
        ('data', 'data'),
        ('config.ini', '.'),
        ('prompt_part1_fixed_rules.md', '.'),
        ('prompt_part2_variables.md', '.'),
        ('prompt_part3_output_rules.md', '.'),
        ('prompt_weekend_time_rules.md', '.'),
        ('prompt_weekday_weekend_separation.md', '.'),
        ('prompt_work_order_uniqueness.md', '.'),
    ],
    hiddenimports=[
        'json_to_xlsx',
        'openpyxl',
        'pandas',
        'openai',
        'statistics',
        'openpyxl.styles',
        'openpyxl.utils',
        'openpyxl.styles.fonts',
        'openpyxl.styles.alignment',
        'openpyxl.styles.border',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
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
    name='schedule-v2',
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
