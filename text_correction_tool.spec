# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=['y:\\02_程式\\10_program\\win11_windsurf_project\\editertool_version07'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.simpledialog',
        'docx2txt',
        'msoffcrypto',
        'opencc',
        'docx',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'checknumber_word',
        'paragraph_formatter',
        'typo_corrector',
        'tkdnd_wrapper',
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

# Add data files (configuration, resources, etc.)
a.datas += [('protected_words.json', 'y:\\02_程式\\10_program\\win11_windsurf_project\\editertool_version07\\protected_words.json', 'DATA')]
a.datas += [('settings.json', 'y:\\02_程式\\10_program\\win11_windsurf_project\\editertool_version07\\settings.json', 'DATA')]

# Add Python modules
a.datas += [('checknumber_word.py', 'y:\\02_程式\\10_program\\win11_windsurf_project\\editertool_version07\\checknumber_word.py', 'DATA')]
a.datas += [('paragraph_formatter.py', 'y:\\02_程式\\10_program\\win11_windsurf_project\\editertool_version07\\paragraph_formatter.py', 'DATA')]
a.datas += [('typo_corrector.py', 'y:\\02_程式\\10_program\\win11_windsurf_project\\editertool_version07\\typo_corrector.py', 'DATA')]
a.datas += [('tkdnd_wrapper.py', 'y:\\02_程式\\10_program\\win11_windsurf_project\\editertool_version07\\tkdnd_wrapper.py', 'DATA')]

# Create logs directory in the executable
a.datas += [('logs/.placeholder', 'y:\\02_程式\\10_program\\win11_windsurf_project\\editertool_version07\\logs\\.placeholder', 'DATA')]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='編審神器',
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
    icon=None,  # Replace with path to your icon file if available
)
