a = Analysis(
    ['run.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config.yaml', '.'),
        ('templates/采购合同模板.docx', 'templates'),
        ('templates/销售合同模板.docx', 'templates'),
        ('templates/销售订单溢价版.docx', 'templates'),
    ],
    # tkinter/filedialog 在函数内导入，PyInstaller 静态分析发现不了，
    # 必须显式收集，否则打包后"2 = 自定义文件夹"的弹窗不可用
    hiddenimports=['tkinter', 'tkinter.filedialog'],
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
    name='合同生成工具',
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
