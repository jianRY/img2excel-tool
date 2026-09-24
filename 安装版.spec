# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置 —— 图片表格转 Excel 助手【目录版 / 安装包负载】

本项目固定产出两个版本：
  1. 单文件版  —— 图片表格转Excel助手.spec（--onefile），绿色版，双击即用
  2. 目录版    —— 本文件（--onedir），仅作为 Inno Setup 安装包的负载

为什么安装版用目录版而不是单文件版：
  单文件版每次启动都要把 ~120MB 解包到临时目录，装到硬盘上毫无意义。
  目录版一次解包到位集合目录，启动快得多 —— 这才是安装版该有的形态。
"""
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = [("models/slanet-plus.onnx", "models")]
datas += [("assets/app.ico", "assets")]
hiddenimports = []
datas += collect_data_files("rapidocr_onnxruntime")
datas += collect_data_files("rapid_table")
hiddenimports += collect_submodules("rapid_table")
hiddenimports += collect_submodules("rapidocr_onnxruntime")

excludes = [
    "torch", "torchvision", "tensorflow", "matplotlib", "scipy",
    "pandas", "IPython", "notebook", "pytest", "setuptools",
]

a = Analysis(
    ["gui.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure)

# onedir 模式：EXE 与 COLLECT 配合，产物是「一个目录 + 目录内的 exe」
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="图片表格转Excel助手",
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
    icon="assets/app.ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="图片表格转Excel助手",
)
