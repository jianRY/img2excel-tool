# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller 打包配置 —— 图片表格转 Excel 助手

包含：
  - 内置 SLANet-Plus 表格结构模型
  - rapidocr / rapid_table 的数据文件与子模块
  - 应用图标（exe 图标 + 运行期窗口图标）
"""
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = [("models/slanet-plus.onnx", "models")]
datas += [("assets/app.ico", "assets")]
hiddenimports = []
datas += collect_data_files("rapidocr_onnxruntime")
datas += collect_data_files("rapid_table")
hiddenimports += collect_submodules("rapid_table")
hiddenimports += collect_submodules("rapidocr_onnxruntime")

# 排除用不到的重量级依赖，控制体积
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

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="图片表格转Excel助手",
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
    icon="assets/app.ico",
)
