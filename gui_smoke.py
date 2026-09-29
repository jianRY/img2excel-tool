# -*- coding: utf-8 -*-
"""GUI 冒烟测试：构建窗口、检查控件、模拟一次转换流程（不进入 mainloop）"""
import sys, os, threading, time
sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import gui

app = gui.App()
app.update()
print("窗口标题:", app.title())
print("尺寸:", app.geometry())

# 控件存在性
for attr in ("tree", "progress", "log_text", "btn_go", "btn_open", "lbl_status", "btn_upd"):
    print(f"  {attr:12s}", "OK" if hasattr(app, attr) else "缺失")
print("keep_raw 默认:", app.keep_raw.get())
print("per_file 默认:", app.per_file.get())
print("自动更新：程序目录可写 =", gui.can_self_update())
print("自动更新：APP_NAME 与 exe 名一致 =", gui.APP_NAME == "图片表格转Excel助手")
print("自动更新：配置文件路径 =", gui.config_path())

# 测 _add_paths（用真实样本；目录来自本机配置，仓库里不写死路径）
import glob
import devconfig

sample = devconfig.sample_dir()
_exts = (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp")
fs = sorted(f for f in glob.glob(os.path.join(sample, "*"))
            if os.path.splitext(f)[1].lower() in _exts)[:2]
print("样本目录:", sample, "| 取到", len(fs), "个")
app._add_paths(fs)
app.update()
print("加入文件数:", len(app.files), "| 列表行数:", len(app.tree.get_children()))
print("默认输出路径:", os.path.basename(app.out_path.get()))

# 测输出路径校验逻辑
app.out_path.set(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "gui_smoke.xlsx"))
print("输出后缀检查:", app.out_path.get().lower().endswith(".xlsx"))

# 测移除/清空
app.remove_selected()
app.update()
print("移除后文件数:", len(app.files))
app.clear_files()
app.update()
print("清空后文件数:", len(app.files))

app._log("冒烟测试日志正常", "ok")
app.destroy()
print("\nGUI 冒烟测试通过")
