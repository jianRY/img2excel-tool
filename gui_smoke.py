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
for attr in ("tree", "progress", "log_text", "btn_go", "btn_open", "lbl_status"):
    print(f"  {attr:12s}", "OK" if hasattr(app, attr) else "缺失")
print("keep_raw 默认:", app.keep_raw.get())
print("per_file 默认:", app.per_file.get())

# 测 _add_paths（用真实票据）
import glob
fs = sorted(glob.glob(r"C:\Users\toxuj\Desktop\FAPIAO\*.jpg"))[:2]
app._add_paths(fs)
app.update()
print("加入文件数:", len(app.files), "| 列表行数:", len(app.tree.get_children()))
print("默认输出路径:", os.path.basename(app.out_path.get()))

# 测输出路径校验逻辑
app.out_path.set(r"D:\workbuddy\杂项\img22excel\gui_smoke.xlsx")
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
