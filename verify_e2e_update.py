# -*- coding: utf-8 -*-
"""真机验收：更新后「自动打开的新版」联网是否正常 —— 本次修复的核心验收点

复现真实更新路径：把 exe 命名成 `<基准名>_更新中.exe`（更新完成后新进程用的正是这个名字）
再启动，检查三件事：

  ① **运行期间不给自己改名** —— 运行中改自己的名字会让该进程出网能力永久失效
     （关掉重开才好），这正是「更新完自动打开的新版点下载就卡住」的根因；
  ② **运行期间联网正常** —— 日志里出现「已从 … 获取版本信息」，证明出网没坏；
  ③ **退出时才改名** —— 且改成 `desired_base_name()` 该有的名字。

两个场景都要过：
  A. 带版本号的绿色版（真实主路径）：`X_v<旧版>_更新中.exe` → 退出后 `X_v<新版>.exe`
  B. 光名字的绿色版：`X_更新中.exe` → 退出后 `X.exe`
     ⚠️ 这不是 bug —— `desired_base_name()` 的判据是「当前文件名是不是 app_name 这个光名字」，
     光名字一律当**安装版**处理（保持固定名），宁可漏改也绝不误伤安装版（改安装版名会
     断掉开始菜单/卸载/下一次更新）。见技能 tkinter-autoupdate。

判定依据 = 程序自己落盘的日志 + 文件系统里的实际文件名。
用**隔离的 APPDATA**，绝不碰用户真实配置与日志。

前置：先跑过 release.py（dist/onefile 下有当前版本的 exe）。
用法：python verify_e2e_update.py
"""
import ctypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent
APP_NAME = "图片表格转Excel助手"
APP_TITLE = "图片表格转 Excel 助手"
WM_CLOSE = 0x0010
WAIT_AFTER_START = 14

u32 = ctypes.windll.user32


def read_ver():
    s = (ROOT / "core.py").read_text(encoding="utf-8")
    m = re.search(r'APP_VERSION\s*=\s*"([\d.]+)"', s)
    if not m:
        raise SystemExit("core.py 里找不到 APP_VERSION")
    return m.group(1)


def enum_windows():
    out = []
    EnumProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

    def cb(hwnd, _lparam):
        n = u32.GetWindowTextLengthW(hwnd)
        if n:
            buf = ctypes.create_unicode_buffer(n + 1)
            u32.GetWindowTextW(hwnd, buf, n + 1)
            if buf.value.strip():
                out.append((hwnd, buf.value))
        return True

    u32.EnumWindows(EnumProc(cb), 0)
    return out


def close_windows(*title_substrings, wait=0.8):
    """按标题片段关窗。⚠️ 不能按 PID 找 —— onefile 的引导器会派生子进程，
    真正建窗口的是子进程（实测踩过：按 PID 找永远关不掉）。"""
    for sub in title_substrings:
        for hwnd, title in enum_windows():
            if sub in title:
                u32.PostMessageW(hwnd, WM_CLOSE, 0, 0)
                print("   已请求关闭窗口：%s" % title)
                time.sleep(wait)


def run_case(exe, start_name, expect_name, label):
    """跑一个场景，返回 (passed, failed) 两个列表。"""
    print()
    print("=" * 62)
    print("场景：%s" % label)
    print("  启动名：%s" % start_name)
    print("  期望：  %s" % expect_name)
    print("=" * 62)

    passed, failed = [], []

    def chk(t, cond, extra=""):
        (passed if cond else failed).append(t)
        print("  [%s] %s %s" % ("PASS" if cond else "FAIL", t, extra if not cond else ""))

    work = Path(tempfile.mkdtemp(prefix="img2excel_e2e_"))
    appdata = work / "appdata"
    appdata.mkdir(parents=True, exist_ok=True)
    staging = work / start_name
    shutil.copy2(exe, staging)

    env = dict(os.environ)
    env["APPDATA"] = str(appdata)          # 隔离配置与日志，不碰真实用户数据
    env.pop("PYTHONPATH", None)

    p = subprocess.Popen([str(staging)], cwd=str(work), env=env)
    print("  已启动，等待联网检查与日志落盘…")
    time.sleep(WAIT_AFTER_START)

    logf = appdata / APP_NAME / "日志" / ("运行日志_%s.log" % time.strftime("%Y%m%d"))
    text = logf.read_text(encoding="utf-8", errors="replace") if logf.exists() else ""

    print("  --- 程序日志 ---")
    for line in (text.strip().splitlines()[-6:] if text.strip() else ["(无日志)"]):
        print("     " + line)

    chk("运行期间没给自己改名", staging.exists())
    chk("运行期间联网正常（日志含「获取版本信息」）",
        "获取版本信息" in text, "日志没出现该行 → 出网可能已失效")
    chk("已登记「退出时改名」", "将在退出时改为" in text)
    chk("进程仍在运行", p.poll() is None)

    print("  关闭程序（触发退出阶段的改名）…")
    close_windows("发现新版本", APP_TITLE)
    time.sleep(3)
    if p.poll() is None:
        close_windows(APP_TITLE)
        time.sleep(4)
    try:
        p.wait(timeout=20)
    except Exception:  # noqa: BLE001
        p.kill()

    text2 = logf.read_text(encoding="utf-8", errors="replace") if logf.exists() else ""
    files = sorted(f.name for f in work.glob("*.exe"))
    print("  --- 退出后目录 ---")
    print("     %s" % files)

    chk("退出后文件名 = %s" % expect_name, (work / expect_name).exists(), "实际=%s" % files)
    chk("中间名文件已被清掉", not staging.exists())
    chk("日志已记录改名结果", "已把程序文件名改为" in text2)

    if p.poll() is None:
        subprocess.run(["taskkill", "/F", "/T", "/PID", str(p.pid)], capture_output=True)

    print("  测试目录：%s" % work)
    return passed, failed


def main():
    ver = read_ver()
    prev = ".".join(str(max(0, int(x) - (1 if i == 2 else 0)))
                    for i, x in enumerate(ver.split(".")))   # 1.2.1 -> 1.2.0
    exe = ROOT / "dist" / "onefile" / (APP_NAME + ".exe")
    if not exe.exists():
        raise SystemExit("找不到 %s —— 请先跑 release.py 打包" % exe)
    print("被测 exe ：%s" % exe)
    print("当前版本 ：v%s（上一版按 v%s 模拟）" % (ver, prev))

    all_pass, all_fail = [], []
    for start, expect, label in (
        ("%s_v%s_更新中.exe" % (APP_NAME, prev),
         "%s_v%s.exe" % (APP_NAME, ver),
         "A. 带版本号的绿色版（真实主路径）"),
        ("%s_更新中.exe" % APP_NAME,
         "%s.exe" % APP_NAME,
         "B. 光名字的绿色版（按安装版处理，保持固定名）"),
    ):
        a, b = run_case(exe, start, expect, label)
        all_pass += a
        all_fail += b

    print()
    print("=" * 62)
    print("验收结果：%d 通过 / %d 失败" % (len(all_pass), len(all_fail)))
    for f in all_fail:
        print("  FAIL: " + f)
    print("=" * 62)
    return 1 if all_fail else 0


if __name__ == "__main__":
    sys.exit(main())
