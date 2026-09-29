#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""图片表格转 Excel 助手 —— 一键发版脚本

固定产出「两个版本」，与团队其他项目约定一致：
  1. 单文件版（绿色版）  dist/onefile/图片表格转Excel助手.exe        —— 双击即用
  2. 安装版（带卸载）    dist/installer/Img2ExcelTool_vX.Y.Z_setup.exe

为什么安装版用 onedir 而不是 onefile：
  单文件版每次启动都要把 ~120MB 解包到临时目录，装到硬盘上毫无意义。
  安装包负载用目录版（一次解包到位集合目录），启动快得多。

用法：
  python release.py                    # 完整发版：打包 -> 自检 -> 签名 -> 交付 -> 发布 GitHub
  python release.py --version 1.3.0    # 指定版本号（会写回 core.py）
  python release.py --bump patch       # 1.2.0 -> 1.2.1
  python release.py --bump minor       # 1.2.0 -> 1.3.0
  python release.py --bump major       # 1.2.0 -> 2.0.0
  python release.py --no-installer     # 只出单文件版
  python release.py --skip-sign        # 跳过签名（仅调试）
  python release.py --skip-selftest    # 跳过打包产物自检
  python release.py --no-publish       # 只做本地打包/签名/交付，不推 GitHub

    set RELEASE_PROXY=http://127.0.0.1:7890   指定代理
    set RELEASE_PROXY=                        强制直连
"""
import argparse
import glob
import hashlib
import json
import os
import re
import shutil
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request

import devconfig   # 同目录模块：本机配置读取（环境变量 -> local_config.json -> 默认值）

ROOT = os.path.dirname(os.path.abspath(__file__))
CORE_PY = os.path.join(ROOT, "core.py")
DIST = os.path.join(ROOT, "dist")
CACHE = os.path.join(ROOT, ".pybuild_cache")

# ---------------- 本机专属配置（不入库） ----------------
# 仓库里**只放相对路径与通用默认值**：凡与开发机绑定的绝对路径
# （构建用 venv、签名脚本、Inno Setup、测试样本目录…）一律走 devconfig 的取值链：
#     ① 环境变量                          —— 临时覆盖，优先级最高
#     ② .pybuild_cache/local_config.json   —— 本机长期配置，已被 .gitignore 排除
#     ③ 下面的通用默认值 / 相对路径        —— 换台机器也能直接跑
# 目的：公开仓库里搜不到任何开发机信息（用户名、目录结构、其他项目位置）。
# 各配置项含义见 local_config.example.json。
_cfg = devconfig.get


# 交付目录（两个 exe + 使用说明一起放）
# 默认 = 源码目录的兄弟目录（本项目的交付位置本来就是这样），本机可用配置覆盖。
DELIVERY = _cfg("IMG2EXCEL_DELIVERY", "delivery_dir",
                os.path.normpath(os.path.join(ROOT, os.pardir, "图片表格转Excel助手")))

APP_NAME = "图片表格转Excel助手"          # 用于 exe 文件名（保持 ASCII 友好）
APP_TITLE = "图片表格转 Excel 助手"        # 用于显示
EXE_NAME = "图片表格转Excel助手.exe"
ONEFILE_DIR = os.path.join(DIST, "onefile")
ONEDIR_DIR = os.path.join(DIST, "onedir")
INSTALLER_DIR = os.path.join(DIST, "installer")

# ---------------- GitHub ----------------
OWNER_REPO = "jianRY/img2excel-tool"
REPO_URL = "https://github.com/" + OWNER_REPO
MAIN_BRANCH = "main"
# Release 资产名保持 ASCII（中文名会被 GitHub 回退成 default.exe）
ASSET_ONEFILE = "Img2ExcelTool_v%s_onefile.exe"
ASSET_SETUP = "Img2ExcelTool_v%s_setup.exe"
ASSET_META = "update.json"          # ⚠️ 名字不能改：服务器脚本按这个名字去
                                    #    releases/latest/download/update.json 抄到 /updates/<app>.json

# ---------------- 自有下载站（兜底源） ----------------
# 客户端源顺序：加速镜像 → GitHub 原站 → 自有站 → GitHub API。
# 短名 APP_KEY 必须三处一致：这里 / autoupdate.py 的 _REPO_TO_APP / 服务器 REPOS。
# ⚠️ 这个地址客户端 exe 里必然明文存在（要连它才能兜底），改它等于让所有旧版失去兜底源；
#    要覆盖请用环境变量 IMG2EXCEL_SITE_URL 或本机配置，**不要直接改这里的默认值**。
SITE_URL = _cfg("IMG2EXCEL_SITE_URL", "site_url", "http://download.internal:8888")
APP_KEY = "img2excel"

# 构建/签名用的解释器：默认就用「当前正在跑本脚本的解释器」——发版脚本本来就
# 是拿那个带 tkinter 的 venv 运行的，天生正确，不必写死 venv 路径。
PY = _cfg("IMG2EXCEL_PY", "python", sys.executable)
PYINSTALLER = _cfg("IMG2EXCEL_PYINSTALLER", "pyinstaller",
                   os.path.join(os.path.dirname(PY),
                                "pyinstaller.exe" if os.name == "nt" else "pyinstaller"))
# 签名脚本：默认在项目自己的 .pybuild_cache/signing/ 下找；本机若用外部共享的
# 签名目录，在 local_config.json 里写 "sign_py" 指定即可。
SIGN_PY = _cfg("IMG2EXCEL_SIGN_PY", "sign_py",
               os.path.join(CACHE, "signing", "sign.py"))


def _pick_git():
    """挑一个真能用的 git。

    坑（2026-09-16 实测）：PortableGit 的 git.exe 把 git-remote-https.exe 放在
    mingw64\\bin（不在 exec-path 里），git 靠翻 PATH 找远程 helper。本机 shell 的
    PATH 被裁剪过，PortableGit 的 bin 不一定在里面 —— 于是 push 报
    `git: 'remote-https' is not a git command`；就算把 PATH 补上也只是变成
    **静默失败**（returncode 128、零输出），更难看懂。
    系统 Git（2.55）helper 布局正常，实测可直接推送，所以优先用它。
    """
    for c in (r"C:\Program Files\Git\cmd\git.exe",
              r"C:\Program Files (x86)\Git\cmd\git.exe"):
        if os.path.exists(c):
            return c
    return None


GIT = _pick_git()


def _pick_wincred():
    """从系统 Git 的安装位置推导 wincred helper —— 不写死任何本机路径。

    系统 Git 布局固定：<git_root>\\cmd\\git.exe
                 与 <git_root>\\mingw64\\libexec\\git-core\\git-credential-wincred.exe
    （**不要**用 PortableGit：它把 helper 放在 mingw64\\bin，不在 exec-path 里，
     git 靠翻 PATH 找不到，会静默失败。）
    """
    if not GIT:
        return ""
    cand = os.path.join(os.path.dirname(os.path.dirname(GIT)),
                        "mingw64", "libexec", "git-core",
                        "git-credential-wincred.exe")
    return cand if os.path.exists(cand) else ""


# 取 PAT 的 helper（wincred）：默认从系统 Git 位置推导
WCRED = _cfg("IMG2EXCEL_WCRED", "wincred", _pick_wincred())

# ⚠️ 代理不硬编码：由 pick_proxy() 在发版开始时**实测**挑选，None 表示直连。
#    历史教训：写死端口会因代理关闭/换端口而在最后一步 push 挂掉；
#    「探到端口开着就用」也是错的（TCP 能连不代表能出网）。
PROXY = None

# Inno Setup 编译器：本机安装位置走配置，仓库里只留通用安装路径。
ISCC_CANDIDATES = ([_cfg("IMG2EXCEL_ISCC", "iscc")] +
                   [r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
                    r"C:\Program Files\Inno Setup 7\ISCC.exe",
                    r"C:\Program Files\Inno Setup 6\ISCC.exe"])


def log(msg):
    print(">> " + msg, flush=True)


def run(cmd, cwd=None, check=True, env=None, timeout=None):
    try:
        p = subprocess.run(cmd, cwd=cwd or ROOT, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", env=env, timeout=timeout)
    except subprocess.TimeoutExpired:
        raise SystemExit("命令超时(%ss)：%s" % (timeout, os.path.basename(str(cmd[0]))))
    if check and p.returncode != 0:
        print((p.stdout or "")[-2000:])
        print((p.stderr or "")[-2000:])
        raise SystemExit("命令失败(%d): %s" % (p.returncode, os.path.basename(str(cmd[0]))))
    return p


# ---------------- 代理 ----------------
def proxy_candidates():
    """按优先级给出代理候选；None 表示「不用代理、直连」。

    ① 环境变量 RELEASE_PROXY —— 显式指定，最高优先级；**空串表示强制直连**
    ② 本机代理软件常用端口（10808 / 10809 / 7890 / 7897）
    ③ 环境变量里的 https_proxy / http_proxy（放最后：常是沙箱注入的，未必能出网）
    ④ None（直连）
    """
    explicit = os.environ.get("RELEASE_PROXY")
    if explicit is not None:
        return [explicit.strip() or None]

    out = []
    for port in (10808, 10809, 7890, 7897):
        u = "http://127.0.0.1:%d" % port
        if u not in out:
            out.append(u)
    for key in ("https_proxy", "HTTPS_PROXY", "http_proxy", "HTTP_PROXY"):
        v = (os.environ.get(key) or "").strip()
        if v and v not in out:
            out.append(v)
    out.append(None)
    return out


def proxy_works(proxy, timeout=6):
    """实测该代理能否访问 GitHub；proxy=None 时测的是直连。

    ⚠️ 两个域名**都必须通**才算可用：
      · api.github.com —— 建 Release、传资产走它
      · github.com     —— git push 走它（CONNECT 隧道到 443）
    踩过的坑：某个代理只放行 api.github.com，于是被判为「可用」，
    打包/签名/元数据全做完，最后 git push 报 CONNECT tunnel failed 502。
    """
    for url in ("https://api.github.com", "https://github.com"):
        try:
            handlers = ([urllib.request.ProxyHandler({"http": proxy, "https": proxy})]
                        if proxy else [urllib.request.ProxyHandler({})])
            op = urllib.request.build_opener(*handlers)
            req = urllib.request.Request(url, headers={"User-Agent": "release-probe"})
            with op.open(req, timeout=timeout) as resp:
                if getattr(resp, "status", 0) != 200:
                    return False
        except Exception:  # noqa: BLE001
            return False
    return True


def pick_proxy(log_fn=None):
    """挑第一个**实测能用的**代理；全不行就直连。"""
    for cand in proxy_candidates():
        if cand:
            try:
                hp = cand.replace("http://", "").replace("https://", "")
                h, pt = hp.rsplit(":", 1)
                s = socket.create_connection((h, int(pt)), timeout=2)
                s.close()
            except Exception:  # noqa: BLE001
                continue  # 端口都不通，直接跳过，省一次 ~12s 探测
        if proxy_works(cand):
            if log_fn:
                log_fn("代理：%s" % (cand or "不使用（直连）"))
            return cand
    if log_fn:
        log_fn("代理：候选全部不可用，按直连处理")
    return None


def git_env(proxy=None):
    """git 子进程环境：补齐 PATH，并把代理统一成 proxy 参数指定的那一个。

    为什么要显式清掉环境变量里的代理：
      ① git config 的 http.proxy 优先级**高于**环境变量，残留一个失效代理会让
         push 直接失败；
      ② 环境变量里的代理是别人注入的，未必能出网。
      所以一律清空，再由 proxy 参数写入唯一的一个，行为可预期。
    """
    e = dict(os.environ)
    if GIT:
        e["PATH"] = os.path.dirname(GIT) + os.pathsep + e.get("PATH", "")
    for k in ("http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY",
              "all_proxy", "ALL_PROXY"):
        e.pop(k, None)
    if proxy:
        e["http_proxy"] = e["https_proxy"] = proxy
        e["HTTP_PROXY"] = e["HTTPS_PROXY"] = proxy
    return e


# ---------------- 版本号 ----------------
def read_version():
    s = open(CORE_PY, encoding="utf-8").read()
    m = re.search(r'^APP_VERSION\s*=\s*"([\d.]+)"', s, re.M)
    if not m:
        raise SystemExit("在 core.py 中找不到 APP_VERSION 常量")
    return m.group(1)


def write_version(v):
    s = open(CORE_PY, encoding="utf-8").read()
    s2 = re.sub(r'^APP_VERSION\s*=\s*"[\d.]+"', 'APP_VERSION = "%s"' % v,
                s, count=1, flags=re.M)
    if s2 == s:
        raise SystemExit("版本号写入失败")
    open(CORE_PY, "w", encoding="utf-8").write(s2)


def bump(ver, kind):
    parts = [int(x) for x in ver.split(".")]
    while len(parts) < 3:
        parts.append(0)
    if kind == "major":
        parts = [parts[0] + 1, 0, 0]
    elif kind == "minor":
        parts = [parts[0], parts[1] + 1, 0]
    else:
        parts = [parts[0], parts[1], parts[2] + 1]
    return ".".join(str(x) for x in parts)


def build_env():
    """打包/签名子进程用的环境：剥掉外部注入的 PYTHONPATH 和代理变量。

    ① PYTHONPATH：开发沙箱会通过它注入 sitecustomize.py，给 shutil.rmtree /
       os.remove 挂上「安全删除」钩子。PyInstaller 覆盖 dist 里的旧产物时会踩到
       钩子，于是 **打包静默失败、留下上一版 exe**，而脚本还以为是新包
       （症状：体积和上版一模一样、时间戳是旧的）。剥掉后行为与命令行直跑一致。

    ② 代理变量：沙箱会注入 http_proxy/https_proxy 指向本地代理，而该代理通常
       不放行 osslsigncode 的 TLS 隧道 → 所有时间戳服务器全部失败且每次都要等
       完整 TCP 超时（5 服务器 × 3 轮），签名看起来像「卡死」，把脚本拖过超时
       而被杀（2026-09-24 实测踩坑：600s 超时，实际单次直连只要 4s）。
    """
    e = dict(os.environ)
    for k in ("PYTHONPATH", "PYTHONSTARTUP",
              "http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY",
              "all_proxy", "ALL_PROXY"):
        e.pop(k, None)
    return e


def _assert_fresh(exe, t0, label):
    """确认产物是本次新建的，而不是没删掉的上版残留。"""
    if not os.path.exists(exe):
        raise SystemExit("%s：产物不存在 %s" % (label, exe))
    if os.path.getmtime(exe) < t0 - 2:
        raise SystemExit("%s：产物是旧的（PyInstaller 未真正重建），中止发版" % label)


# ---------------- 单文件版 ----------------
def build_onefile():
    log("打包【单文件版】…")
    os.makedirs(ONEFILE_DIR, exist_ok=True)
    bld = os.path.join(CACHE, "bld_onefile_%s" % time.strftime("%Y%m%d_%H%M%S"))
    spec = os.path.join(ROOT, "图片表格转Excel助手.spec")
    t0 = time.time()
    p = run([PYINSTALLER, spec, "--noconfirm", "--clean",
             "--distpath", ONEFILE_DIR, "--workpath", bld],
            check=False, env=build_env())
    exe = os.path.join(ONEFILE_DIR, EXE_NAME)
    if p.returncode != 0:
        print((p.stdout or "")[-2500:])
        print((p.stderr or "")[-2500:])
        raise SystemExit("单文件版构建失败（返回码 %s）" % p.returncode)
    _assert_fresh(exe, t0, "单文件版")
    size = os.path.getsize(exe) / 1048576
    log("  单文件版完成：%.1f MB" % size)
    if size < 60:
        raise SystemExit("产物仅 %.1f MB，疑似 tkinter/onnxruntime 未打进来，中止发版" % size)
    return exe


# ---------------- 目录版（安装包负载） ----------------
def build_onedir():
    log("打包【目录版】（安装包负载）…")
    os.makedirs(ONEDIR_DIR, exist_ok=True)
    bld = os.path.join(CACHE, "bld_onedir_%s" % time.strftime("%Y%m%d_%H%M%S"))
    spec = os.path.join(ROOT, "安装版.spec")
    t0 = time.time()
    p = run([PYINSTALLER, spec, "--noconfirm", "--clean",
             "--distpath", ONEDIR_DIR, "--workpath", bld],
            check=False, env=build_env())
    d = os.path.join(ONEDIR_DIR, "图片表格转Excel助手")
    target = os.path.join(d, EXE_NAME)
    if p.returncode != 0:
        print((p.stdout or "")[-2500:])
        print((p.stderr or "")[-2500:])
        raise SystemExit("目录版构建失败（返回码 %s）" % p.returncode)
    _assert_fresh(target, t0, "目录版")
    total = sum(os.path.getsize(os.path.join(r, f))
                for r, _dirs, fs in os.walk(d) for f in fs)
    log("  目录版完成：%.1f MB（%d 个文件）"
        % (total / 1048576, sum(len(fs) for _r, _d, fs in os.walk(d))))
    return d


# ---------------- 安装版 ----------------
def build_installer(ver):
    iscc = next((p for p in ISCC_CANDIDATES if os.path.exists(p)), None)
    if not iscc:
        log("!! 未找到 ISCC.exe，跳过安装版")
        return None
    payload = os.path.join(ONEDIR_DIR, "图片表格转Excel助手")
    if not os.path.exists(os.path.join(payload, EXE_NAME)):
        log("!! 目录版负载不存在（%s），跳过安装版" % payload)
        return None
    log("编译【安装版】…")
    os.makedirs(INSTALLER_DIR, exist_ok=True)
    iss = os.path.join(ROOT, "installer.iss")
    # ⚠️ Inno Setup 7 的命令行选项是 `-o<path>` / `-d<name>=<value>`，
    #    旧的 `/O` `/D` 写法在 ISCC 7 上会被当成第二个脚本文件名而报错
    #    「You may not specify more than one script filename」。
    p = run([iscc, "-dMyAppVersion=%s" % ver,
             "-o" + INSTALLER_DIR.replace("\\", "/"), iss],
            cwd=ROOT, check=False)
    exe = os.path.join(INSTALLER_DIR, "Img2ExcelTool_v%s_setup.exe" % ver)
    if not os.path.exists(exe):
        txt = (p.stdout or "") + "\n" + (p.stderr or "")
        errs = [l for l in txt.splitlines()
                if re.search(r"error|Error|not found|cannot|Cannot|无效|找不到", l)]
        print("\n".join(errs[-14:]) if errs else txt[-2000:])
        raise SystemExit("安装版编译失败")
    log("  安装版完成：%.1f MB" % (os.path.getsize(exe) / 1048576))
    return exe


# ---------------- 自检 ----------------
def selftest(exe, label, expect_ver=None):
    """用产物跑一次无界面自检：确认包内引擎可用，且版本号就是本次目标版本。"""
    log("校验 %s 可运行…" % label)
    out = os.path.join(CACHE, "selftest_%s_%s"
                       % (label, time.strftime("%Y%m%d_%H%M%S")))
    os.makedirs(out, exist_ok=True)
    v = run([exe, "--version"], check=False, timeout=300)
    vout = (v.stdout or "").strip()
    if "1." not in vout and "2." not in vout:
        raise SystemExit("%s --version 无输出，产物异常" % label)
    # ⚠️ 版本闸门第二道：确认打进包里的是本次目标版本（防「打包沿用了旧产物」）
    if expect_ver and expect_ver not in vout:
        raise SystemExit("%s 的 --version 输出「%s」，不含目标版本 %s，中止发版"
                         % (label, vout, expect_ver))
    log("  %s" % vout)

    # 识别自检用的样本目录：本机配置里指定真实票据目录；未配置时退回项目内 verify/。
    # ⚠️ 仓库里绝不写死样本路径 —— 那会把开发机的目录结构一起公开出去。
    sample = _cfg("IMG2EXCEL_SAMPLE", "sample_dir", os.path.join(ROOT, "verify"))
    if not os.path.isdir(sample):
        log("  !! 找不到测试样本目录（%s），跳过识别自检" % sample)
        return True
    exts = (".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp", ".pdf")
    pics = sorted(f for f in glob.glob(os.path.join(sample, "*"))
                  if os.path.splitext(f)[1].lower() in exts)[:2]
    if not pics:
        log("  !! 样本目录里没有图片（%s），跳过识别自检" % sample)
        return True
    log("  样本目录：%s" % sample)
    tgt = os.path.join(out, "in")
    os.makedirs(tgt, exist_ok=True)
    for f in pics:
        shutil.copy2(f, tgt)
    xlsx = os.path.join(out, "result.xlsx")
    run([exe, "--selftest", tgt, xlsx], check=False, timeout=900)
    logf = os.path.splitext(xlsx)[0] + ".selftest.log"
    text = open(logf, encoding="utf-8", errors="replace").read() if os.path.exists(logf) else ""
    tail = [l for l in text.splitlines() if l.strip()][-4:]
    for line in tail:
        log("  | " + line)
    if not os.path.exists(xlsx):
        raise SystemExit("%s 识别自检未产出 xlsx，中止" % label)
    log("  %s 自检通过" % label)
    return True


# ---------------- 签名 ----------------
def sign(exe, title):
    if not os.path.exists(SIGN_PY):
        log("!! 找不到签名脚本，跳过签名：" + SIGN_PY)
        return False
    log("签名：%s" % os.path.basename(exe))
    # 传干净环境（无坏代理）；sign.py 内部单次 TSA 25s 硬超时，
    # 最坏 5 服务器 × 3 轮 ≈ 400s，外层给 900s 留足余量。
    p = run([PY, SIGN_PY, exe, title, REPO_URL], check=False,
            env=build_env(), timeout=900)
    out = (p.stdout or "") + "\n" + (p.stderr or "")
    ok = p.returncode == 0 and "签名完成" in out
    if not ok:
        print(out[-900:])
        log("!! 签名未成功（请检查证书）")
    else:
        log("  签名 ok")
    return ok


# ---------------- 交付 ----------------
def copy_delivery(onefile, installer, ver):
    os.makedirs(DELIVERY, exist_ok=True)
    # 交付文件名带版本号：放桌面/U 盘一眼看出版本，不会新旧混淆。
    # GitHub Release 资产名保持 ASCII（见 ASSET_* 常量）。
    base = APP_NAME
    # ⚠️ 单文件版交付名**不带** `_单文件版` 后缀：自动更新会把绿色版归一成
    #    `图片表格转Excel助手_vX.Y.Z.exe`，交付时就用同名，用户不会遇到「名字突变」。
    pairs = [(onefile, "%s_v%s.exe" % (base, ver))]
    if installer:
        pairs.append((installer, "%s_v%s_安装版.exe" % (base, ver)))
    # 历史命名归档，避免新旧并存让用户不知道点哪个
    legacy_dir = os.path.join(DELIVERY, "_旧版备份")
    new_names = {nm for _, nm in pairs}
    legacy_files = []
    for pat in ("%s-单文件版.exe", "%s-安装版.exe",
                "%s_v*_单文件版.exe", "%s_v*_安装版.exe"):
        legacy_files += glob.glob(os.path.join(DELIVERY, pat % base))
    for p in legacy_files:
        nm = os.path.basename(p)
        if nm in new_names:
            continue
        try:
            os.makedirs(legacy_dir, exist_ok=True)
            shutil.move(p, os.path.join(legacy_dir, nm))
            log("交付 -> 旧命名归档 %s" % nm)
        except (OSError, shutil.Error) as e:
            log("!! 旧命名归档失败（忽略）：%s" % e)
    for src, nm in pairs:
        if src and os.path.exists(src):
            dst = os.path.join(DELIVERY, nm)
            shutil.copy2(src, dst)
            log("交付 -> %s（%.1f MB）" % (dst, os.path.getsize(dst) / 1048576))
    manual = os.path.join(ROOT, "使用说明.md")
    if os.path.exists(manual):
        shutil.copy2(manual, os.path.join(DELIVERY, "使用说明.md"))
        log("交付 -> %s" % os.path.join(DELIVERY, "使用说明.md"))


# ---------------- GitHub 发布 ----------------
def get_token():
    """取 GitHub PAT：优先缓存，其次系统 Git 的 wincred。"""
    f = os.path.join(CACHE, "token.txt")
    if os.path.exists(f):
        t = open(f, encoding="utf-8").read().strip()
        if t:
            return t
    if not WCRED or not os.path.exists(WCRED):
        raise SystemExit("找不到 git-credential-wincred.exe，无法取 token"
                         "（可用环境变量 IMG2EXCEL_WCRED 指定其路径）")
    out = subprocess.run([WCRED, "get"], input="protocol=https\nhost=github.com\n\n",
                         capture_output=True, text=True, timeout=30).stdout
    for line in out.splitlines():
        if line.startswith("password="):
            t = line[9:].strip()
            os.makedirs(CACHE, exist_ok=True)
            open(f, "w", encoding="utf-8").write(t)
            return t
    raise SystemExit("拿不到 GitHub token（wincred 里没有 github.com 凭据）")


def api(path, token, data=None, method=None, timeout=120):
    proxy = urllib.request.ProxyHandler({"http": PROXY, "https": PROXY})
    op = urllib.request.build_opener(proxy)
    h = {"Authorization": "Bearer " + token, "User-Agent": "curl/8",
         "Accept": "application/vnd.github+json"}
    body = None
    if data is not None:
        body = data if isinstance(data, bytes) else json.dumps(data).encode()
        h["Content-Type"] = "application/json"
    req = urllib.request.Request("https://api.github.com" + path, data=body,
                                 headers=h, method=method)
    return op.open(req, timeout=timeout)


def ensure_repo(token):
    """仓库不存在则创建（public）。不带 auto_init，避免后续推送时远端已有初始提交。"""
    try:
        json.load(api("/repos/%s" % OWNER_REPO, token))
        log("仓库已存在：%s" % OWNER_REPO)
        return False
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
    log("创建仓库 %s…" % OWNER_REPO)
    json.load(api("/user/repos", token, {
        "name": OWNER_REPO.split("/", 1)[1],
        "description": "把照片/PDF 里的表格识别出来汇总成 Excel。完全离线运行。",
        "private": False,
        "has_issues": True, "has_wiki": False, "has_projects": False,
        "auto_init": False,
    }))
    log("  仓库已创建")
    return True


def git(*args, check=True):
    return run([GIT] + list(args), cwd=ROOT, check=check, env=git_env(PROXY))


def git_init_if_needed():
    if os.path.isdir(os.path.join(ROOT, ".git")):
        return
    log("初始化 git 仓库…")
    git("init", "-q", "-b", MAIN_BRANCH)
    git("config", "user.name", "jianRY")
    git("config", "user.email", "jianRY@users.noreply.github.com")


def git_commit_tag_push(ver, token):
    log("git 提交 + 打 tag…")
    git("add", "-A")
    st = git("status", "--short").stdout.strip()
    if not st:
        log("  无文件变更，跳过 commit")
    else:
        git("commit", "-q", "-m", "release: v%s" % ver)
    git("tag", "-a", "v" + ver, "-m", "%s v%s" % (APP_TITLE, ver), check=False)
    log("推送%s…" % ("（经 %s）" % PROXY if PROXY else "（直连）"))
    # ⚠️ 先清掉仓库里可能残留的 http.proxy：git config 的优先级**高于**环境变量，
    #    残留一个已失效的代理会直接让 push 失败。
    git("config", "--unset", "http.proxy", check=False)
    git("config", "--unset", "https.proxy", check=False)
    url = "https://x-access-token:%s@github.com/%s.git" % (token, OWNER_REPO)
    try:
        git("remote", "remove", "origin", check=False)
        git("remote", "add", "origin", "https://github.com/%s.git" % OWNER_REPO)
        print("   $ git push <token>@github.com/%s.git %s v%s" % (OWNER_REPO, MAIN_BRANCH, ver))
        p = run([GIT, "push", url, "%s:%s" % (MAIN_BRANCH, MAIN_BRANCH)],
                check=False, env=git_env(PROXY))
        if p.returncode != 0:
            print(re.sub(r"x-access-token:[^@\s]+@", "x-access-token:***@",
                         p.stderr or "")[-800:])
            log("  提示：可显式指定代理重试 —— set RELEASE_PROXY=http://127.0.0.1:7890")
            log("        或强制直连 —— set RELEASE_PROXY=")
            raise SystemExit("push 分支失败")
        # ⚠️ tag 推送必须校验返回值：若 tag 没推上去，建出的 Release 可能指向错误提交。
        p2 = run([GIT, "push", url, "v" + ver], check=False, env=git_env(PROXY))
        if p2.returncode != 0:
            err = (p2.stderr or "")
            if "already exists" in err or "up to date" in err or "Everything up-to-date" in err:
                log("  tag v%s 已存在于远端，继续" % ver)
            else:
                print(re.sub(r"x-access-token:[^@\s]+@", "x-access-token:***@", err)[-600:])
                raise SystemExit("push tag v%s 失败，中止（避免 Release 指向错误提交）" % ver)
        log("  推送完成")
    finally:
        git("config", "--unset", "http.proxy", check=False)
        git("config", "--unset", "https.proxy", check=False)


def release_body(ver, onefile_mb, setup_mb, notes=""):
    """Release 正文：本次更新 + 下载表 + 使用提示。

    只写用户看得懂的内容 —— 打包方式、自动更新的技术原理、排查过程这类
    内部细节一律不写进来。
    """
    parts = []
    if notes:
        parts.append("## 本次更新\n\n%s\n\n" % notes.strip())
    parts.append(
        "## 下载\n\n"
        "| 版本 | 文件 | 大小 | 说明 |\n|---|---|---|---|\n"
        "| **单文件版**（绿色版） | `Img2ExcelTool_v%s_onefile.exe` | %.1f MB | 双击即用，无需安装，支持程序内一键升级 |\n"
        "| **安装版** | `Img2ExcelTool_v%s_setup.exe` | %.1f MB | 装到 Program Files，带开始菜单/桌面快捷方式/卸载 |\n\n"
        "两个版本的识别功能完全一致。安装版升级请重新下载安装包。\n\n"
        % (ver, onefile_mb, ver, setup_mb))
    parts.append(
        "## 功能\n\n"
        "- 支持 JPG / PNG / BMP / TIF / WEBP 图片与 PDF（自动逐页处理）\n"
        "- RapidOCR 中文识别 + SLANet-Plus 表格结构还原，**完全离线**\n"
        "- 多文件可合并为同一工作表，也可每个文件单独工作表\n"
        "- 金额千分位误判自动修正，并在「列N原文」保留原始文本便于核对\n"
        "- 纯数字自动转成数值格式（可求和、可排序），日期归一为 `YYYY/MM/DD`\n\n")
    parts.append(
        "## 注意\n\n"
        "- 识别质量取决于原图：拍照尽量平整、光线均匀、正对拍摄\n"
        "- 票据类图片属「无表格线」版式，列的含义需要人工判断\n"
        "- **重要金额请人工复核**，用「列N原文」列对照\n\n"
        "exe 已使用 jianRY 自签名证书签名（SHA256 + RFC3161 时间戳）。\n")
    return "".join(parts)


def create_release(token, ver, body):
    try:
        d = json.load(api("/repos/%s/releases/tags/v%s" % (OWNER_REPO, ver), token))
        log("Release v%s 已存在，复用 id=%s" % (ver, d["id"]))
        # 重跑发版时正文也要跟着刷新，否则线上还挂着旧说明
        try:
            json.load(api("/repos/%s/releases/%s" % (OWNER_REPO, d["id"]), token,
                          {"body": body}, method="PATCH"))
            log("  已刷新 Release 说明")
        except Exception as e:  # noqa: BLE001
            log("  !! Release 说明刷新失败（忽略）：%s" % e)
        return d["id"]
    except urllib.error.HTTPError as e:
        if e.code != 404:
            raise
    log("创建 Release v%s…" % ver)
    d = json.load(api("/repos/%s/releases" % OWNER_REPO, token, {
        "tag_name": "v" + ver,
        "name": "%s v%s" % (APP_TITLE, ver),
        "body": body,
        "draft": False, "prerelease": False,
    }))
    return d["id"]


def upload_asset(token, rid, path, name):
    size = os.path.getsize(path)
    log("上传资产 %s（%.1f MB）…" % (name, size / 1048576))
    # 同名资产先删除，保证脚本可重入
    try:
        d = json.load(api("/repos/%s/releases/%s/assets" % (OWNER_REPO, rid), token))
        for a in d:
            if a["name"] == name:
                api("/repos/%s/releases/assets/%s" % (OWNER_REPO, a["id"]),
                    token, method="DELETE").read()
                log("  已删除同名旧资产 %s" % name)
    except Exception as e:  # noqa: BLE001
        log("  同名资产检查失败（忽略）：%s" % e)
    op = urllib.request.build_opener(
        urllib.request.ProxyHandler({"http": PROXY, "https": PROXY}))
    with open(path, "rb") as f:
        blob = f.read()
    url = ("https://uploads.github.com/repos/%s/releases/%s/assets?name=%s"
           % (OWNER_REPO, rid, urllib.parse.quote(name)))
    for attempt in (1, 2, 3):
        try:
            req = urllib.request.Request(url, data=blob, method="POST", headers={
                "Authorization": "Bearer " + token, "User-Agent": "curl/8",
                "Content-Type": "application/octet-stream",
                "Content-Length": str(len(blob)),
            })
            d = json.load(op.open(req, timeout=1800))
            log("  ok -> %s (%.1f MB)" % (d["name"], d["size"] / 1048576))
            return True
        except Exception as e:  # noqa: BLE001
            log("  第 %d 次上传失败：%s" % (attempt, e))
            time.sleep(4)
    return False


# ---------------- 更新元数据（客户端自动更新的数据源） ----------------
def _sha256_of(path, chunk=1024 * 1024):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for blk in iter(lambda: f.read(chunk), b""):
            h.update(blk)
    return h.hexdigest()


def make_update_json(ver, onefile, installer, notes):
    """生成 update.json —— 随 Release 作为附件上传，客户端靠它判断与校验更新。

    字段契约（客户端 fetch_update_info 直接读这些键，改键名等于断更）：
        version / url（主源）/ fallback_url（兜底）/ size / sha256 / notes

    ⚠️ 两个源的语义别搞反（2026-09-22 对调过）：
      url          = GitHub Release 直链 → 客户端再展开成加速镜像变体并测速择优
      fallback_url = 自有站直链，只在镜像与原站都不可用时才用
    新旧客户端都按 url → fallback_url 顺序试，所以对调不破坏已发布版本。

    ⚠️ 附件名必须就叫 update.json：服务器脚本按这个名字去
      releases/latest/download/update.json 抄到 /updates/<APP_KEY>.json，改名即断链。
    """
    tag = "v" + ver
    gh = "https://github.com/" + OWNER_REPO
    base_asset = ASSET_ONEFILE % ver
    setup_asset = ASSET_SETUP % ver
    site = SITE_URL.rstrip("/")

    data = {
        "app": APP_KEY,
        "name": APP_TITLE,
        "version": ver,
        "asset": base_asset,
        "notes": notes or "",
        "url": "%s/releases/download/%s/%s" % (gh, tag, base_asset),
        "fallback_url": "%s/files/%s" % (site, base_asset),
        "release_url": "%s/releases/tag/%s" % (gh, tag),
        "site_url": site + "/",
        "size": os.path.getsize(onefile),
        "sha256": _sha256_of(onefile),
        "published": time.strftime("%Y-%m-%d %H:%M:%S"),
    }
    if installer and os.path.exists(installer):
        data["setup_url"] = "%s/releases/download/%s/%s" % (gh, tag, setup_asset)
        data["setup_fallback_url"] = "%s/files/%s" % (site, setup_asset)
        data["setup_size"] = os.path.getsize(installer)
        data["setup_sha256"] = _sha256_of(installer)

    os.makedirs(CACHE, exist_ok=True)
    path = os.path.join(CACHE, "update.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    log("更新元数据 -> %s" % path)
    log("  url          = %s" % data["url"])
    log("  fallback_url = %s" % data["fallback_url"])
    log("  size=%d  sha256=%s…" % (data["size"], data["sha256"][:12]))
    return path, data


def publish(onefile, installer, ver, notes=""):
    """推 GitHub + 建 Release + 上传两个 exe + update.json。"""
    if not GIT:
        log("!! 找不到可用的 git.exe，跳过 GitHub 发布")
        return False
    token = get_token()
    ensure_repo(token)
    git_init_if_needed()
    git_commit_tag_push(ver, token)

    onefile_mb = os.path.getsize(onefile) / 1048576 if onefile and os.path.exists(onefile) else 0
    setup_mb = (os.path.getsize(installer) / 1048576
                if installer and os.path.exists(installer) else 0)

    # 先算元数据：sha256 必须在 exe 定稿之后算，否则哈希与资产对不上
    meta_path = None
    if onefile and os.path.exists(onefile):
        meta_path, _ = make_update_json(ver, onefile, installer, notes)
    else:
        log("!! 没有单文件版产物，跳过 update.json（自动更新将不可用）")

    rid = create_release(token, ver, release_body(ver, onefile_mb, setup_mb, notes))
    ok = True
    if onefile:
        ok &= upload_asset(token, rid, onefile, ASSET_ONEFILE % ver)
    if installer:
        ok &= upload_asset(token, rid, installer, ASSET_SETUP % ver)
    if meta_path:
        # ⚠️ 上传失败**不阻断发版**：自动更新仍可用（检查退回 GitHub API，只是慢些）
        if not upload_asset(token, rid, meta_path, ASSET_META):
            log("!! update.json 上传失败（不阻断发版）—— 自动更新会退回 GitHub API 检查")
    log("Release 页面：%s/releases/tag/v%s" % (REPO_URL, ver))
    return ok


# ---------------- 主流程 ----------------
def main():
    ap = argparse.ArgumentParser(description="%s 一键发版" % APP_TITLE)
    ap.add_argument("--version", default=None, help="指定版本号（写回 core.py）")
    ap.add_argument("--bump", choices=["patch", "minor", "major"], default=None)
    ap.add_argument("--no-installer", action="store_true", help="只出单文件版")
    ap.add_argument("--skip-sign", action="store_true", help="跳过签名（仅调试）")
    ap.add_argument("--skip-selftest", action="store_true", help="跳过产物自检")
    ap.add_argument("--no-publish", action="store_true", help="只做本地，不推 GitHub")
    ap.add_argument("--notes", default="", help="本次更新说明（写进 Release 正文与 update.json）")
    ap.add_argument("--notes-file", default=None, help="从文件读取更新说明（优先于 --notes）")
    args = ap.parse_args()

    global PROXY
    PROXY = pick_proxy(log)

    os.makedirs(CACHE, exist_ok=True)
    os.makedirs(DIST, exist_ok=True)

    notes = args.notes
    if args.notes_file:
        notes = open(args.notes_file, encoding="utf-8").read()
    notes = (notes or "").strip()

    cur = read_version()
    ver = args.version or (bump(cur, args.bump) if args.bump else cur)
    if ver != cur:
        write_version(ver)
        log("版本号：%s -> %s" % (cur, ver))
    else:
        log("版本号：v%s（未变更）" % ver)

    # ⚠️ 版本闸门：写回后立刻读回核对。源码常量必须与 tag / 资产名 / update.json 同源，
    #    否则客户端「远端版本 > 本地版本」永远成立 → 每次启动都弹更新（升级死循环）。
    back = read_version()
    if back != ver:
        raise SystemExit("版本常量写回失败（core.py 仍是 %s，目标 %s），中止发版" % (back, ver))

    onefile = build_onefile()
    if not args.skip_selftest:
        selftest(onefile, "单文件版", ver)
    if not args.skip_sign:
        sign(onefile, APP_TITLE)

    installer = None
    if not args.no_installer:
        build_onedir()
        installer = build_installer(ver)
        if installer and not args.skip_sign:
            sign(installer, APP_TITLE + " 安装版")

    copy_delivery(onefile, installer, ver)

    published = False
    if not args.no_publish:
        published = publish(onefile, installer, ver, notes)

    log("")
    log("=" * 62)
    log("发版完成 v%s" % ver)
    log("  单文件版：%s" % onefile)
    if installer:
        log("  安装版  ：%s" % installer)
    log("  交付目录：%s" % DELIVERY)
    if published:
        log("  Release ：%s/releases/tag/v%s" % (REPO_URL, ver))
    log("=" * 62)


if __name__ == "__main__":
    main()
