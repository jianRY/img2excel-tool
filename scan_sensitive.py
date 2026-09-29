#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""入库前敏感信息扫描 —— 防止开发机信息与凭据被推到公开仓库。

为什么必须有它：本仓库是 public。一旦内部路径/凭据进了提交，`git rm` 只是从
最新快照里删掉，**历史依然公开可查**（本项目就踩过这个坑：内部记忆目录进过一次
提交，后来才发现历史没清）。把扫描卡在提交之前，比事后清洗历史便宜得多。

用法：
    python scan_sensitive.py              # 扫当前已入库文件（git ls-files）
    python scan_sensitive.py --staged     # 只扫暂存区（pre-commit 钩子用）
    python scan_sensitive.py --all        # 连未入库文件一起扫（含 .gitignore 排除项）
    python scan_sensitive.py --history    # 扫全部 git 历史，找历史残留

退出码：0 = 干净；1 = 有命中（提交应中止）。

激活 pre-commit 钩子（每台机器只需一次）：
    git config core.hooksPath .githooks

⚠️ 本脚本内的规则**全部在运行时构造**，源码里不出现任何真实的用户名 / 家目录 /
   开发机路径 —— 否则扫描脚本自己就成了泄露源。
"""
import argparse
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))

# 脚本自身与钩子文件里必然出现规则关键字，不参与扫描
SELF_SKIP = {"scan_sensitive.py", "pre-commit", "scan_sensitive"}

# 允许出现的通用系统路径前缀（任何 Windows 机器都一样，不含个人身份）
PATH_ALLOW = (
    "Program Files (x86)",
    "Program Files",
    "C:\\Windows",
    "C:/Windows",
    "ProgramData",
    "%APPDATA%",
    "%LOCALAPPDATA%",
    "%TEMP%",
    "%USERPROFILE%",
    "{autopf}",
    "{localappdata}",
    "{userappdata}",
)

# 只扫这些扩展名（其余按二进制跳过）
TEXT_EXTS = {
    ".py", ".md", ".txt", ".json", ".iss", ".spec", ".cfg", ".ini", ".toml",
    ".yml", ".yaml", ".bat", ".cmd", ".ps1", ".sh", ".html", ".css", ".js",
    ".ts", ".xml", ".csv", ".example", ".gitignore", ".gitattributes",
}

MAX_SIZE = 8 * 1024 * 1024


def _pick_git():
    for c in (r"C:\Program Files\Git\cmd\git.exe",
              r"C:\Program Files (x86)\Git\cmd\git.exe"):
        if os.path.exists(c):
            return c
    for d in os.environ.get("PATH", "").split(os.pathsep):
        p = os.path.join(d, "git.exe" if os.name == "nt" else "git")
        if os.path.exists(p):
            return p
    return "git"


GIT = _pick_git()


def _git(*args):
    """跑一条 git 命令并返回 stdout 文本；失败返回 None。

    加 core.quotepath=false 是为了中文路径不被转义成八进制。
    """
    try:
        p = subprocess.run([GIT, "-c", "core.quotepath=false"] + list(args),
                           cwd=ROOT, capture_output=True)
    except Exception:  # noqa: BLE001
        return None
    if p.returncode != 0:
        return None
    return p.stdout.decode("utf-8", errors="replace")


# ---------------- 规则 ----------------
def build_rules():
    """运行时构造规则表：[(规则名, 正则, 修复建议), ...]。"""
    rules = []

    # ① 本机身份：用户名 / 家目录 / 机器名（都从环境取，不写死在源码里）
    user = (os.environ.get("USERNAME") or os.environ.get("USER") or "").strip()
    home = os.path.expanduser("~")
    host = (os.environ.get("COMPUTERNAME") or os.environ.get("HOSTNAME") or "").strip()
    if user and user.lower() not in ("user", "administrator", "root"):
        rules.append(("本机用户名", re.compile(re.escape(user), re.I),
                      "别把它写进代码或注释；路径改成相对路径 / 环境变量"))
    if home and len(home) > 4:
        rules.append(("家目录绝对路径",
                      re.compile(re.escape(home).replace("\\\\", "[\\\\/]"), re.I),
                      "改用 os.path.expanduser('~') 或相对路径"))
    if host and len(host) > 3:
        rules.append(("本机计算机名", re.compile(re.escape(host), re.I),
                      "删掉；机器名没有出现在公开仓库的理由"))

    # ② 开发机工作区路径：本脚本所在目录及其上级，运行时才算出
    here = ROOT.replace("\\", "[\\\\/]")
    parent = os.path.dirname(ROOT)
    if len(parent) > 4 and os.path.dirname(parent):
        rules.append(("开发机工作区路径",
                      re.compile(re.escape(parent).replace("\\\\", "[\\\\/]"), re.I),
                      "改成相对路径；本机专属路径写进 .pybuild_cache/local_config.json"))
    rules.append(("项目所在盘符路径", re.compile(here, re.I),
                  "写成相对路径（os.path.dirname(__file__) 等）"))

    # ③ 任何形如 X:\\Users\\<名字> 的路径（即使不是当前用户）
    rules.append((r"用户目录绝对路径", re.compile(r"[A-Za-z]:[\\/]Users[\\/]", re.I),
                  "改用 %USERPROFILE% / ~ / 相对路径"))

    # ④ 通用盘符绝对路径（白名单里的系统路径会在后面过滤掉）
    rules.append(("盘符绝对路径", re.compile(r"[A-Za-z]:[\\/][A-Za-z_][^\s\"'<>|,;)]{3,}"),
                  "改成相对路径或环境变量；确需固定就放 local_config.json（不入库）"))

    # ⑤ 凭据特征
    rules.append(("GitHub 令牌",
                  re.compile(r"\b(gh[pousr]_[A-Za-z0-9]{20,}|github_pat_[A-Za-z0-9_]{20,})"),
                  "立刻吊销该令牌并换成从凭据管理器读取"))
    rules.append(("私钥内容",
                  re.compile(r"-----BEGIN [A-Z ]*PRIVATE KEY-----"),
                  "私钥绝不入库；改用签名工具从库外读取"))
    rules.append(("疑似硬编码口令",
                  re.compile(r"(?i)\b(pass(word|wd)?|secret|api[_-]?key|token)\b\s*[:=]\s*"
                             r"[\"'][^\"'\s{}<>$]{6,}[\"']"),
                  "改成从环境变量 / 凭据管理器读取"))
    rules.append(("内联 URL 凭据",
                  re.compile(r"[a-z]+://[^/\s:@]{3,}:[^/\s:@]{3,}@", re.I),
                  "别把账号密码拼进 URL；用凭据管理器"))

    # ⑥ 内部资料 / 不该公开的文件名
    rules.append(("内部记忆目录", re.compile(r"\.workbuddy[\\/]"),
                  "内部记忆不属于公开仓库；确认已 gitignore 且从未进过历史"))

    return rules


# 允许的通用系统路径：命中这些前缀的行不再算「盘符绝对路径」
def _allowed_line(line):
    return any(a in line for a in PATH_ALLOW)


BAD_PATH_PATTERNS = (
    (r"(^|/)\.workbuddy(/|$)", "内部记忆目录"),
    (r"使用说明\.md$", "面向本机的交付文档"),
    (r"\.(pfx|p12|cer|key|pem)$", "证书 / 私钥"),
    (r"(^|/)token\.txt$", "凭据缓存"),
    (r"(^|/)(dist|build)/", "构建产物"),
    (r"(^|/)\.pybuild_cache(/|$)", "本机构建缓存"),
    (r"(^|/)(archive|verify)/", "本机工作暂存"),
    (r"\.exe$", "二进制产物（应走 Release 资产）"),
)


def is_text_file(name):
    base = os.path.basename(name)
    if base in ("LICENSE", "Makefile", "Dockerfile", ".gitignore"):
        return True
    return os.path.splitext(name)[1].lower() in TEXT_EXTS


def scan_text(text, label, rules, findings):
    """扫一段文本，把命中写进 findings。"""
    base = os.path.basename(label.replace("\\", "/"))
    lines = text.splitlines()
    for name, pat, hint in rules:
        # .gitignore 里出现 .workbuddy/ 是**在排除它**，不是泄露，跳过这条规则
        if name == "内部记忆目录" and base in (".gitignore", ".gitattributes"):
            continue
        for i, line in enumerate(lines, 1):
            m = pat.search(line)
            if not m:
                continue
            if name == "盘符绝对路径" and _allowed_line(line):
                continue
            findings.append((label, i, name, hint))
            break  # 每个规则每个文件只报第一处，避免刷屏


def scan_files(files, rules, quiet=False):
    findings = []
    scanned = 0
    for rel in files:
        if rel in SELF_SKIP or os.path.basename(rel) in SELF_SKIP:
            continue
        if not is_text_file(rel):
            continue
        p = os.path.join(ROOT, rel.replace("/", os.sep))
        if not os.path.isfile(p) or os.path.getsize(p) > MAX_SIZE:
            continue
        try:
            with open(p, encoding="utf-8", errors="replace") as f:
                text = f.read()
        except OSError:
            continue
        if "\0" in text[:4096]:
            continue
        scanned += 1
        scan_text(text, rel, rules, findings)
    return scanned, findings


def report(scanned, findings):
    if findings:
        print(">> 命中 %d 处（扫描 %d 个文本文件）：\n" % (len(findings), scanned))
        for label, line, name, hint in findings:
            print("   [危险] %s:%d" % (label, line))
            print("          规则：%s" % name)
            print("          处理：%s" % hint)
        print()
        return False
    print(">> 干净：扫描 %d 个文本文件，未发现开发机信息 / 凭据。" % scanned)
    return True


def scan_history(rules):
    """扫全部 git 历史：先看曾经跟踪过哪些路径，再逐 blob 扫内容。"""
    print(">> 扫描 git 历史…")
    names = _git("log", "--all", "--pretty=format:", "--name-only", "--diff-filter=A")
    if names is None:
        print("   无法读取 git 历史（不是仓库或没有 git）")
        return True
    paths = sorted({l.strip() for l in names.splitlines() if l.strip()})
    bad = []
    for p in paths:
        for pat, why in BAD_PATH_PATTERNS:
            if re.search(pat, p, re.I):
                bad.append((p, why))
                break
    if bad:
        print("   !! 历史中出现过不该入库的路径：")
        for p, why in bad:
            print("      %s   （%s）" % (p, why))
        print("   → 这些内容在 GitHub 上仍可被翻到，必须重写历史或删库重建。")
    else:
        print("   历史路径检查：干净（%d 个曾出现的路径）" % len(paths))

    # 逐 blob 扫内容（用「首次出现的路径」作标签，便于定位与误报豁免）
    commits = _git("rev-list", "--all")
    if commits is None:
        return not bad
    hashes = {}
    for c in commits.split():
        tree = _git("ls-tree", "-r", "--format=%(objectname)\t%(path)", c)
        if not tree:
            continue
        for line in tree.splitlines():
            parts = line.split("\t", 1)
            if len(parts) == 2:
                hashes.setdefault(parts[0], parts[1])
    findings = []
    scanned = 0
    for h, path in sorted(hashes.items(), key=lambda kv: kv[1]):
        # 扫描脚本自身、钩子文件里必然出现规则关键字，跳过（否则自己报自己）
        if os.path.basename(path) in SELF_SKIP:
            continue
        try:
            p = subprocess.run([GIT, "cat-file", "blob", h],
                               cwd=ROOT, capture_output=True)
        except Exception:  # noqa: BLE001
            continue
        if p.returncode != 0:
            continue
        raw = p.stdout
        if b"\0" in raw[:4096] or len(raw) > MAX_SIZE:
            continue
        text = raw.decode("utf-8", errors="replace")
        scanned += 1
        scan_text(text, path, rules, findings)

    if findings:
        print("\n   !! 历史内容命中 %d 处：" % len(findings))
        for label, line, name, hint in findings:
            print("      历史版本 %s:%d  → %s" % (label, line, name))
        print("   → 历史一旦推送就无法靠 git rm 消除，需要重写历史 / 删库重建。")
    else:
        print("   历史内容检查：干净（%d 个 blob）" % scanned)
    ok = not bad and not findings
    print(">> 历史扫描结果：%s\n" % ("干净" if ok else "有问题，需处理"))
    return ok


def collect_files(mode):
    if mode == "all":
        out = []
        for dirpath, dirnames, filenames in os.walk(ROOT):
            dirnames[:] = [d for d in dirnames
                           if d not in (".git", "__pycache__", "dist", "build")]
            for fn in filenames:
                full = os.path.join(dirpath, fn)
                out.append(os.path.relpath(full, ROOT).replace(os.sep, "/"))
        return out
    if mode == "staged":
        out = _git("diff", "--cached", "--name-only", "--diff-filter=ACMR")
    else:
        out = _git("ls-files")
    if out is None:
        return []
    return [l.strip() for l in out.splitlines() if l.strip()]


def main():
    ap = argparse.ArgumentParser(description="入库前敏感信息扫描")
    ap.add_argument("--staged", action="store_true", help="只扫暂存区（pre-commit 用）")
    ap.add_argument("--all", action="store_true", help="连未入库文件一起扫")
    ap.add_argument("--history", action="store_true", help="扫全部 git 历史")
    args = ap.parse_args()

    rules = build_rules()

    if args.history:
        return 0 if scan_history(rules) else 1

    mode = "staged" if args.staged else ("all" if args.all else "tracked")
    label = {"staged": "暂存区", "all": "全部文件（含未入库）", "tracked": "已入库文件"}[mode]
    files = collect_files(mode)
    if not files:
        print(">> 没有可扫描的文件（%s）" % label)
        return 0

    print(">> 扫描%s：%d 个文件" % (label, len(files)))
    scanned, findings = scan_files(files, rules)
    ok = report(scanned, findings)
    if not ok:
        print(">> 提交已中止。修正上面各项后重试；")
        print("   确认是误报可用：git commit --no-verify")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
