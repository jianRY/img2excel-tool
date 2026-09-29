# -*- coding: utf-8 -*-
"""自动更新模块回归测试（不联网）

覆盖最容易被改坏的几条硬约束：
  ① 「镜像优先、自有站兜底」—— 自有站哪怕测速最快也必须排最后
  ② 版本比较补零到 3 段、且十位数不能按字符串比
  ③ 安装版保持固定名 / 绿色版带版本号
  ④ 安装包资产不能被当作可更新文件
  ⑤ update.json 与自有站元数据的地址推导

用法：python test_autoupdate.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import autoupdate as A  # noqa: E402

REPO_API = "https://api.github.com/repos/jianRY/img2excel-tool/releases/latest"
MIRROR = A.MIRROR_PREFIXES[0]
GH = "https://github.com/jianRY/img2excel-tool/releases/download/v1.2.1/Img2ExcelTool_v1.2.1_onefile.exe"
SITE = A.SITE_URL.rstrip("/") + "/files/Img2ExcelTool_v1.2.1_onefile.exe"

PASS = FAIL = 0


def chk(label, cond, extra=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print("  [PASS] %s" % label)
    else:
        FAIL += 1
        print("  [FAIL] %s   %s" % (label, extra))


def test_version():
    print("\n[1] 版本比较")
    chk("parse_version 补零到 3 段", A.parse_version("2.9") == (2, 9, 0), A.parse_version("2.9"))
    chk("parse_version 去掉 v 前缀", A.parse_version("v2.10.0") == (2, 10, 0))
    chk("1.2.1 > 1.2.0", A.version_greater("1.2.1", "1.2.0"))
    chk("2.10.0 > 2.9.0（十位数不能按字符串比）", A.version_greater("2.10.0", "2.9.0"))
    chk("2.9.0 不大于 2.9.0（防升级死循环）", not A.version_greater("2.9.0", "2.9.0"))
    chk("两段式 tag 与三位常量等价 v2.8.0 vs 2.8", not A.version_greater("2.8.0", "2.8"))


def test_src_rank():
    print("\n[2] 源分组")
    chk("镜像 → 0", A._src_rank(MIRROR + GH) == 0)
    chk("GitHub 原站 → 1", A._src_rank(GH) == 1)
    chk("自有站 → 2", A._src_rank(SITE) == 2)
    chk("镜像标签", A._src_label(MIRROR + GH).startswith("加速镜像"))
    chk("自有站标签", A._src_label(SITE) == "自有服务器")


def test_build_sources():
    print("\n[3] 下载候选展开")
    out = A.build_download_sources([GH, SITE])
    chk("每个镜像各派生一份", sum(1 for u in out if u.startswith(MIRROR)) == 1)
    chk("GitHub 原链保留", GH in out)
    chk("自有站保留", SITE in out)
    chk("镜像排在 GitHub 原站之前", out.index(MIRROR + GH) < out.index(GH))


def test_rank_keeps_site_last():
    """硬约束：自有站测速再快，也必须排最后。只按速度排就会破坏这条。"""
    print("\n[4] 排序：自有站永远兜底（关键回归）")
    orig = A.probe_speed
    speeds = {}

    def fake(url, nbytes=A.PROBE_BYTES, timeout=A.PROBE_TIMEOUT):
        return speeds.get(url, 0.0)

    A.probe_speed = fake
    try:
        m, g, s = MIRROR + GH, GH, SITE
        speeds[m] = 50.0        # 镜像很慢（远低于门限）
        speeds[g] = 30.0        # 原站更慢
        speeds[s] = 9999.0      # 自有站最快 —— 故意诱导「按速度排」
        out = A.rank_sources([s, g, m])
        chk("自有站排最后", out[-1] == s, out)
        chk("镜像排最前", out[0] == m, out)

        # 镜像全挂（探测 0）也不能被丢弃
        speeds2 = {m: 0.0, g: 0.0, s: 120.0}
        speeds.clear()
        speeds.update(speeds2)
        out2 = A.rank_sources([m, g, s])
        chk("探测失败的源不被丢弃", len(out2) == 3, out2)
        chk("镜像挂了仍是组优先（排原站前）", out2.index(m) < out2.index(g), out2)
        chk("自有站依旧兜底", out2[-1] == s, out2)
    finally:
        A.probe_speed = orig


def test_installer_detect():
    print("\n[5] 安装包资产识别")
    for n in ("Img2ExcelTool_v1.2.1_setup.exe", "App_installer.exe", "图片助手_安装版.exe"):
        chk("跳过 %s" % n, A._looks_like_installer(n))
    for n in ("Img2ExcelTool_v1.2.1_onefile.exe", "图片表格转Excel助手_v1.2.1.exe"):
        chk("消费 %s" % n, not A._looks_like_installer(n))


def test_naming():
    print("\n[6] 更新后的文件名")
    app = "图片表格转Excel助手"
    chk("安装版（光名字）保持不变",
        A.desired_base_name(app, "1.2.1", app) == app)
    chk("绿色版带上新版本号",
        A.desired_base_name(app, "1.2.1", app + "_v1.2.0") == app + "_v1.2.1")
    chk("未接 app_name 时原样返回（保护其他项目）",
        A.desired_base_name(None, None, "随便") == "随便")
    chk("_base_stem 去 _更新中", A._base_stem(app + "_更新中") == app)
    chk("_base_stem 去 _更新中_序号", A._base_stem(app + "_更新中_123") == app)
    chk("_base_stem 去 _旧版_时间戳", A._base_stem(app + "_旧版_20260927") == app)


def test_urls():
    print("\n[7] 元数据地址推导")
    chk("Release 里的 update.json 地址",
        A._release_meta_url(REPO_API)
        == "https://github.com/jianRY/img2excel-tool/releases/latest/download/update.json")
    chk("自有站元数据地址",
        A._server_meta_url(REPO_API) == A.SITE_URL.rstrip("/") + "/updates/img2excel.json")
    chk("_REPO_TO_APP 短名与发版脚本一致", A._REPO_TO_APP.get("jianry/img2excel-tool") == "img2excel")
    chk("认不出的仓库返回 None（自动跳过该级）", A._server_meta_url("https://x/y") is None)


def test_update_json_contract():
    """发版脚本生成的 update.json 必须与客户端读的键一致。"""
    print("\n[8] update.json 字段契约（发版侧 ↔ 客户端）")
    import release as R
    import json as _json
    import tempfile

    d = tempfile.mkdtemp()
    exe = os.path.join(d, "a.exe")
    with open(exe, "wb") as f:
        f.write(b"x" * 2048)
    _, data = R.make_update_json("1.2.1", exe, None, "测试")

    need = ("version", "url", "fallback_url", "size", "sha256", "notes")
    for k in need:
        chk("含字段 %s" % k, k in data)
    chk("app 短名三处一致", data["app"] == A._REPO_TO_APP["jianry/img2excel-tool"])
    chk("url 是 GitHub 直链（主源）", "github.com" in data["url"])
    chk("fallback_url 是自有站（兜底）", data["fallback_url"].startswith(A.SITE_URL))
    chk("源语义没接反", "github.com" not in data["fallback_url"])
    chk("sha256 与文件一致", data["sha256"] == R._sha256_of(exe))
    chk("落盘可解析", _json.load(open(data.get("_path", os.path.join(R.CACHE, "update.json")),
                                    encoding="utf-8"))["version"] == "1.2.1")


if __name__ == "__main__":
    print("=" * 60)
    print("自动更新模块回归测试")
    print("=" * 60)
    test_version()
    test_src_rank()
    test_build_sources()
    test_rank_keeps_site_last()
    test_installer_detect()
    test_naming()
    test_urls()
    test_update_json_contract()
    print("\n" + "=" * 60)
    print("结果：%d 通过 / %d 失败" % (PASS, FAIL))
    print("=" * 60)
    raise SystemExit(1 if FAIL else 0)
