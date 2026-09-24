# img2excel 项目长期约定

## 环境

- 工作目录：`D:\workbuddy\杂项\img22excel`（**已从 C 盘 WorkBuddy 目录迁来**，文档里的旧 C 盘路径已失效）
- 可用 Python 环境：`C:\Users\toxuj\.workbuddy\binaries\python\envs\court_build_v13\Scripts\python.exe`
  - 该环境依赖齐全（rapidocr-onnxruntime / rapid-table / onnxruntime / opencv / PyMuPDF / openpyxl / PyInstaller）
  - 原专用 venv `envs\img2excel` 已不存在，勿引用
- 测试样本：`C:\Users\toxuj\Desktop\FAPIAO\`（66 个医疗票据实拍图）

## 技术约定

- **识别管线**：RapidOCR 文字识别 → RapidTable(SLANet-Plus) 表格结构 → 退化则走 `fallback_lines` 兜底 → 数值清洗 → Excel
- **绝不允许脏值进 Excel**：空单元格必须是真空值，不得出现字面 `'None'` 字符串
- **模糊数值不猜测**：`12.345` 这类无法判定是小数还是被误插小数点的，一律保持原文本输出
- **金额修正必须在「原文」列留痕**：任何自动修正都要可追溯、可核对

## 已知陷阱（血泪教训）

1. **不要尝试自动旋转纠偏**——实测原图方向永远最优，旋转只会劣化（详见 2026-09-24 日志）
2. **`(?<=\d)\s+(?=\d)` lookbehind 形式在本环境不生效**，用 `\s+` 全清
3. **SLANet 失败时静默**——它会把整页塞进 1 个单元格且不报错，必须用 `table_quality_ok` 拦截
4. **多选删除不要用 `tree.index()`**——索引会随删除位移，导致删错文件
5. **不要用 PowerShell 工具跑 `Get-AuthenticodeSignature`**（沙箱报 decisionRecord 错），
   改用签名目录里的 `bin/osslsigncode.exe verify -in <exe>`
6. **发版脚本签名「卡死」= 代理污染，不是网络/文件大小**：
   沙箱注入的 `http_proxy/https_proxy` 指向本地代理（**端口随机，别写死**），
   osslsigncode 走它会连不上 TSA 且**每次等完整 TCP 超时**，5 服务器×3 轮拖死外层超时。
   正解：`release.py::build_env()` 剥代理 + `sign.py` 内单服务器 `timeout=25` 硬超时。
   诊断捷径：手工直跑 `sign.py`，秒过即证明是调用方 env 被污染。详见技能 `exe-self-sign`。
7. **Inno Setup 7 命令行语法**：用 `-d<name>=<value>` / `-o<path>`，旧的 `/D` `/O` 会报
   「You may not specify more than one script filename」。


## 发布约定

- 所有打包的 exe **必须**用 jianRY 自签名证书签名（见技能 `exe-self-sign`）
- 签名工具：`D:\workbuddy\诉讼案件网站\.pybuild_cache\signing\sign.py`
- 打包命令：`python -m PyInstaller "图片表格转Excel助手.spec" --noconfirm --clean`
- 发版前跑三连自检：单图 / PDF / 文件夹批量，并复核 `.selftest.log`
