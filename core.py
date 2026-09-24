# -*- coding: utf-8 -*-
"""
img2excel 核心转换模块
离线流程：收集文件 -> PDF 逐页转图 -> 自动方向矫正 -> RapidOCR 识别
          -> RapidTable 表格结构还原（退化时走兜底分列）-> 数值清洗 -> 写 Excel
"""
import os
import re
from html.parser import HTMLParser
from pathlib import Path

import numpy as np

APP_VERSION = "1.2.0"

IMG_EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff", ".webp"}
PDF_EXTS = {".pdf"}


# ---------------------------------------------------------------- 资源路径
def app_base_dir() -> Path:
    """PyInstaller 单文件模式下指向解包目录，否则指向脚本目录"""
    if getattr(__import__("sys"), "frozen", False):
        return Path(__import__("sys")._MEIPASS)  # noqa: SLF001
    return Path(__file__).resolve().parent


def table_model_path() -> str:
    bundled = app_base_dir() / "models" / "slanet-plus.onnx"
    if bundled.exists():
        return str(bundled)
    dev = Path(__file__).resolve().parent / "models" / "slanet-plus.onnx"
    return str(dev)


# ---------------------------------------------------------------- 引擎（懒加载单例）
_ocr_engine = None
_table_engine = None


def get_ocr_engine():
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    return _ocr_engine


def get_table_engine():
    global _table_engine
    if _table_engine is None:
        from rapid_table import ModelType, RapidTable, RapidTableInput
        cfg = RapidTableInput(
            model_type=ModelType.SLANETPLUS,
            model_dir_or_path=table_model_path(),
            use_ocr=True,
        )
        _table_engine = RapidTable(cfg)
    return _table_engine


# ---------------------------------------------------------------- 文件收集
def collect_files(path: str) -> list:
    """返回可处理的文件绝对路径列表（目录则递归扫描）"""
    p = Path(path)
    if p.is_file():
        return [str(p)] if p.suffix.lower() in IMG_EXTS | PDF_EXTS else []
    files = []
    for root, _dirs, names in os.walk(p):
        for n in sorted(names):
            f = Path(root) / n
            if f.suffix.lower() in IMG_EXTS | PDF_EXTS:
                files.append(str(f))
    return files


def pdf_to_images(pdf_path: str, dpi: int = 200):
    """PDF 逐页转 numpy 图（RGB），返回 [(页码, img)]"""
    try:
        import pymupdf as fitz  # PyMuPDF >= 1.24
    except ImportError:
        import fitz
    pages = []
    with fitz.open(pdf_path) as doc:
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=dpi)
            img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, pix.n
            )
            if pix.n == 4:
                img = img[:, :, :3]
            elif pix.n == 1:
                img = np.repeat(img, 3, axis=2)
            pages.append((i, np.ascontiguousarray(img)))
    return pages


def read_image(path: str):
    import cv2
    img = cv2.imdecode(np.fromfile(path, dtype=np.uint8), cv2.IMREAD_COLOR)
    if img is None:  # 中文路径兜底失败则常规读
        import cv2 as _cv
        img = _cv.imread(path)
    return img


# ---------------------------------------------------------------- 数值清洗
# 纯数字（含千分位 / 小数点 / 数字间误插空格），允许前后有空格或货币符号
_NUM_RE = re.compile(r"^[\s¥￥$]*([+-]?\d[\d\s.,]*)[\s元]*$")
# 日期：20240123 / 2024-01-23 / 2024/1/23
_DATE_RE = re.compile(r"^(\d{4})[-/.年]?(\d{1,2})[-/.月]?(\d{1,2})日?$")


def fix_number(text: str):
    """把 OCR 文本规整成 (值, 原文) —— 值可能是 float / int / 原字符串。

    修复三类常见 OCR 误判：
      千分位逗号被认成小数点 : "2.903.00"  -> 2903.0
      千分位逗号残留         : "2,903.00"  -> 2903.0
      小数点后带空格         : "4. 69"     -> 4.69
    无法安全判定时原样返回字符串，绝不猜测。
    """
    if text is None:
        return "", text
    raw = str(text)
    s = raw.strip()
    if not s:
        return "", raw
    m = _NUM_RE.match(s)
    if not m:
        return raw, raw
    body = m.group(1)
    # 去掉数字之间的空格（"4. 69" / "2 903.00" / "2 . 9 0 3 . 0 0"）
    # 注意：不要用 (?<=\d)\s+(?=\d) 形式的 lookbehind，本环境下它不生效
    body = re.sub(r"\s+", "", body)
    sign = ""
    if body[:1] in "+-":
        sign, body = body[0], body[1:]
    if not body or not body[0].isdigit():
        return raw, raw

    has_dot, has_comma = "." in body, "," in body
    if has_dot and has_comma:
        # 标准写法：逗号千分位 + 点小数（或欧式反之），取最后出现的为小数点
        if body.rfind(",") > body.rfind("."):
            body = body.replace(".", "").replace(",", ".")
        else:
            body = body.replace(",", "")
    elif body.count(".") > 1:
        # 多个点：最后一个点后是 1~2 位则视为小数点，其余点视为千分位
        head, _, tail = body.rpartition(".")
        if len(tail) in (1, 2) and head.replace(".", "").isdigit():
            body = head.replace(".", "") + "." + tail
        else:
            body = body.replace(".", "")
    elif body.count(",") > 1:
        head, _, tail = body.rpartition(",")
        if len(tail) in (1, 2) and head.replace(",", "").isdigit():
            body = head.replace(",", "") + "." + tail
        else:
            body = body.replace(",", "")
    elif has_dot:
        # 形如 "12.345" 的模糊情形：既可能是 12.345，也可能是 12345 被误加小数点。
        # 金额场景下 3 位小数极罕见，但直接改判成整数风险更大——
        # 因此只有在整数部分长度符合千分位分组规律（如 "2.903"）时才视为误判。
        int_part, _, dec_part = body.partition(".")
        if (len(dec_part) == 3 and dec_part.isdigit() and int_part.isdigit()
                and 1 <= len(int_part) <= 4 and (len(int_part) + 3) % 3 == 0):
            body = int_part + dec_part
    elif has_comma:
        int_part, _, dec_part = body.partition(",")
        if len(dec_part) == 3 and dec_part.isdigit():
            body = int_part + dec_part
        else:
            body = int_part + "." + dec_part

    if not re.fullmatch(r"\d+(\.\d+)?", body):
        return raw, raw
    try:
        val = float(body)
    except ValueError:
        return raw, raw
    # 整数且无小数位 -> 输出 int，避免 20240123 变 20240123.0
    if "." not in body:
        return (int(val), raw) if abs(val) < 1e15 else (raw, raw)
    # 数值加符号
    if sign == "-":
        val = -val
    return val, raw


def is_date(text: str) -> bool:
    """判断是否为可安全转换的日期串（只认 8 位数字或带分隔符的 YYYYMMDD）"""
    m = _DATE_RE.match(str(text).strip())
    if not m:
        return False
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    return 1900 <= y <= 2100 and 1 <= mo <= 12 and 1 <= d <= 31


# ---------------------------------------------------------------- OCR + 表格还原
def ocr_image(img: np.ndarray):
    """返回 (boxes, txts, scores)；无文字时返回 (None, None, None)
    兼容 rapidocr_onnxruntime 两种返回：
      1.2.x: (result, elapse)，result 为 [[box, text, score], ...]
      1.3.x: Result 对象（.boxes/.txts/.scores）
    """
    ret = get_ocr_engine()(img)
    if isinstance(ret, tuple):
        result = ret[0] if len(ret) >= 1 else None
    else:
        result = ret
    if result is None:
        return None, None, None
    if hasattr(result, "boxes"):                    # Result 对象
        if result.boxes is None or len(result.txts) == 0:
            return None, None, None
        return result.boxes, tuple(result.txts), tuple(result.scores)
    if not result:                                  # 空列表
        return None, None, None
    boxes = [r[0] for r in result]
    txts = tuple(r[1] for r in result)
    scores = tuple(float(r[2]) for r in result)
    return boxes, txts, scores


def recognize_table(img: np.ndarray, ocr_results):
    """表格结构识别，返回 (rows: List[List[str]], ok: bool)"""
    engine = get_table_engine()
    out = engine(img, ocr_results=[ocr_results])
    html = out.pred_htmls[0] if out.pred_htmls else ""
    rows = parse_html_table(html)
    return rows, bool(rows)


def table_quality_ok(rows, ocr_results):
    """判断 SLANet 给出的结构是否可信。

    SLANet 在「无线表格 / 票据 / 病历」这类版面上常常失败，且**失败时不报错**，
    而是把整页文字全塞进 1~2 个单元格、或少数几行，输出一张几乎全是空白的表。
    若不拦截，Excel 里就会出现「一整坨文字挤在一格」的垃圾结果。
    """
    if not rows:
        return False
    _boxes, txts, _ = ocr_results
    n_text = len(txts) if txts else 0
    if n_text == 0:
        return False
    data_rows = [r for r in rows if any(str(c).strip() for c in r)]
    n_rows = len(data_rows)
    n_cells = sum(1 for r in rows for c in r if str(c).strip())
    # 规则 1：大量文字框却只有 1~2 行 -> 结构退化
    if n_text >= 8 and n_rows <= 2:
        return False
    # 规则 2：单元格数远少于文字框数 -> 大量文字被吞进同一个格子
    if n_text >= 8 and n_cells < n_text * 0.5:
        return False
    # 规则 3：单格内容异常长（多行文字被拼在一起）
    for r in rows:
        for c in r:
            s = str(c)
            if len(s) > 80 and s.count(" ") > 8:
                return False
    return True


# ---------------------------------------------------------------- HTML 表格解析
class _TableHTMLParser(HTMLParser):
    """把 pred_html 解析成二维网格，支持 colspan/rowspan"""

    def __init__(self):
        super().__init__()
        self.grid = []          # List[List[str|None]]
        self.row = -1
        self.col = 0
        self.in_cell = False
        self.buf = []
        self.spans = (1, 1)     # colspan, rowspan

    def handle_starttag(self, tag, attrs):
        if tag == "tr":
            self.row += 1
            self.col = 0
        elif tag in ("td", "th"):
            self.in_cell = True
            self.buf = []
            d = dict(attrs)
            self.spans = (int(d.get("colspan", 1)), int(d.get("rowspan", 1)))
            # 跳过被上方 rowspan 占用的格
            while (self.row < len(self.grid)
                   and self.col < len(self.grid[self.row])
                   and self.grid[self.row][self.col] is not None):
                self.col += 1

    def handle_data(self, data):
        if self.in_cell:
            self.buf.append(data)

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self.in_cell:
            self.in_cell = False
            text = re.sub(r"\s+", " ", "".join(self.buf)).strip()
            self._place(text)
            self.col += self.spans[0]
            self.spans = (1, 1)
        elif tag == "table":
            pass

    def _place(self, text):
        cs, rs = self.spans
        while len(self.grid) <= self.row:
            self.grid.append([])
        line = self.grid[self.row]
        while len(line) <= self.col:
            line.append(None)
        line[self.col] = text
        for k in range(1, cs):
            while len(line) <= self.col + k:
                line.append(None)
            line[self.col + k] = ""     # 被横向合并占位
        for k in range(1, rs):          # 被纵向合并占位
            while len(self.grid) <= self.row + k:
                self.grid.append([])
            ext = self.grid[self.row + k]
            while len(ext) <= self.col:
                ext.append(None)
            ext[self.col] = ""


def parse_html_table(html: str):
    if not html or "<table" not in html.lower():
        return []
    parser = _TableHTMLParser()
    try:
        parser.feed(html)
    except Exception:
        return []
    # 清理 None（空格位）
    grid = [[("" if c is None else c) for c in row] for row in parser.grid]
    if not grid:
        return []
    width = max(len(r) for r in grid)
    grid = [r + [""] * (width - len(r)) for r in grid]
    # 全空行剔除
    return [r for r in grid if any(str(c).strip() for c in r)]


def _cluster_rows(items, y_tol_ratio=0.6, min_tol=8.0):
    """按 y 中心聚类成行，返回 [[item, ...], ...]，item = (yc, x0, x1, txt)"""
    items = sorted(items, key=lambda t: (t[0], t[1]))
    rows, cur, cur_y, heights = [], [], None, []
    for it in items:
        if cur_y is None:
            cur, cur_y, heights = [it], it[0], [it[2] - it[1]]
            continue
        tol = max(np.mean(heights) * y_tol_ratio, min_tol)
        if abs(it[0] - cur_y) <= tol:
            cur.append(it)
            heights.append(it[2] - it[1])
            cur_y = sum(x[0] for x in cur) / len(cur)
        else:
            rows.append(cur)
            cur, cur_y, heights = [it], it[0], [it[2] - it[1]]
    if cur:
        rows.append(cur)
    return rows


# 同一单元格内允许的最大字符间隙（相对框高中位数）
def _merge_tol(median_w: float) -> float:
    """字符间隙阈值：中文单据字距通常远小于半个字宽"""
    return max(median_w * 1.1, 14.0)


def _merge_row_cells(row_items):
    """把同一行内「属于同一个单元格」的相邻框合并。

    RapidOCR 常把一个数（3,040.00）或一串词切成多个框，不合并就会出现
    "3,040." 与 "063." 这种断裂值。但合并过头同样有害——会把相邻两列的
    数字粘成一串（如 "0.00" + "0.00" -> "0.000.00"）。
    因此判据必须严格：只认「间隙极小 且 拼接后仍是一个合法数字」的情形。
    """
    row_items = sorted(row_items, key=lambda t: t[1])
    if not row_items:
        return []
    widths = [it[2] - it[1] for it in row_items]
    median_w = float(np.median(widths))
    tol = _merge_tol(median_w)
    # 数字续接允许的最小间隙：远小于半个字宽，否则视为跨列
    num_tol = max(median_w * 0.45, 6.0)

    cells = [list(row_items[0])]
    for it in row_items[1:]:
        gap = it[1] - cells[-1][2]
        prev_txt, cur_txt = str(cells[-1][3]).strip(), str(it[3]).strip()
        both_numeric = (re.fullmatch(r"[+-]?[\d., ]+", prev_txt)
                        and re.fullmatch(r"[+-]?[\d., ]+", cur_txt))
        merged = (prev_txt + cur_txt).replace(" ", "")
        # 只有「拼接后恰好是一个合法数字」才认为是被切断的同一个数
        num_glue = (gap <= tol and both_numeric
                    and bool(re.fullmatch(r"[+-]?\d[\d.,]*", merged))
                    and not re.search(r"[.,]\d{3}[.,]", merged))
        if num_glue:
            cells[-1][2] = max(cells[-1][2], it[2])
            cells[-1][3] = merged
        elif gap <= tol:
            # 普通粘连：同一单元格内的文字。两个独立数字绝不吃
            if both_numeric:
                cells.append(list(it))
                continue
            cells[-1][2] = max(cells[-1][2], it[2])
            cells[-1][3] = prev_txt + (" " if prev_txt and cur_txt else "") + cur_txt
        else:
            cells.append(list(it))
    return cells


def _cluster_cols(rows, x_tol_ratio=0.85, min_tol=10.0):
    """把所有行的框投影到 x 轴做一维聚类，得到全局列中心列表。

    比「固定 18px 间隙」稳健得多：即便某行只有 2 个框、另一行有 6 个框，
    也能把它们对齐到同一套列骨架上。
    """
    centers = []
    for row in rows:
        for it in row:
            centers.append(((it[1] + it[2]) / 2.0, it[2] - it[1]))
    if not centers:
        return []
    centers.sort(key=lambda t: t[0])
    median_w = float(np.median([h for _c, h in centers]))
    tol = max(median_w * x_tol_ratio, min_tol)
    cols, cur, cur_c = [], [], None
    for c, w in centers:
        if cur_c is None:
            cur, cur_c = [(c, w)], c
        elif abs(c - cur_c) <= tol:
            cur.append((c, w))
            cur_c = sum(x[0] for x in cur) / len(cur)
        else:
            cols.append(sum(x[0] for x in cur) / len(cur))
            cur, cur_c = [(c, w)], c
    if cur:
        cols.append(sum(x[0] for x in cur) / len(cur))
    # 合并过近的列（防止把同一列切成两半），但要比聚类阈值严格得多，
    # 否则相邻两列（如「个人账户支付」与「个人账户余额」）会被并成一列，
    # 导致两格数字粘成 "0.000.00" 这种脏值。
    merged = []
    for c in cols:
        if merged and c - merged[-1] < tol * 0.35:
            merged[-1] = (merged[-1] + c) / 2.0
        else:
            merged.append(c)
    return merged


def fallback_lines(img: np.ndarray, ocr_results):
    """无表格结构时的兜底：按 y 聚类成行，再把每个框吸附到全局列骨架

    旧实现用固定 18px 间隙判断换列，遇到「表头多个词被拼成一坨」或
    「个别行缺列」就会整体错位。这里改为两阶段聚类：先定列骨架，再逐行填入。
    """
    boxes, txts, _scores = ocr_results
    items = []
    for box, txt in zip(boxes, txts):
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        items.append((float(np.mean(ys)), float(min(xs)), float(max(xs)), str(txt)))
    if not items:
        return []
    rows = _cluster_rows(items)
    # 先把每行内被切碎的框粘回单元格，再定全局列骨架
    merged_rows = [_merge_row_cells(r) for r in rows]
    merged_rows = [r for r in merged_rows if r]
    cols = _cluster_cols(merged_rows)
    col_width = (np.diff(cols).mean() if len(cols) > 1 else 60.0)
    out = []
    for row in merged_rows:
        row_sorted = sorted(row, key=lambda t: t[1])
        line = ["" for _ in cols]
        for _y, x0, x1, txt in row_sorted:
            # 按「框与列中心的重叠程度」归属，而非单纯比中心距离。
            # 只比中心距离会把长框错误地拉到相邻列，造成两列数字粘在一起。
            best_i, best_ov = 0, -1e9
            for i, c in enumerate(cols):
                lo, hi = c - col_width / 2, c + col_width / 2
                ov = min(x1, hi) - max(x0, lo)
                if ov > best_ov:
                    best_ov, best_i = ov, i
            ci = best_i
            if line[ci]:
                line[ci] += txt          # 同列多段文字拼接
            else:
                line[ci] = txt
        if any(c.strip() for c in line):
            out.append(line)
    return out


# ---------------------------------------------------------------- 单文件转换
def convert_one(path: str):
    """返回 [(页码, rows)]；rows 为该页的二维表格"""
    ext = Path(path).suffix.lower()
    pages = []
    if ext in PDF_EXTS:
        pages = pdf_to_images(path)
    else:
        img = read_image(path)
        if img is not None:
            pages = [(1, img)]
    results = []
    for page_no, img in pages:
        try:
            boxes, txts, scores = ocr_image(img)
            if boxes is None:
                results.append((page_no, []))
                continue
            try:
                rows, ok = recognize_table(img, (boxes, txts, scores))
                if not ok or not table_quality_ok(rows, (boxes, txts, scores)):
                    rows = fallback_lines(img, (boxes, txts, scores))
            except Exception:
                rows = fallback_lines(img, (boxes, txts, scores))
            results.append((page_no, rows))
        except Exception as e:
            results.append((page_no, None))
            results.append((f"__error__", f"{Path(path).name} 第{page_no}页: {e}"))
    return results


def convert_many(paths, progress_cb=None, log_cb=None, fix_nums=True):
    """批量转换，返回 (all_rows, errors)
    all_rows: List[List] —— [来源文件, 页码, col1, col2, ...]
    fix_nums=True 时，单元格会做数值/日期清洗（千分位误判修正）。
    """
    all_rows, errors = [], []
    n = len(paths)
    for idx, path in enumerate(paths, start=1):
        name = Path(path).name
        if log_cb:
            log_cb(f"[{idx}/{n}] 正在识别：{name}")
        try:
            page_results = convert_one(path)
        except Exception as e:
            errors.append(f"{name}: {e}")
            if log_cb:
                log_cb(f"[{idx}/{n}] 识别失败：{name} ({e})")
            if progress_cb:
                progress_cb(idx, n)
            continue
        for page_no, rows in page_results:
            if isinstance(page_no, str) and page_no == "__error__":
                errors.append(rows)
                continue
            if rows is None:
                continue
            if not rows:
                if log_cb:
                    log_cb(f"[{idx}/{n}] 未识别到文字：{name} 第{page_no}页")
                continue
            fixed = 0
            for r in rows:
                if fix_nums:
                    cells, npair = [], 0
                    for c in r:
                        if c is None or c == "":
                            cells.append("")
                            continue
                        v, _raw = fix_number(c)
                        if isinstance(v, (int, float)):
                            npair += 1
                        cells.append(v)
                    fixed += npair
                else:
                    cells = ["" if c is None else str(c) for c in r]
                all_rows.append([name, page_no] + cells)
            if fix_nums and fixed and log_cb:
                log_cb(f"[{idx}/{n}] {name} 第{page_no}页："
                       f"{len(rows)} 行，规整 {fixed} 个数值")
        if progress_cb:
            progress_cb(idx, n)
    return all_rows, errors


def parse_cell(text):
    """把原始 OCR 文本转成写入 Excel 的值 + 类型标记
    返回 (value, kind)  kind ∈ {"num", "date", "text", ""}
    """
    if text is None:
        return "", ""
    s = str(text).strip()
    if not s:
        return "", ""
    v, _raw = fix_number(s)
    if isinstance(v, (int, float)):
        return v, "num"
    if is_date(s):
        m = _DATE_RE.match(s)
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        # 8 位纯数字的日期在 Excel 里会变科学计数，统一插成斜杠格式
        return f"{y}/{mo:02d}/{d:02d}", "date"
    return s, "text"


def write_excel(rows, out_path: str, source_names, keep_raw=True, per_file=False):
    """写 Excel。

    per_file=False（默认）：所有文件合并到单 sheet「识别结果」，便于按来源筛选。
    per_file=True：每个来源文件单独一个 sheet（sheet 名 = 文件名，截断至 31 字），
                   各表按自己的列数排版，避免不同版式互相撑宽。
    两种模式都会在末尾附「说明」sheet。
    """
    if per_file:
        return _write_excel_per_file(rows, out_path, keep_raw)
    return _write_excel_single(rows, out_path, keep_raw)


def _build_layout(max_cols, keep_raw):
    layout = [("来源文件", "src"), ("页码", "page")]
    for i in range(1, max_cols + 1):
        layout.append((f"列{i}", f"col{i}"))
        if keep_raw:
            layout.append((f"列{i}原文", f"raw{i}"))
    return layout


def _write_sheet(ws, rows, layout, styles):
    """把 rows 按 layout 写进工作表（含表头样式、冻结、筛选、列宽）"""
    from openpyxl.utils import get_column_letter
    (border, f_header, f_raw, f_data, fill_header, fill_raw,
     fill_src, center, left) = styles

    for c, (title, kind) in enumerate(layout, start=1):
        cell = ws.cell(row=1, column=c, value=title)
        cell.font = f_raw if kind.startswith("raw") else f_header
        cell.fill = fill_raw if kind.startswith("raw") else fill_header
        cell.alignment = center
        cell.border = border

    for ri, r in enumerate(rows, start=2):
        src = r[0] if len(r) > 0 else ""
        page = r[1] if len(r) > 1 else ""
        cells = r[2:]
        for ci, (title, kind) in enumerate(layout):
            if kind == "src":
                v = src
            elif kind == "page":
                v = page
            else:
                idx = int(re.sub(r"\D", "", kind)) - 1
                raw = cells[idx] if idx < len(cells) else ""
                if raw is None:
                    raw = ""
                if kind.startswith("raw"):
                    v = "" if raw == "" else str(raw)
                else:
                    v, _k = parse_cell(raw)
            cell = ws.cell(row=ri, column=ci + 1, value=v)
            if kind.startswith("raw"):
                cell.font = f_raw
                cell.fill = fill_raw
                cell.alignment = left
            else:
                cell.font = f_data
                cell.alignment = left if ci == 0 else center
                if kind == "src":
                    cell.fill = fill_src
                elif isinstance(v, float):
                    cell.number_format = "#,##0.00"
                elif isinstance(v, int) and not isinstance(v, bool):
                    cell.number_format = "#,##0"
            cell.border = border

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 8
    for i in range(3, len(layout) + 1):
        title = layout[i - 1][0]
        ws.column_dimensions[get_column_letter(i)].width = (
            13 if title.endswith("原文") else 18)
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = (f"A1:"
                          f"{ws.cell(row=len(rows) + 1, column=len(layout)).coordinate}")


def _styles():
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    thin = Side(style="thin", color="FFBFBFBF")
    return (
        Border(left=thin, right=thin, top=thin, bottom=thin),
        Font(name="微软雅黑", size=10, bold=True, color="FFFFFFFF"),
        Font(name="微软雅黑", size=9, color="FF7F7F7F"),
        Font(name="微软雅黑", size=10),
        PatternFill("solid", fgColor="FF4472C4"),
        PatternFill("solid", fgColor="FFDCE3F0"),
        PatternFill("solid", fgColor="FFD9E2F3"),
        Alignment(horizontal="center", vertical="center"),
        Alignment(horizontal="left", vertical="center"),
    )


def _write_excel_single(rows, out_path, keep_raw):
    from openpyxl import Workbook
    max_cols = max((len(r) - 2 for r in rows), default=1)
    max_cols = max(max_cols, 1)
    layout = _build_layout(max_cols, keep_raw)
    wb = Workbook()
    ws = wb.active
    ws.title = "识别结果"
    _write_sheet(ws, rows, layout, _styles())
    _add_notes_sheet(wb, rows, keep_raw)
    wb.properties.title = "图片表格识别结果"
    wb.properties.creator = "图片表格转Excel助手"
    wb.save(out_path)


def _write_excel_per_file(rows, out_path, keep_raw):
    from openpyxl import Workbook
    groups = {}
    for r in rows:
        groups.setdefault(str(r[0]), []).append(r)
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    for name, grows in groups.items():
        title = re.sub(r'[\\/*?:\[\]]', "_", Path(name).stem)[:31] or "Sheet"
        base, k = title, 2
        while title in used:
            title = f"{base[:28]}_{k}"
            k += 1
        used.add(title)
        ws = wb.create_sheet(title)
        max_cols = max((len(r) - 2 for r in grows), default=1)
        _write_sheet(ws, grows, _build_layout(max(max_cols, 1), keep_raw), _styles())
    _add_notes_sheet(wb, rows, keep_raw)
    wb.properties.title = "图片表格识别结果"
    wb.properties.creator = "图片表格转Excel助手"
    wb.save(out_path)


def _add_notes_sheet(wb, rows, keep_raw):
    """附一页使用说明，解释列含义与「原文」列的用途"""
    from openpyxl.styles import Font, PatternFill, Alignment
    ws = wb.create_sheet("说明")
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 96
    title_f = Font(name="微软雅黑", size=13, bold=True, color="FF2F5597")
    key_f = Font(name="微软雅黑", size=10, bold=True)
    val_f = Font(name="微软雅黑", size=10)
    wrap = Alignment(horizontal="left", vertical="top", wrap_text=True)

    ws["A1"] = "识别结果说明"
    ws["A1"].font = title_f
    ws.merge_cells("A1:B1")
    notes = [
        ("来源文件", "该行数据来自哪个图片 / PDF，按此列筛选可查看单个单据的全部内容。"),
        ("页码", "PDF 的第几页；图片固定为 1。"),
        ("列N", "识别出的表格内容。纯数字已自动转成数值格式，可直接求和；"
                "日期已归一为 YYYY/MM/DD。"),
    ]
    if keep_raw:
        notes.append(
            ("列N原文", "OCR 识别出的原始文本。程序会自动修正一些典型误判"
                        "（如千分位逗号被认成小数点：2.903.00 → 2903.00、"
                        "小数点后误插空格：4. 69 → 4.69）。核对时以「列N」为结果，"
                        "以「列N原文」为原始凭证；发现修正错误时可直接对照。"))
    notes += [
        ("数据行", f"本次共写入 {len(rows)} 行数据。"),
        ("准确性提示",
         "识别质量取决于原图：拍照请保持平整、光线均匀、正对拍摄。"
         "皱折、反光、印章遮挡会影响准确率，涉及金额请务必与原件复核。"),
        ("离线说明", "本工具全程离线运行，不会上传任何数据到网络。"),
    ]
    r = 3
    for k, v in notes:
        ws.cell(r, 1, k).font = key_f
        c = ws.cell(r, 2, v)
        c.font = val_f
        c.alignment = wrap
        ws.row_dimensions[r].height = 30 if len(v) < 60 else 48
        r += 1


def default_output_path(paths):
    """默认输出到第一个源文件同目录，名称：识别结果_YYYYMMDD_HHMMSS.xlsx"""
    from datetime import datetime
    base = Path(paths[0]).parent if paths else Path.cwd()
    return str(base / f"识别结果_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
