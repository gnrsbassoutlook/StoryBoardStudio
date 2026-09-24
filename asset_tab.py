# -*- coding: utf-8 -*-
"""StoryBoardStudio · 第一部分：影视资产表（X / Y / Z）工作台

设计原则
--------
1. **不侵入第二部分**。本模块自带索引、渲染、样式、脚本与保存接口，只通过
   StoryBoardStudio.py 里的四处增量挂载生效：导入 → CUSTOM_CSS/JS 追加 →
   gr.Tabs() 里调用 build_asset_tab() → 注册 /api/save_excel_a。
2. **右端固定预览列**。表格最右侧固定「📷 图片预览」「🔊 声音预览」两列，用 CSS
   position:sticky 吸附在视口右侧 —— 横向拖动表格时始终可见。这两列**不参与
   --a-scale 缩放**（宽度恒定），否则缩到 0.8 倍以下时 <audio> 的原生播放条会被
   挤没、缩略图也会小到看不清。图片预览的高度由 JS 逐行对齐到「本行最高的单元格」；
   一条编号只显示它自己那张图（衍生形态不再附带基底图）。
3. **所见即所改**。表内单元格可直接编辑，点「同步并保存到 Excel」写回源文件，
   写前自动备份到 **<资产目录>/excel_backup/**（资产目录没填时退回 Excel 同目录 `_backups/`）。
4. **复用第二部分的媒体通道**。图片/音频都走已有的 `/api/local_media`，
   图片灯箱复用第二部分的 `openLightbox()`。

资产编码约定
------------
X 人物 / Y 场景 / Z 道具；基本形态 `X01`、`Y01`、`Z01`，衍生形态加小写后缀
`X01a`、`X01b`、`Y01a`…；音频按 `编号.名称` 命名（`X01.何功伟.mp3`）。
衍生形态的「参考音频」列只填**去掉后缀的裸基底编号**（`X01a` → `X01`）。

编码规范（严格，硬规则 —— 用户定义）
------------------------------------
一条文件名要能被本模块识别，必须同时满足：

1. `X` / `Y` / `Z` 之后是**两位数字**，个位数前补 0：`X01`、`Y02`、`Z01`
2. 后缀字母**必须小写**：`X01a` ✅、`X01A` ❌
3. 之后若要接描述文字，**必须用一个英文句点 `.` 隔开**：
   `X01.何功伟.jpg` ✅、`X01a.何功伟狱中.png` ✅、`X01何功伟行刑.jpg` ❌（少了句点）
4. 不认空格、短横、括号等任何其他分隔符：`X1a 何功伟绝食状态-.jpg` ❌

匹配策略（严格精确，无任何兜底）
--------------------------------
只认完全相同的规范化编码：`X01a` 只找 `X01a.*`，`X01` 只找 `X01.*`。
**`X01` ≠ `X1`** —— 不做前导零归一、不做模糊/近似匹配。

不合规命名（`X1.jpg`、`Y2.wav`、`X01C.png`、`X4A父亲.png`、`X01何功伟.jpg` …）
**一律不参与匹配**，只登记进「违规命名」清单，在页面顶部红字标出，
并给出建议的正确写法（`X4A父亲.png` → `X04a.父亲.png`）。改名后重新加载即生效。

素材目录
--------
图片目录 = 「资产目录」；音频目录 = 显式「音频目录」> `资产目录/Audio` > `资产目录`。
**递归扫描子文件夹**，`X01.flac` 放在 `Audio/02_主角/` 里一样能认到。
"""

import html
import os
import re
import shutil
import urllib.parse
from datetime import datetime

import gradio as gr
import openpyxl
from pydantic import BaseModel

# ==============================================================
# 0. 常量与工具
# ==============================================================
IMG_EXTS = {'.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tif', '.tiff'}
AUD_EXTS = {'.mp3', '.wav', '.aac', '.m4a', '.flac'}

MEDIA_URL = "/api/local_media?filepath={}"

# 合规命名的核心片段：X/Y/Z + 两位数字（个位补 0）+ 可选小写后缀字母
_STRICT_CODE = r"[XYZxyz]\d{2}[a-z]?"

# 一条完整文件名（去掉扩展名后的主干）是否合规：编码 [+ 英文句点 . + 任意描述]
_STEM_CODE_RE = re.compile(rf"^({_STRICT_CODE})(?:\..*)?$")
# 从任意文本里捞合规编码（用于「参考音频」列）
_CODE_IN_TEXT_RE = re.compile(rf"(?<![A-Za-z0-9]){_STRICT_CODE}(?![A-Za-z0-9])")
# 以 X/Y/Z 开头、紧跟数字或分隔符的 —— 一律视为「想写资产编码」，不合规就要报警
_ASSETLIKE_RE = re.compile(r"^[XYZxyz](?=\d|[._\-\s(（])")
# 衍生形态 → 基底：X01a → X01
_VARIANT_RE = re.compile(r"^([XYZ]\d{2})([A-Z]+)$")

# 中间内容列宽度（px，1.0 倍率下）。用户口径：A/B/C 收窄，台词列给足但别太宽。
# **只有中间这些列参与 --a-scale 缩放**（见 _W_IMGVIEW / _W_AUDIOVIEW 的说明）；
# 调宽某列 = 别处等量收回，否则中间部分会被缩得更小。
_COL_WIDTH = {
    "code": 96, "name": 150, "prompt": 250, "refaudio": 150,
    "voice": 256, "line": 272, "sound": 256, "other": 200,
}
# 右端两个固定预览列宽度（px）——**不参与 --a-scale 缩放**。
# 原因：声音预览里的 <audio> 原生控件有个最小可用宽度，整表等比缩小后
# 播放进度条会被挤没（用户实测：MacBook Air「更多内容」缩放下的典型症状）。
# 这两列是工具列，保持恒定尺寸才能保证缩略图和播放条始终看得见。
_W_IMGVIEW, _W_AUDIOVIEW = 210, 300

# 各种列的文本域默认高度（px）。用户口径：偏矮一些，纵向一屏多看到几行。
_TEXT_HEIGHT = {"prompt": 136, "voice": 106, "line": 92, "sound": 106, "other": 92}

# 表头配色（沿用第二部分深色体系）
_TH_CLASS = {
    "name": "a-col-name", "prompt": "a-col-text", "refaudio": "a-col-ref",
    "voice": "a-col-snd", "line": "a-col-snd", "sound": "a-col-snd",
}


def _esc(s) -> str:
    return html.escape("" if s is None else str(s), quote=True)


def media_url(path: str) -> str:
    return MEDIA_URL.format(urllib.parse.quote(path))


# 「复制文件地址」可能带上的引号：资源管理器给的是英文双引号，某些输入法 /
# 聊天工具 / 文档里粘出来的是中文引号，这里一律剥掉。（与主程序保持同一份实现）
_QUOTE_CHARS = "\"'“”‘’「」『』"


def _trim_sep(p: str) -> str:
    """去掉路径尾部的分隔符，但保住盘符根目录与 / 。"""
    if not p or p == "/" or re.match(r"^[A-Za-z]:[\\/]?$", p):
        return p
    return p.rstrip("/\\")


def clean_path(path_str: str) -> str:
    """把用户粘进来的路径洗成干净的绝对路径。

    兼容：首尾空格 / 全角空格 / BOM、被英文或中文引号包裹（含只粘到左引号）、
    浏览器复制来的 file:// 前缀、终端拖拽产生的「反斜杠 + 空格」转义。
    """
    if not path_str:
        return ""
    p = str(path_str).lstrip("\ufeff").replace("\u3000", " ").strip()
    if p[:7].lower() == "file://":
        p = urllib.parse.unquote(p[7:])
        if re.match(r"^/[A-Za-z]:", p):        # file:///C:/x → C:/x
            p = p[1:]
    for _ in range(3):                          # 反复剥，兼容 外层"内层' 这种嵌套
        p = p.strip().strip("\u200b")
        if len(p) >= 2 and p[0] in _QUOTE_CHARS and p[-1] in _QUOTE_CHARS:
            p = p[1:-1]
        elif p and p[0] in _QUOTE_CHARS:        # 只粘到左引号的情况
            p = p[1:]
        else:
            break
    p = p.strip().replace("\\ ", " ")
    if not p:
        return ""
    return os.path.abspath(p)


def natural_sort_key(s: str):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', str(s))]


def col_letter(idx: int) -> str:
    """0 → A、1 → B …… 25 → Z、26 → AA（够用即可）。"""
    n, out = int(idx), ""
    while True:
        out = chr(ord('A') + n % 26) + out
        n = n // 26 - 1
        if n < 0:
            return out


def base_of(code: str) -> str:
    """X01a → X01；本身已是基底（或不合规）则返回空串。"""
    m = _VARIANT_RE.match(str(code or "").strip().upper())
    return m.group(1) if m else ""


def suggest_code(code: str) -> str:
    """给出合规写法建议：X1 → X01、X4A → X04a、X01C → X01c；已合规则返回空串。

    只按硬规则判定：数字必须两位、后缀字母必须小写。编码首字母的大小写不参与判定
    （匹配不区分大小写，`x01` 与 `X01` 视为同一编号，不催改名）。
    """
    raw = str(code or "").strip()
    m = re.match(r"^([XYZxyz])(\d+)([a-zA-Z]*)$", raw)
    if not m:
        return ""
    fixed = f"{m.group(1).upper()}{int(m.group(2)):02d}{m.group(3).lower()}"
    same = fixed == f"{raw[0].upper()}{raw[1:]}"
    return "" if same else fixed


def suggest_filename(fname: str) -> str:
    """给一个不合规文件名，返回建议的合规文件名（无法判断时返回空串）。"""
    stem, ext = os.path.splitext(fname)
    m = re.match(r"^([XYZxyz])(\d+)([a-zA-Z]*)(.*)$", stem)
    if not m:
        return ""
    code_new = f"{m.group(1).upper()}{int(m.group(2)):02d}{m.group(3).lower()}"
    rest = m.group(4).strip().strip("._-—\u3000 ")
    return f"{code_new}.{rest}{ext}" if rest else f"{code_new}{ext}"


def codes_in(text: str) -> list:
    """从一段文本里按出现顺序取出资产编码（去重、**保留原始大小写**）。

    保留大小写是为了让「X01a」不会被显示成 X01A 而误判为后缀大写违规；
    真正的查找由 MediaIndex.resolve 内部统一大写处理。
    """
    out, seen = [], set()
    for m in _CODE_IN_TEXT_RE.finditer(str(text or "")):
        c = m.group(0)
        if c.upper() not in seen:
            seen.add(c.upper())
            out.append(c)
    return out


# ==============================================================
# 1. 媒体索引（递归扫描子文件夹）
# ==============================================================
class MediaIndex:
    """一个目录下的图片或音频索引：严格编码表 + 重复编码冲突 + 违规命名清单。

    只索引合规编码（`X01` / `X01a` / `Y02` / `Z01`）。不合规命名（`X1`、`Y2`、`X001`）
    一律不建索引，只记进 ``violations``，由界面红字提示去改名 —— 保证 `X01` 永不落到 `X1` 上。
    """

    def __init__(self, root_dir: str, valid_exts: set):
        self.root = clean_path(root_dir) if root_dir else ""
        self.index = {}      # 合规编码（已大写）→ 路径
        self.codes = []      # 合规编码清单（磁盘真实大小写）
        self.conflicts = []  # 同一编码多个文件
        self.violations = []  # [(文件名, 建议名)]
        self.files = 0       # 扫描到的媒体文件总数
        self._scan(valid_exts)

    def _scan(self, valid_exts: set):
        if not self.root or not os.path.isdir(self.root):
            return
        groups = {}
        # os.walk 天然递归：子文件夹里的 X01.flac 一样能认到
        for root, dirs, files in os.walk(self.root):
            dirs[:] = [d for d in dirs if not d.startswith('.')]
            for f in sorted(files, key=natural_sort_key):
                if os.path.splitext(f)[1].lower() not in valid_exts:
                    continue
                self.files += 1
                full = os.path.join(root, f)
                stem = os.path.splitext(f)[0]
                m = _STEM_CODE_RE.match(stem)
                if m:
                    groups.setdefault(m.group(1).upper(), []).append((f, full, m.group(1)))
                elif _ASSETLIKE_RE.match(f):
                    # X1 / Y2 / X01C / X4A父亲 / X01何功伟行刑 … 不合规
                    # 只登记不索引，绝不参与匹配
                    self.violations.append((f, suggest_filename(f)))
                else:
                    # 与资产编码无关的命名（如镜头号 s01-01.mp3），按整名登记
                    self.index.setdefault(stem.upper(), full)

        for code, items in groups.items():
            items.sort(key=lambda x: natural_sort_key(x[0]))
            self.index.setdefault(code, items[0][1])
            # 同一编码多个文件 → 按「子目录深度浅者优先」提示，便于判断
            if len(items) > 1:
                self.conflicts.append(
                    f"{code} → " + "、".join(
                        os.path.relpath(p, self.root) for _n, p, _c in items
                    )
                )
        self.codes = sorted({it[2] for items in groups.values() for it in items})

    def resolve(self, code):
        """严格精确匹配：只返回完全同名编码的文件路径，命中不到就是 None。"""
        c = str(code or "").strip().upper()
        if not c:
            return None
        return self.index.get(c)

    def __len__(self):
        return len(self.index)


def find_audio_dir(asset_dir: str, audio_dir: str) -> str:
    """音频目录优先级：显式目录 > 资产目录/Audio > 资产目录（与第二部分一致）。"""
    a = clean_path(asset_dir) if asset_dir else ""
    au = clean_path(audio_dir) if audio_dir else ""
    if au:
        return au
    sub = os.path.join(a, "Audio") if a else ""
    return sub if sub and os.path.isdir(sub) else a


# ==============================================================
# 2. 表结构解析 / 渲染
# ==============================================================
def classify_columns(headers: list):
    """把资产表表头映射成 [(kind, 显示名, 列下标)]。

    kind ∈ code / name / prompt / refaudio / voice / line / sound / other
    用关键词识别，兼容用户增删列或改列名。
    """
    cols = []
    for i, h in enumerate(headers):
        t = str(h or "").strip()
        flat = t.replace(" ", "").replace("\u3000", "")
        if "编号" in flat or "代号" in flat:
            kind = "code"
        elif "名称" in flat or "说明" in flat:
            kind = "name"
        elif "参考音频" in flat or "参考音" in flat:
            kind = "refaudio"
        elif "台词" in flat:
            # 必须排在「捏声」之前：表头常写成「台词（捏声所用）」，含「捏声」二字
            kind = "line"
        elif "声音特质" in flat or "捏声" in flat:
            kind = "voice"
        elif "声音生成" in flat or "环境声" in flat or "音效" in flat:
            kind = "sound"
        elif "文生图" in flat or "图生图" in flat or "提示词" in flat:
            kind = "prompt"
        else:
            kind = "other"
        if i == 0 and kind == "other":    # 位置兜底：第一列必是编号
            kind = "code"
        elif i == 1 and kind == "other":  # 第二列必是名称
            kind = "name"
        cols.append((kind, t or f"列{i + 1}", i))
    return cols


def _text_box(value: str, cls: str, height_px: int, col: int, row: int) -> str:
    """可编辑文本域：直接改，改完点「同步并保存到 Excel」写回源表。

    data-base-h 存住「本列默认高度的基准值」：这个框允许手动拖高（resize: vertical），
    一拖，浏览器就把行内的 calc(...) 换成固定 px，那一格连同整行、连带图片预览的
    高度都会被锁死缩不回去。有了基准值，aRestoreTextHeights() 才能把它真正还原。
    """
    return (
        f'<textarea class="a-text {cls}" data-col="{col}" data-row="{row}" '
        f'data-base-h="{height_px}" '
        f'spellcheck="false" '
        f'style="height:calc({height_px}px * var(--a-row-scale))">{_esc(value)}</textarea>'
    )


def _img_card(code: str, img_idx: "MediaIndex", tag: str = "") -> str:
    path = img_idx.resolve(code)
    if path:
        url = media_url(path)
        return (
            f'<div class="a-card" onclick="openLightbox(\'{url}\')" '
            f'title="点击放大：{_esc(code)}{_esc(tag)}｜{_esc(os.path.relpath(path, img_idx.root))}">'
            f'<img src="{url}" class="a-thumb" alt="{_esc(code)}" loading="lazy"/>'
            f'<span class="a-badge">{_esc(code)}{_esc(tag)}</span></div>'
        )
    fix = suggest_code(code)
    cls = "a-card a-missing" + (" a-badcode" if fix else "")
    tip = (f"编号写法不合规：「{code}」应为 {fix}（X/Y/Z + 两位数字，后缀小写）"
           if fix else "资产目录里没有该编码的图片文件")
    return (
        f'<div class="{cls}" title="{_esc(tip)}">'
        f'<span class="a-badge">{_esc(code)}{_esc(tag)}</span>'
        f'<div class="a-missing-txt">{_esc(fix) if fix else "无文件"}</div></div>'
    )


def _audio_card(code: str, aud_idx: "MediaIndex", tag: str = "") -> str:
    path = aud_idx.resolve(code)
    if path:
        url = media_url(path)
        return (
            f'<div class="a-audio" title="{_esc(os.path.relpath(path, aud_idx.root))}">'
            f'<span class="a-audio-badge">{_esc(code)}{_esc(tag)}</span>'
            f'<audio controls preload="none" src="{url}" class="a-audio-player" '
            f'onclick="event.stopPropagation()" playsinline></audio></div>'
        )
    fix = suggest_code(code)
    cls = "a-audio a-audio-missing" + (" a-badcode" if fix else "")
    txt = f"编号应为 {fix}" if fix else "无音频文件"
    return (
        f'<div class="{cls}">'
        f'<span class="a-audio-badge">{_esc(code)}{_esc(tag)}</span>'
        f'<span class="a-audio-na">{_esc(txt)}</span></div>'
    )


def render_table_a_html(sheet: dict, img_idx: "MediaIndex", aud_idx: "MediaIndex"):
    """把一张资产表（sheet dict）渲染成带右端固定预览列、且可直接编辑的 HTML 表格。"""
    if not sheet or not sheet.get("rows"):
        return "<div class='a-empty'>该资产表为空（没有数据行）</div>"

    headers = sheet["headers"]
    rows = sheet["rows"]
    sheet_name = sheet["name"]

    all_cols = classify_columns(headers)
    code_cols = [c for c in all_cols if c[0] == "code"]
    code_ci = code_cols[0][2] if code_cols else 0

    # 丢掉「无表头且整列为空」的占位列，避免表格里出现莫名其妙的空列
    def _col_all_empty(ci: int) -> bool:
        for r in rows:
            v = r["values"][ci] if ci < len(r["values"]) else None
            if v is not None and str(v).strip() != "":
                return False
        return True

    body_cols = []
    for kind, title, ci in all_cols:
        if kind == "code":
            continue
        if not str(title).strip() and _col_all_empty(ci):
            continue
        if str(title).startswith("列") and _col_all_empty(ci):
            continue
        body_cols.append((kind, title, ci))

    # ---- 列宽：显式 colgroup + 表格总宽，杜绝 table-layout:fixed 下的挤压 ----
    # 中间列 * --a-scale（可缩放）；右端两个预览列恒定 px（不缩放），
    # 所以表格总宽 = 中间之和 * scale + 固定宽度。
    scale_w = [_COL_WIDTH["code"]] + [_COL_WIDTH.get(k, 220) for k, _t, _c in body_cols]
    fixed_w = _W_IMGVIEW + _W_AUDIOVIEW
    base_w = sum(scale_w)
    colgroup = "".join(f'<col style="width:calc({w}px * var(--a-scale))"/>' for w in scale_w)
    colgroup += f'<col style="width:{_W_IMGVIEW}px"/><col style="width:{_W_AUDIOVIEW}px"/>'

    # ---- 表头：A.资产编号 / B.资产名称 / … 前缀用 Excel 真实列号 ----
    headers_html = [
        f'<th class="a-col-code">{col_letter(code_ci)}.'
        f'{_esc(str(headers[code_ci] or "资产编号").strip())}</th>'
    ]
    for kind, title, ci in body_cols:
        headers_html.append(
            f'<th class="{_TH_CLASS.get(kind, "a-col-text")}">{col_letter(ci)}.{_esc(title)}</th>'
        )
    headers_html.append('<th class="a-col-imgview">📷 图片预览</th>')
    headers_html.append('<th class="a-col-audioview">🔊 声音预览</th>')

    ref_ci = next((ci for kind, _t, ci in body_cols if kind == "refaudio"), None)

    rows_html = []
    for r in rows:
        vals = r["values"]
        excel_row = r["excel_row"]

        def _v(ci):
            x = vals[ci] if ci < len(vals) else None
            return "" if x is None else str(x)

        code = _v(code_ci).strip()

        # ---- 首列：可编辑编号 + 不合规提示 ----
        fix_self = suggest_code(code)
        code_inner = (
            f'<input type="text" class="a-code-input" data-col="{code_ci}" data-row="{excel_row}" '
            f'value="{_esc(code)}" spellcheck="false"/>'
        )
        if fix_self:
            code_inner += f'<div class="a-code-warn">应为 {_esc(fix_self)}</div>'
        cells = [f'<td class="a-col-code">{code_inner}</td>']

        for kind, _title, ci in body_cols:
            val = _v(ci)
            if kind == "name":
                cells.append(
                    f'<td class="a-col-name">'
                    f'<textarea class="a-text a-name-box" data-col="{ci}" data-row="{excel_row}" '
                    f'data-autogrow="1" spellcheck="false" '
                    f'style="height:calc(62px * var(--a-row-scale))">{_esc(val)}</textarea></td>'
                )
            elif kind == "refaudio":
                # 裸基底编号：框高由 JS 按内容自动撑开（aAutoGrowShort），不再把字裁掉
                cells.append(
                    f'<td class="a-col-ref"><div class="a-ref">'
                    f'<textarea class="a-text a-ref-box" data-col="{ci}" data-row="{excel_row}" '
                    f'data-autogrow="1" spellcheck="false" rows="1">{_esc(val)}</textarea></div></td>'
                )
            else:
                cls = {"prompt": "a-prompt", "voice": "a-voice",
                       "line": "a-line", "sound": "a-sound"}.get(kind, "")
                cells.append(
                    f'<td class="{_TH_CLASS.get(kind, "a-col-text")}">'
                    f'{_text_box(val, cls, _TEXT_HEIGHT.get(kind, 104), ci, excel_row)}</td>'
                )

        # ---- 右端固定：图片预览（只显示本条编号的图；X01a 不再附带 X01 基座图）----
        img_cards = _img_card(code, img_idx) if code else ""
        cells.append(f'<td class="a-col-imgview"><div class="a-grid">{img_cards}</div></td>')

        # ---- 右端固定：声音预览（本条音频优先，其次「参考音频」列里的基底）----
        ref_text = _v(ref_ci) if ref_ci is not None else ""
        aud_codes = [c for c in [code] if c]
        for c in codes_in(ref_text):
            if c.upper() not in [x.upper() for x in aud_codes]:
                aud_codes.append(c)
        cards = "".join(
            _audio_card(c, aud_idx, "" if i == 0 else "·基底")
            for i, c in enumerate(aud_codes[:3])
        )
        cells.append(f'<td class="a-col-audioview"><div class="a-audio-stack">{cards}</div></td>')

        rows_html.append(
            f'<tr class="a-row" data-code="{_esc(code)}" data-row="{excel_row}" '
            f'onclick="onAssetRowClick(this)">' + "".join(cells) + "</tr>"
        )

    return (
        f'<div class="a-viewport" id="assetViewport" data-sheet-name="{_esc(sheet_name)}">'
        f'<table class="a-table" id="assetTable" data-base-width="{base_w}" '
        f'data-fixed-width="{fixed_w}" '
        f'style="width:calc({base_w}px * var(--a-scale) + {fixed_w}px)">'
        f'<colgroup>{colgroup}</colgroup>'
        f'<thead><tr>{"".join(headers_html)}</tr></thead>'
        f'<tbody>{"".join(rows_html)}</tbody></table></div>'
    )


# ==============================================================
# 3. 读取 / 保存 Excel
# ==============================================================
def read_sheet(excel_path: str, sheet_name: str) -> dict:
    """读一张工作表为 {name, headers, rows:[{excel_row, values}], ncols}。

    用 openpyxl 直读并**保留每行真实的 Excel 行号**，这样保存时可以按绝对行写回，
    即使表中有空行、也不会发生行错位。
    """
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    try:
        ws = wb[sheet_name]
        raw = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()

    if not raw:
        return {"name": sheet_name, "headers": [], "rows": [], "ncols": 0}

    headers = raw[0]
    # 有些表第一行是空/标题行 → 若首行全空则跳过
    body = raw[1:]
    start_row = 2
    if all(v is None or str(v).strip() == "" for v in headers) and body:
        headers, body, start_row = body[0], body[1:], 3

    ncols = max([len(headers)] + [len(r) for r in body]) if body else len(headers)
    if ncols < len(headers):
        ncols = len(headers)

    def _pad(row):
        return list(row) + [None] * (ncols - len(row))

    rows = []
    for i, r in enumerate(body):
        r = _pad(r)
        if all(v is None or str(v).strip() == "" for v in r):
            continue  # 整行空 → 不渲染（也不会被保存覆盖）
        rows.append({"excel_row": start_row + i, "values": r})

    return {"name": sheet_name, "headers": _pad(headers), "rows": rows, "ncols": ncols}


def backup_excel_file(filepath: str, asset_dir: str = "") -> str:
    """写前备份：**<资产目录>/excel_backup/<原名>_<时间戳><扩展名>**。

    用户口径：Excel 和它的历史版本归到资产目录下的同一个 `excel_backup/` 里，好找。
    资产目录没设置 / 不存在时退回老位置 —— Excel 同目录的 `_backups/`。
    返回备份文件路径；源文件不存在时返回空串。
    """
    if not os.path.exists(filepath):
        return ""
    base = clean_path(asset_dir) if asset_dir else ""
    if base and os.path.isdir(base):
        backup_dir = os.path.join(base, "excel_backup")
    else:
        backup_dir = os.path.join(os.path.dirname(filepath), "_backups")
    os.makedirs(backup_dir, exist_ok=True)
    name, ext = os.path.splitext(os.path.basename(filepath))
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(backup_dir, f"{name}_{ts}{ext}")
    shutil.copy2(filepath, dest)
    return dest


def write_sheet(excel_path: str, sheet_name: str, rows: list, asset_dir: str = "") -> str:
    """把 rows（[{row, cells:{列下标: 值}}]）写回指定工作表。返回备份文件路径。"""
    if not os.path.exists(excel_path):
        raise FileNotFoundError("Excel 文件不存在")
    backup_path = backup_excel_file(excel_path, asset_dir)

    wb = openpyxl.load_workbook(excel_path)
    if sheet_name not in wb.sheetnames:
        raise KeyError(f"工作表「{sheet_name}」不存在")
    ws = wb[sheet_name]

    for item in rows:
        r = int(item.get("row") or 0)
        if r < 1:
            continue
        for col_key, val in (item.get("cells") or {}).items():
            c = int(col_key) + 1
            ws.cell(row=r, column=c, value=("" if val is None else str(val)))
    wb.save(excel_path)
    return backup_path


# ==============================================================
# 4. 后端控制器
# ==============================================================
class AssetTableApp:
    def __init__(self):
        self.excel_path = ""
        self.sheets = []
        self.current = ""
        self.img_idx = MediaIndex("", IMG_EXTS)
        self.aud_idx = MediaIndex("", AUD_EXTS)
        self.img_root = ""
        self.aud_root = ""

    def refresh_index(self, asset_dir, audio_dir):
        """**每次都重扫**：子文件夹里新丢进来的 X01.flac 立刻可见，不吃缓存。"""
        self.img_root = clean_path(asset_dir)
        self.aud_root = find_audio_dir(asset_dir, audio_dir)
        self.img_idx = MediaIndex(self.img_root, IMG_EXTS)
        self.aud_idx = MediaIndex(self.aud_root, AUD_EXTS)

    def load_project(self, path_a, asset_dir, audio_dir, current_table=None):
        """读 sheet 列表并渲染默认表。返回 (Dropdown 更新, HTML, 统计 HTML)。"""
        self.refresh_index(asset_dir, audio_dir)
        self.excel_path = clean_path(path_a)

        if not self.excel_path or not os.path.exists(self.excel_path):
            self.sheets, self.current = [], ""
            return gr.update(choices=[], value=None), (
                "<div class='a-empty'>没找到资产表文件。<br>请在顶部「表格 A (资产表)」填好"
                "资产 Excel 的完整路径，再点 🚀 加载/刷新工程。</div>"
            )

        try:
            wb = openpyxl.load_workbook(self.excel_path, read_only=True)
            # 严格按源文件里 sheet 的原有顺序，不做任何排序
            self.sheets = list(wb.sheetnames)
            wb.close()
        except Exception as e:  # noqa: BLE001
            self.sheets, self.current = [], ""
            return gr.update(choices=[], value=None), (
                f"<div class='a-empty'>读取资产表失败：{_esc(e)}</div>"
            )

        if not self.sheets:
            return gr.update(choices=[], value=None), "<div class='a-empty'>资产表里没有工作表</div>"

        target = current_table if current_table in self.sheets else self.sheets[0]
        self.current = target
        return gr.update(choices=self.sheets, value=target), self.load_table(target)

    def load_table(self, sheet_name):
        if not self.excel_path or not os.path.exists(self.excel_path) or not sheet_name:
            return "<div class='a-empty'>请先加载资产表</div>"
        self.current = sheet_name
        try:
            sheet = read_sheet(self.excel_path, sheet_name)
        except Exception as e:  # noqa: BLE001
            return f"<div class='a-empty'>读取工作表「{_esc(sheet_name)}」失败：{_esc(e)}</div>"
        return render_table_a_html(sheet, self.img_idx, self.aud_idx)

    def step_table(self, direction, current):
        if not self.sheets or current not in self.sheets:
            return current, self.load_table(current)
        i = self.sheets.index(current) + direction
        if 0 <= i < len(self.sheets):
            return self.sheets[i], self.load_table(self.sheets[i])
        return current, self.load_table(current)

    def stats_html(self):
        parts = [f"📷 图片资产 <b>{len(self.img_idx.codes)}</b> 个编号",
                 f"🔊 声音资产 <b>{len(self.aud_idx.codes)}</b> 个编号"]
        if self.img_idx.conflicts:
            parts.append(f"<span class='a-warn'>⚠ 图片编号重复 {len(self.img_idx.conflicts)} 组</span>")
        if self.aud_idx.conflicts:
            parts.append(f"<span class='a-warn'>⚠ 声音编号重复 {len(self.aud_idx.conflicts)} 组</span>")
        bad = self.img_idx.violations + self.aud_idx.violations
        if bad:
            shown = "；".join(
                f"{_esc(f)} → <b>{_esc(s)}</b>" if s else _esc(f)
                for f, s in bad[:3]
            )
            more = f"，另有 {len(bad) - 3} 个" if len(bad) > 3 else ""
            parts.append(
                f"<span class='a-badcode-txt'>⛔ 不合规命名 {len(bad)} 个，<b>不参与匹配</b>：{shown}{more}</span>"
            )
        parts.append(f"🔍 音频目录：<code>{_esc(self.aud_root or '（未设置）')}</code>（含子文件夹）")
        return f"<div class='a-hint'>{' ｜ '.join(parts)}</div>"


# ==============================================================
# 5. 样式（追加到 CUSTOM_CSS 末尾，全用 .a- 前缀，不影响第二部分）
# ==============================================================
ASSET_CSS = """
/* ================= 第一部分：资产表预览 ================= */
.a-hint {
    flex: 1;
    font-size: 11px;
    color: #8c9ba5;
    line-height: 1.55;
    padding: 2px 6px;
}
.a-hint b { color: #ebcb8b; }
.a-hint code {
    color: #a3be8c;
    background: #16181f;
    padding: 0 3px;
    border-radius: 3px;
    font-size: 10.5px;
}
.a-warn { color: #f0a0a8; }
.a-badcode-txt { color: #d08770; }
.a-empty {
    padding: 28px 16px;
    text-align: center;
    color: #8c9ba5;
    font-size: 13px;
    line-height: 1.8;
    background: #181a20;
    border: 1px dashed #2b303c;
    border-radius: 8px;
}

.a-viewport {
    width: 100%;
    overflow: auto;
    max-height: 72vh;
    border: 1px solid #2b303c;
    border-radius: 8px;
    background: #181a20;
    -webkit-overflow-scrolling: touch;
}

.a-table {
    border-collapse: separate;
    border-spacing: 0;
    table-layout: fixed;
    font-size: 12px;
}

.a-table thead th {
    position: sticky;
    position: -webkit-sticky;
    top: 0;
    background: #20242c;
    color: #c8d1e0;
    font-weight: 600;
    text-align: left;
    height: 34px;
    line-height: 34px;
    padding: 0 8px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    border-bottom: 2px solid #2e3440;
    z-index: 10;
}

.a-table tbody td {
    border-bottom: 1px solid #282c37;
    padding: 5px 6px;
    vertical-align: top;
    overflow: hidden;
}

/* 首列表头是「A.资产编号」，字号略小、内边距收窄；放不下时**折成两行**而不是省略成
   「A.资…」（窄视口下中间列会被等比缩小，靠省略号会看不出是哪一列）。 */
.a-table thead th.a-col-code {
    font-size: 11px;
    padding: 0 4px;
    white-space: normal;
    line-height: 14px;
    overflow: hidden;
    text-overflow: clip;
    word-break: break-all;
}

/* 统计行（右侧那一大段提示）限高 + 内部滚动：不合规命名多时不会把上面那行撑成一大块，
   纵向空间留给表格本体，一屏能多看好几行。 */
.a-stats-box { max-height: 82px; overflow-y: auto; }
.a-stats-box::-webkit-scrollbar { width: 6px; }
.a-stats-box::-webkit-scrollbar-thumb { background: #3b4252; border-radius: 3px; }

/* 表格 A 的「整行大字」预览框：比第二部分矮一些，给表格多留点纵向空间 */
#masterBoxA { height: 62px; }

.a-row:hover { background-color: #1a1e27; }
.a-row.a-selected { background-color: #242e3f !important; }

/* 列宽由 <colgroup> 统一控制，表格总宽在渲染时按列宽之和给出 */

/* ---- 右端固定预览列：横向滚动时永远贴在右边 ----
   注意：这两列**不参与 --a-scale 缩放**（宽度恒为 _W_IMGVIEW / _W_AUDIOVIEW px），
   所以 sticky 偏移量也得是常数，不能再乘 var(--a-scale)，否则播放条会被挤掉。 */
.a-table thead th.a-col-imgview {
    position: sticky;
    position: -webkit-sticky;
    top: 0;
    right: __W_AUDVIEW__;
    z-index: 22;
    background: #262c37;
    box-shadow: -1px 0 0 #2e3440;
}
.a-table thead th.a-col-audioview {
    position: sticky;
    position: -webkit-sticky;
    top: 0;
    right: 0;
    z-index: 22;
    background: #262c37;
}
.a-table tbody td.a-col-imgview {
    position: sticky;
    position: -webkit-sticky;
    right: __W_AUDVIEW__;
    z-index: 6;
    background: #1b1f27;
    box-shadow: -1px 0 0 #2e3440;
}
.a-table tbody td.a-col-audioview {
    position: sticky;
    position: -webkit-sticky;
    right: 0;
    z-index: 6;
    background: #1b1f27;
}

/* 编号输入 */
.a-code-input {
    width: 100%;
    box-sizing: border-box;
    font-family: monospace;
    font-size: 13px;
    font-weight: 700;
    color: #88c0d0;
    background: #161d26;
    border: 1px solid #2f4650;
    border-radius: 4px;
    text-align: center;
    padding: 5px 2px;
    outline: none;
}
.a-code-input:focus { border-color: #88c0d0; }
.a-code-warn {
    margin-top: 3px;
    font-size: 9.5px;
    line-height: 1.3;
    color: #d08770;
    text-align: center;
    border: 1px dashed #8a4b3a;
    border-radius: 3px;
    padding: 1px 2px;
}

/* 参考音频（D 列）：高度交给 aAutoGrowShort() 按内容撑开，绝不把字裁掉 */
.a-ref { display: block; }
.a-ref-box {
    font-family: monospace;
    font-size: 10.5px !important;
    line-height: 1.5 !important;
    color: #a3be8c !important;
    background: #16181f !important;
    border: 1px solid #3e4452 !important;
    min-height: 32px;
    overflow: hidden;
    resize: none;
    overflow-wrap: anywhere;
}

/* 可编辑文本域 */
.a-text {
    background: #14161b;
    border: 1px solid #2e3440;
    border-radius: 4px;
    color: #d8dee9;
    padding: 6px;
    font-size: 11.5px;
    line-height: 1.55;
    width: 100%;
    box-sizing: border-box;
    resize: vertical;
    overflow-y: auto;
    font-family: inherit;
    outline: none;
}
.a-text:focus { border-color: #5b6b86; background: #171a21; }
.a-name-box {
    font-size: 12px;
    color: #f4f6fa;
    height: auto;
    min-height: 44px;
    resize: none;
}
.a-prompt { color: #eceff4; }
.a-voice  { color: #d7e0c8; }
.a-line   { color: #f0dcb4; }
.a-sound  { color: #c5d5e8; }

/* 图片预览：整列宽 × 「本行最高单元格」高（--a-img-h 由 aFitImageHeights() 逐行写入）。
   一条编号只显示自己那张图 —— 衍生形态（X01a）不再附带 X01 基座图。 */
.a-grid { display: block; width: 100%; }
.a-card {
    position: relative;
    width: 100%;
    height: var(--a-img-h, 96px);
    border-radius: 3px;
    overflow: hidden;
    border: 1px solid #3b4252;
    background: #0a0a0f;
    box-sizing: border-box;
    cursor: zoom-in;
}
.a-thumb { width: 100%; height: 100%; object-fit: contain; display: block; }
.a-card.a-missing {
    border: 1px dashed #bf616a;
    cursor: default;
    height: auto;
    min-height: 56px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
}
.a-missing-txt { font-size: 9px; color: #bf616a; }
.a-card.a-badcode { border-color: #d08770; border-style: solid; }
.a-card.a-badcode .a-missing-txt { color: #d08770; font-weight: 700; }
.a-audio.a-badcode { border: 1px dashed #d08770; }
.a-audio.a-badcode .a-audio-na { color: #d08770; }
.a-badge {
    position: absolute;
    bottom: 0; left: 0; right: 0;
    background: rgba(0, 0, 0, 0.86);
    color: #88c0d0;
    font-size: 9px;
    line-height: 14px;
    text-align: center;
}

/* 声音预览 */
.a-audio-stack { display: flex; flex-direction: column; gap: 4px; }
.a-audio {
    display: flex;
    align-items: center;
    gap: 6px;
    background: #232731;
    border-radius: 4px;
    padding: 3px 6px;
}
.a-audio-missing { opacity: 0.55; border: 1px dashed #5c2e2e; }
.a-audio-badge {
    font-family: monospace;
    font-size: 10px;
    color: #a3be8c;
    font-weight: 700;
    min-width: 46px;
    flex-shrink: 0;
}
.a-audio-na { font-size: 10px; color: #bf616a; }
.a-audio-player { height: 28px; width: 100%; flex: 1; min-width: 140px; }
"""

# 把 CSS 里的固定预览列宽占位符换成真实常量值（两处常量只在这里定义一次，不会写歪）
ASSET_CSS = ASSET_CSS.replace("__W_AUDVIEW__", f"{_W_AUDIOVIEW}px")


# ==============================================================
# 6. 前端脚本（追加到 CUSTOM_JS 末尾）
# ==============================================================
ASSET_JS = """
/* ================= 第一部分：资产表视图微调 ================= */
let aScale = parseFloat(localStorage.getItem('assetScale')) || 1.0;
let aRowScale = parseFloat(localStorage.getItem('assetRowScale')) || 1.0;

function applyAssetScales() {
    const root = document.documentElement;
    root.style.setProperty('--a-scale', aScale);
    root.style.setProperty('--a-row-scale', aRowScale);
}

/* 把「手动拖过」的文本域高度还原成默认的 calc(...)（基准值在 data-base-h 里）。
   否则内联的固定 px 会盖住 --a-row-scale：这一格缩不回去，行高缩不回去，
   图片预览的 --a-img-h（量行高得来）也跟着锁死。                      */
window.aRestoreTextHeights = function() {
    const boxes = document.querySelectorAll('#assetViewport textarea.a-text[data-base-h]');
    for (let i = 0; i < boxes.length; i++) {
        const el = boxes[i];
        el.style.height = 'calc(' + el.dataset.baseH + 'px * var(--a-row-scale))';
    }
};
window.aAdjustWidth = function(delta) {
    aScale = Math.max(0.5, Math.min(2.5, aScale + delta));
    applyAssetScales();
    aAutoGrowShort();
    aFitImageHeights();
};
window.aAdjustHeight = function(delta) {
    aRowScale = Math.max(0.5, Math.min(2.5, aRowScale + delta));
    applyAssetScales();
    aRestoreTextHeights();
    aAutoGrowShort();
    aFitImageHeights();
};
/* 表格总宽 = 中间可缩放列 * --a-scale + 右端两个固定预览列（恒定 px，不缩放） */
window.aTableMetrics = function(table) {
    return {
        base: parseFloat(table.dataset.baseWidth || '0'),
        fixed: parseFloat(table.dataset.fixedWidth || '0')
    };
};
window.aFitWidth = function() {
    const vp = document.getElementById('assetViewport');
    if (!vp) return;
    const table = vp.querySelector('table');
    if (!table) return;
    const m = aTableMetrics(table);
    if (m.base > 0) {
        const avail = vp.clientWidth - 24 - m.fixed;
        aScale = Math.max(0.5, Math.min(1.5, avail / m.base));
        applyAssetScales();
        aAutoGrowShort();
        aFitImageHeights();
    }
};
/* 换表后自动适配：只「收」不「放」，保证最右的台词列不被 sticky 预览列盖住 */
window.aAutoFitIfNeeded = function(force) {
    const vp = document.getElementById('assetViewport');
    if (!vp || !vp.clientWidth) return;          // 隐藏状态下量不到宽度，跳过
    const table = vp.querySelector('table');
    if (!table) return;
    const m = aTableMetrics(table);
    if (!m.base) return;
    const avail = vp.clientWidth - 20 - m.fixed; // 预览列不缩放，先把它们的固定宽度扣掉
    if (!force && m.base * aScale <= avail + 2) return;   // 已经放得下
    aScale = Math.max(0.55, Math.min(1.0, avail / m.base));
    applyAssetScales();
};
/* 图片预览逐行等高：卡片高度 = 本行最高的「非预览列」单元格高度。
   ⚠️ 必须先「清零再量」——
   同一个 <tr> 里所有 <td> 的高度是相等的（行高由最高的那一格决定），所以直接读
   td.offsetHeight，拿到的是「已经被图片自己撑高之后的整行高度」，于是
   「量出来 → 写回图片 → 行更高 → 下次量更大」，只涨不落：
   用户按 Shift+g / r 都缩不回去，就是因为这个自我强化循环。
   现在的顺序是：先把所有图片格的高度归零 → 一次性量行高（此时行高只由文字格决定）
   → 再写回。这样图片永远跟着文字走，缩放和「恢复默认」都能立即回落。 */
window.aFitImageHeights = function() {
    const vp = document.getElementById('assetViewport');
    if (!vp || !vp.clientWidth) return;
    const rows = vp.querySelectorAll('tr.a-row');
    if (!rows.length) return;

    // 1) 全部归零（只写不读，不触发逐个重排）
    const pairs = [];
    for (let i = 0; i < rows.length; i++) {
        const cell = rows[i].querySelector('td.a-col-imgview');
        if (!cell) continue;
        cell.style.setProperty('--a-img-h', '0px');
        pairs.push([rows[i], cell]);
    }
    if (!pairs.length) return;

    // 2) 归零后统一测量：只量中间列，既避免自我强化，也避免和 sticky 列互相影响
    const picks = [];
    for (let k = 0; k < pairs.length; k++) {
        const tr = pairs[k][0];
        let h = 0;
        const tds = tr.querySelectorAll('td');
        for (let j = 0; j < tds.length; j++) {
            const td = tds[j];
            if (td.classList.contains('a-col-imgview') ||
                td.classList.contains('a-col-audioview')) continue;
            if (td.offsetHeight > h) h = td.offsetHeight;
        }
        if (!h) h = tr.offsetHeight;
        picks.push([pairs[k][1], Math.max(56, Math.round(h - 12))]);
    }

    // 3) 写回
    for (let m = 0; m < picks.length; m++) {
        picks[m][0].style.setProperty('--a-img-h', picks[m][1] + 'px');
    }
};
window.aSaveScale = function() {
    localStorage.setItem('assetScale', aScale);
    localStorage.setItem('assetRowScale', aRowScale);
    showToast('资产表栏宽与行高已保存为默认！', 'success');
};
window.aResetScale = function() {
    aScale = 1.0; aRowScale = 1.0;
    localStorage.removeItem('assetScale');
    localStorage.removeItem('assetRowScale');
    applyAssetScales();
    aRestoreTextHeights();
    aAutoGrowShort();
    aFitImageHeights();
    showToast('资产表已恢复默认排版', 'success');
};

/* 点行 → 上方大字看整行文本（与第二部分的 masterPromptBox 同款体验） */
window.onAssetRowClick = function(tr) {
    document.querySelectorAll('tr.a-row').forEach(r => r.classList.remove('a-selected'));
    tr.classList.add('a-selected');
    const code = tr.dataset.code || '';
    const boxes = tr.querySelectorAll('textarea.a-text');
    let parts = [];
    boxes.forEach(b => { if (b.value.trim()) parts.push(b.value); });
    const box = document.getElementById('masterBoxA');
    if (!box) return;
    box.textContent = '';
    box.insertAdjacentHTML('beforeend', "<b style='color:#88c0d0'>[" + code + "]</b> ");
    if (parts.length) {
        const span = document.createElement('span');
        span.style.whiteSpace = 'pre-wrap';
        span.textContent = parts.join('\\n\\n');
        box.appendChild(span);
    } else {
        box.insertAdjacentHTML('beforeend', "<i style='color:#8c9ba5'>该行无文本内容</i>");
    }
};

/* ---------- 表格 A：编辑后就地保存回 Excel（写前自动备份） ---------- */
window.saveAssetToExcel = function() {
    const vp = document.getElementById('assetViewport');
    const sheetName = vp ? vp.dataset.sheetName : '';
    if (!sheetName) { showToast('请先加载资产表', 'error'); return; }

    const trs = vp.querySelectorAll('tbody tr.a-row');
    const rows = [];
    trs.forEach(tr => {
        const r = parseInt(tr.dataset.row, 10);
        if (!r) return;
        const cells = {};
        tr.querySelectorAll('[data-col]').forEach(el => {
            cells[el.dataset.col] = el.value;
        });
        rows.push({row: r, cells: cells});
    });
    if (!rows.length) { showToast('表格数据为空', 'error'); return; }

    fetch('/api/save_excel_a', {
        method: 'POST',
        headers: {'Content-Type': 'application/json'},
        body: JSON.stringify({sheet_name: sheetName, rows: rows})
    })
    .then(res => res.json())
    .then(res => showToast(res.msg, res.status === 'success' ? 'success' : 'error'))
    .catch(() => showToast('保存失败：网络错误', 'error'));
};

/* ---------- 短文本框自动撑高：D 列「参考音频」、B 列「资产名称」 ----------
   这两个框内容短但会被换行，固定高度会把字裁掉（unset 高度又受 rows 属性限制），
   所以渲染后按 scrollHeight 逐个撑开；CSS 里的 min-height 负责兜底最小值。      */
window.aAutoGrowShort = function() {
    const boxes = document.querySelectorAll(
        'textarea.a-ref-box, textarea.a-name-box, textarea[data-autogrow="1"]');
    for (let i = 0; i < boxes.length; i++) {
        const el = boxes[i];
        el.style.height = 'auto';                       // 先归零再量，避免只增不减
        const need = (el.scrollHeight || 0) + 2;
        el.style.height = Math.min(260, Math.max(32, need)) + 'px';
    }
};

/* 手动改动时同步撑高（不用等重新渲染） */
document.addEventListener('input', function(e) {
    const el = e.target;
    if (!el || !el.classList || !el.classList.contains('a-ref-box')) return;
    el.style.height = 'auto';
    el.style.height = Math.min(260, Math.max(32, (el.scrollHeight || 0) + 2)) + 'px';
});

/* ---------- 键盘导航三件套：切 sheet / 切 Tab / 缩放 ----------
   a s  → 当前这张表的上一个 / 下一个 sheet
          （表格 A = 「上一表 / 下一表」，表格 B = 「上一场 / 下一场」）
   d f  → 切 Tab 页面（d → 表格 A，f → 表格 B）
   g h  → 横向缩放；Shift+g / Shift+h → 纵向缩放；r → 恢复默认
   键盘上 a s d f g h 连成一排，手指不用挪窝。                       */

/* 当前可见的 Tab 面板（gradio 通常只挂载活动面板，这里再兜一层防御） */
window.sbsVisiblePane = function() {
    const panels = Array.prototype.slice.call(document.querySelectorAll('[role="tabpanel"]'));
    if (!panels.length) return document;
    const vis = panels.filter(function(p) { return p.offsetParent !== null; });
    return vis.length ? vis[0] : document;
};

/* 在指定范围内按关键词依次找按钮并点击 */
window.sbsClickIn = function(scope, keywords) {
    const btns = Array.prototype.slice.call((scope || document).querySelectorAll('button'));
    for (let i = 0; i < keywords.length; i++) {
        const hit = btns.find(function(b) {
            return (b.textContent || '').indexOf(keywords[i]) >= 0;
        });
        if (hit) { hit.click(); return true; }
    }
    return false;
};

/* a / s：切当前这张表的 sheet（A 表走「上一表/下一表」，B 表走「上一场/下一场」） */
window.sbsStepSheet = function(dir) {
    const pane = sbsVisiblePane();
    const keys = dir < 0 ? ['上一表', '上一场'] : ['下一表', '下一场'];
    return sbsClickIn(pane, keys);
};

/* d / f：在两个 Tab 页面之间切换 */
window.sbsSwitchTab = function(which) {
    const want = (which === 'A' ? '表格A' : '表格B');
    const tabs = Array.prototype.slice.call(document.querySelectorAll('[role="tab"]'));
    const hit = tabs.find(function(t) {
        return (t.textContent || '').replace(/\\s+/g, '').indexOf(want) >= 0;
    });
    if (!hit) return false;
    hit.click();
    // 切回表格 A 时才量得到宽度 → 补一次自动适配 + 短框自适应
    if (which === 'A') {
        setTimeout(function() { aAutoFitIfNeeded(); aAutoGrowShort(); aFitImageHeights(); }, 300);
    }
    return true;
};
/* 兼容旧名（早期版本用的是 switchSbsTab('表格 A')） */
window.switchSbsTab = function(keyword) {
    return sbsSwitchTab((keyword || '').indexOf('A') >= 0 ? 'A' : 'B');
};

/* Shift + 滚轮 → 资产表横向滚动（与第二部分同款手感） */
document.addEventListener('wheel', function(e) {
    const vp = document.getElementById('assetViewport');
    if (!vp || !vp.contains(e.target)) return;
    if (e.shiftKey) {
        e.preventDefault();
        vp.scrollLeft += e.deltaY;
    }
}, { passive: false });

/* 窗口尺寸变化（含 MacBook「更多内容」这类分辨率切换）后重算一次 */
window.addEventListener('resize', function() {
    clearTimeout(window.__aResizeTimer);
    window.__aResizeTimer = setTimeout(function() {
        aAutoGrowShort(); aAutoFitIfNeeded(); aFitImageHeights();
    }, 150);
});

setTimeout(function() { applyAssetScales(); aAutoGrowShort(); aFitImageHeights(); }, 600);
"""


# ==============================================================
# 7. 保存接口（由 StoryBoardStudio.py 注册到同一个 FastAPI 实例）
# ==============================================================
_APP_A = None  # build_asset_tab() 时指向真实的 AssetTableApp 实例


class SaveAssetRequest(BaseModel):
    sheet_name: str
    rows: list


def register_asset_api(fastapi_app):
    """把表格 A 的保存接口挂到主应用的 FastAPI 上。"""

    @fastapi_app.post("/api/save_excel_a")
    def save_excel_a_api(req: SaveAssetRequest):
        if _APP_A is None or not _APP_A.excel_path or not os.path.exists(_APP_A.excel_path):
            return {"status": "error", "msg": "保存失败：资产表文件不存在"}
        if not req.sheet_name or not req.rows:
            return {"status": "error", "msg": "保存失败：数据不完整"}
        try:
            backup = write_sheet(_APP_A.excel_path, req.sheet_name, req.rows, _APP_A.img_root)
        except PermissionError:
            return {"status": "error", "msg": "❌ 保存失败：Excel 正在被 Excel/WPS 占用，请先关闭"}
        except Exception as e:  # noqa: BLE001
            return {"status": "error", "msg": f"❌ 保存失败: {e}"}
        name = os.path.basename(_APP_A.excel_path)
        bak = os.path.basename(backup) if backup else "（无备份）"
        return {"status": "success",
                "msg": f"✅ 已保存 {name} · {req.sheet_name}｜备份：{bak}"}

    return fastapi_app


# ==============================================================
# 8. UI 构建（需在调用方的 gr.Tabs() 上下文里调用）
# ==============================================================
def build_asset_tab(txt_path_a, txt_asset_dir, txt_audio_dir, btn_load_project,
                    label="📊 表格 A · 影视资产表"):
    """在当前 gr.Blocks() 上下文里建出第一部分页面并绑定事件。"""
    global _APP_A
    app_a = AssetTableApp()
    _APP_A = app_a

    with gr.TabItem(label):
        with gr.Row(elem_classes="top-toolbar-row"):
            dd_tables_a = gr.Dropdown(label="当前资产表", choices=[], interactive=True, scale=2)
            btn_prev_a = gr.Button("◀ 上一表", scale=0, min_width=60)
            btn_next_a = gr.Button("下一表 ▶", scale=0, min_width=60)
            html_stats_a = gr.HTML("<div class='a-hint'>等待加载…</div>", elem_classes="a-stats-box")

        # 说明文案（预览列吸附 / 快捷键 / 命名硬规则）已移到 README.md，
        # 界面上不再占纵向空间；`.a-keycap` 样式随之弃用。

        with gr.Row(elem_classes="view-toolbar-row"):
            gr.Markdown("**📐 视图微调**", scale=0)
            a_w_dec = gr.Button("◀ 栏宽收窄", scale=1)
            a_w_inc = gr.Button("▶ 栏宽拉宽", scale=1)
            a_h_dec = gr.Button("▲ 行高收紧", scale=1)
            a_h_inc = gr.Button("▼ 行高拉开", scale=1)
            a_fit = gr.Button("↔ 适配宽度", scale=1)
            a_save = gr.Button("📌 记住排版", scale=1)
            a_reset = gr.Button("🔄 恢复默认", scale=1)

        with gr.Row(elem_classes="prompt-save-row"):
            gr.HTML("""<div id="masterBoxA" class="master-prompt-preview"></div>""")
            a_save_excel = gr.Button("💾 同步并保存到 Excel", variant="primary",
                                     scale=0, elem_classes="save-btn")

        html_asset_view = gr.HTML(
            "<div class='a-empty'>等待加载资产表…<br>"
            "在顶部填好「表格 A (资产表)」路径与「资产目录」后，点 🚀 加载/刷新工程。</div>"
        )

        # ---------- 事件绑定 ----------
        a_w_dec.click(None, None, None, js="aAdjustWidth(-0.1)")
        a_w_inc.click(None, None, None, js="aAdjustWidth(0.1)")
        a_h_dec.click(None, None, None, js="aAdjustHeight(-0.1)")
        a_h_inc.click(None, None, None, js="aAdjustHeight(0.1)")
        a_fit.click(None, None, None, js="aFitWidth()")
        a_save.click(None, None, None, js="aSaveScale()")
        a_reset.click(None, None, None, js="aResetScale()")
        a_save_excel.click(None, None, None, js="saveAssetToExcel()")

        def _load(path_a, asset_dir, audio_dir, cur):
            upd, body = app_a.load_project(path_a, asset_dir, audio_dir, cur)
            return upd, body, app_a.stats_html()

        def _step(direction, cur):
            sheet, body = app_a.step_table(direction, cur)
            return sheet, body, app_a.stats_html()

        # 挂在同一个「加载/刷新工程」按钮上：第二部分那一路照旧，本路只负责资产表
        # 每次换表后补一次自动适配（只收不放）+ 短框撑高 + 图片逐行等高，
        # 避免台词列被预览列压住、D 列裁字、图片预览忽大忽小
        _FIT_JS = ("setTimeout(function(){ aAutoGrowShort(); aAutoFitIfNeeded(); "
                   "aFitImageHeights(); }, 320)")

        btn_load_project.click(
            fn=_load,
            inputs=[txt_path_a, txt_asset_dir, txt_audio_dir, dd_tables_a],
            outputs=[dd_tables_a, html_asset_view, html_stats_a],
        ).then(None, None, None, js=_FIT_JS)
        # .input() 只在用户手动选择时触发，避免与程序化赋值重复渲染
        dd_tables_a.input(fn=app_a.load_table, inputs=[dd_tables_a],
                          outputs=[html_asset_view]).then(None, None, None, js=_FIT_JS)
        btn_prev_a.click(fn=lambda cur: _step(-1, cur), inputs=[dd_tables_a],
                         outputs=[dd_tables_a, html_asset_view, html_stats_a]
                         ).then(None, None, None, js=_FIT_JS)
        btn_next_a.click(fn=lambda cur: _step(1, cur), inputs=[dd_tables_a],
                         outputs=[dd_tables_a, html_asset_view, html_stats_a]
                         ).then(None, None, None, js=_FIT_JS)

    return app_a
