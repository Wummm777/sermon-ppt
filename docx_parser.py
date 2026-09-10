# -*- coding: utf-8 -*-
"""解析带 tag 的讲道稿 docx，输出结构化对象。

tag 规范（每个 tag 独占一段，中文行紧跟英文行配对）：
  [TITLE]/[TITLE_EN]   题目
  [SREF]/[SREF_EN]     主题经文出处
  [POINT]/[POINT_EN]   大点
  [SUB]/[SUB_EN]       小点
  [VERSE]/[VERSE_EN]   经文全文（出处前缀如【创11:4】含在文本内）

经文段落的字符级样式（粗体/斜体/下划线/颜色）会被原样保留。
"""
import re
from dataclasses import dataclass, field

from docx import Document

TAGS = {"TITLE", "SREF", "POINT", "SUB", "VERSE"}
TAG_RE = re.compile(r"^\s*\[([A-Z_]+)\]\s*")

# 行内样式标记：粗体 **…**、下划线 __…__、颜色 {c:RRGGBB}…{/c}
MARKUP_RE = re.compile(r"(\*\*|__|\{c:[0-9A-Fa-f]{6}\}|\{/c\})")

# 经文出处前缀（如【创11:4】），仅用于识别是否混入单独的出处行
VREF_LINE_RE = re.compile(r"^\s*(?:【[^】]{1,40}】|\([^)]{1,40}\)|（[^）]{1,40}）)\s*$")


def parse_line(text):
    """把带样式标记的文本解析成 Run 列表（标记可嵌套，开关式）。"""
    runs = []
    bold = underline = False
    color = None
    for part in MARKUP_RE.split(text):
        if part == "**":
            bold = not bold
        elif part == "__":
            underline = not underline
        elif part.startswith("{c:") and part.endswith("}"):
            color = part[3:-1].upper()
        elif part == "{/c}":
            color = None
        elif part:
            runs.append(Run(text=part, bold=bold, italic=False,
                            underline=underline, color=color))
    return runs


def wrap(text, bold=False, underline=False, color=None):
    """把纯文本包成带样式标记的文本（生成端，与 parse_line 对应）。"""
    out = text
    if underline:
        out = "__" + out + "__"
    if color:
        out = "{c:" + color + "}" + out + "{/c}"
    if bold:
        out = "**" + out + "**"
    return out


@dataclass
class Run:
    """最小富文本单元，保留 Word 里的字符样式。"""
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    color: str | None = None  # RRGGBB 十六进制，None 表示跟随默认


@dataclass
class Pair:
    """中英对照的一对内容，排版与分割时视为一个整体。"""
    zh: list = field(default_factory=list)
    en: list = field(default_factory=list)

    def text(self, lang):
        runs = self.zh if lang == "zh" else self.en
        return "".join(r.text for r in runs).strip()


@dataclass
class VerseBlock:
    ref: Pair          # 经文出处
    text: Pair         # 经文全文


@dataclass
class Sub:
    heading: Pair
    verses: list = field(default_factory=list)


@dataclass
class Point:
    heading: Pair
    subs: list = field(default_factory=list)
    verses: list = field(default_factory=list)  # 直接挂在大点下的经文（少见）


@dataclass
class Sermon:
    title: Pair = None
    sref: Pair = None   # 主题经文出处（无全文）
    points: list = field(default_factory=list)


def _make_run(text, r):
    """从 docx 的 run 提取文本与字符样式（字号/字体名忽略，由脚本预设控制）。"""
    f = r.font
    color = None
    try:
        if f.color is not None and f.color.type is not None and f.color.rgb is not None:
            color = str(f.color.rgb)
    except Exception:
        color = None
    u = r.underline
    underline = u is True or (u not in (None, False) and getattr(u, "name", "") != "NONE")
    return Run(text=text, bold=bool(f.bold), italic=bool(f.italic),
               underline=underline, color=color)


def _slice_runs(para, start):
    """取段落文本 start 偏移之后的 runs（即剥掉 tag 部分），保留样式。"""
    runs = []
    pos = 0
    for r in para.runs:
        seg_start, seg_end = pos, pos + len(r.text)
        pos = seg_end
        if seg_end <= start:
            continue
        text = r.text[max(0, start - seg_start):]
        if text:
            runs.append(_make_run(text, r))
    return runs


def read_raw_paras(path):
    """规则模式入口：读 docx 全部段落的纯文本与原生 runs（样式完整保留）。

    返回 [(行号, 文本, [Run, ...]), ...]，空段跳过。
    """
    doc = Document(path)
    out = []
    for lineno, para in enumerate(doc.paragraphs, 1):
        text = para.text.strip()
        if not text:
            continue
        runs = []
        for r in para.runs:
            if r.text:
                runs.append(_make_run(r.text, r))
        out.append((lineno, text, runs))
    return out


def read_items(path):
    """按顺序读出 docx 中所有带 tag 的段落。

    返回 (items, skipped)：
      items   = [(行号, tag, [Run, ...]), ...]
      skipped = [(行号, 内容开头), ...]  没有 tag 的段落
    """
    doc = Document(path)
    items, skipped = [], []
    for lineno, para in enumerate(doc.paragraphs, 1):
        text = para.text
        if not text.strip():
            continue
        m = TAG_RE.match(text)
        if not m:
            skipped.append((lineno, text.strip()[:40]))
            continue
        tag = m.group(1)
        if tag == "VREF":
            # [VREF] tag 已废弃：出处应与正文同在 [VERSE] 行内
            skipped.append((lineno, "[VREF]已废弃 " + text.strip()[:30]))
            continue
        items.append((lineno, tag, _slice_runs(para, m.end())))
    return items, skipped


def read_items_from_text(text):
    """从纯文本（剪贴板/txt）读带 tag 的行，样式标记解析为 Run。

    兼容处理：若出现单独成行的经文出处（如「【创11:4】」），视为错误报告给用户。
    经文出处应与正文同在 [VERSE] 行内。
    """
    items, skipped = [], []
    for lineno, line in enumerate(text.splitlines(), 1):
        if not line.strip():
            continue
        m = TAG_RE.match(line)
        if not m:
            skipped.append((lineno, line.strip()[:40]))
            continue
        tag = m.group(1)
        content = line[m.end():]
        # 旧版规范残留检测：[VREF] 单独的出处行，或 [VERSE] 行只有出处没有正文
        if tag in ("VREF", "SREF") and content.strip() and VREF_LINE_RE.match(content.strip()):
            # SREF 单独出处是正常结构，仅 VREF 残留才是错误
            if tag == "VREF":
                skipped.append((lineno, line.strip()[:40]))
                continue
        if tag == "VREF":
            # [VREF] tag 已废弃：把它当 skipped 处理，报错提示改用新规范
            skipped.append((lineno, "[VREF]已废弃 " + line.strip()[:30]))
            continue
        items.append((lineno, tag, parse_line(content)))
    return items, skipped


def parse_sermon(items, skipped):
    """把 items 组装成 Sermon。严格模式：任何问题都进 errors，有错则不生成 pptx。"""
    errors = []
    for l, t in skipped:
        if t.startswith("[VREF]已废弃"):
            errors.append(f"第{l}段：[VREF] 已废弃——新版规范中经文出处（如【罗5:1】）应写在经文正文同一行内，"
                          f"请用新版「AI提示词模板」重新加工")
        else:
            errors.append(f"第{l}段：缺少 tag 标记，无法识别（内容开头：{t}…）")
    sermon = Sermon()
    i, n = 0, len(items)

    def pair_at(base):
        """从 items[i] 开始取 [base] + [base_EN] 一对。"""
        nonlocal i
        lineno_zh = items[i][0]
        p = Pair(zh=items[i][2])
        i += 1
        if i < n and items[i][1] == base + "_EN":
            lineno_en = items[i][0]
            p.en = items[i][2]
            i += 1
            if not p.text("zh"):
                errors.append(f"第{lineno_zh}段：[{base}] 中文内容为空")
            if not p.text("en"):
                errors.append(f"第{lineno_en}段：[{base}_EN] 英文内容为空")
        else:
            errors.append(f"第{lineno_zh}段：[{base}] 之后缺少配对的 [{base}_EN]")
        return p

    if n == 0:
        errors.append("文档为空：没有任何带 tag 的段落")
        return sermon, errors

    if items[i][1] == "TITLE":
        sermon.title = pair_at("TITLE")
    else:
        errors.append(f"第{items[i][0]}段：讲道稿必须以 [TITLE] 开头，实际是 [{items[i][1]}]")
        i += 1

    if i < n:
        if items[i][1] == "SREF":
            sermon.sref = pair_at("SREF")
        else:
            errors.append(f"第{items[i][0]}段：题目之后应为 [SREF] 主题经文出处，实际是 [{items[i][1]}]")
            i += 1

    cur_point = None
    cur_sub = None
    while i < n:
        lineno, tag, runs = items[i]
        if tag == "POINT":
            cur_point = Point(heading=pair_at("POINT"))
            sermon.points.append(cur_point)
            cur_sub = None
        elif tag == "SUB":
            if cur_point is None:
                errors.append(f"第{lineno}段：[SUB] 出现在任何 [POINT] 之前")
                cur_point = Point(heading=Pair())
                sermon.points.append(cur_point)
            cur_sub = Sub(heading=pair_at("SUB"))
            cur_point.subs.append(cur_sub)
        elif tag == "VERSE":
            # 经文块两行一组：[VERSE]中文全文（含出处前缀）→ [VERSE_EN]英文全文
            verse_zh = items[i][2]
            lineno_zh = items[i][0]
            i += 1
            verse_en, lineno_ven = [], lineno_zh
            if i < n and items[i][1] == "VERSE_EN":
                verse_en = items[i][2]
                lineno_ven = items[i][0]
                i += 1
            else:
                errors.append(f"第{lineno_zh}段：[VERSE] 之后缺少配对的 [VERSE_EN]")
            if not "".join(r.text for r in verse_zh).strip():
                errors.append(f"第{lineno_zh}段：[VERSE] 中文内容为空")
            if not "".join(r.text for r in verse_en).strip():
                errors.append(f"第{lineno_ven}段：[VERSE_EN] 英文内容为空")
            block = VerseBlock(ref=Pair(), text=Pair(zh=verse_zh, en=verse_en))
            if cur_sub is not None:
                cur_sub.verses.append(block)
            elif cur_point is not None:
                cur_point.verses.append(block)
            else:
                errors.append(f"第{lineno_zh}段：经文出现在任何 [POINT] 之前")
        elif tag in ("TITLE", "SREF"):
            errors.append(f"第{lineno}段：[{tag}] 只能出现一次且必须在开头")
            pair_at(tag)
        elif tag in TAGS:
            errors.append(f"第{lineno}段：[{tag}] 没有对应的中文件，顺序错误")
            i += 1
        else:
            if tag.startswith("VREF"):
                errors.append(f"第{lineno}段：[{tag}] 已废弃——新版规范中经文出处（如【罗5:1】）应写在经文正文同一行内，"
                              f"请用新版「AI提示词模板」重新加工")
            else:
                errors.append(f"第{lineno}段：无法识别的 tag [{tag}]")
            i += 1

    if not sermon.points:
        errors.append("整篇没有 [POINT] 大点")
    for pi, p in enumerate(sermon.points, 1):
        name = p.heading.text("zh")[:20] or "(空)"
        if not p.subs and not p.verses:
            errors.append(f"大点{pi}（{name}）之下没有任何小点或经文")
    return sermon, errors
