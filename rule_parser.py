# -*- coding: utf-8 -*-
"""规则匹配引擎：按行首/关键字规则把无 tag 的讲道稿行分类为结构元素。

判定优先级（先命中先归类）：
  题目/主题经文出处（关键字包含）→ 大点（行首编号/绪论关键字）
  → 小点（行首编号）→ 经文正文（【 前缀）
四者都不命中的行 → 报错（严格模式，不静默丢弃）。

配对规则：中文行（含任意汉字）后紧跟的纯英文行配成一对 {zh, en}；
无配对的行，另一侧为空字符串（渲染时只显单语）。

归组：经文/小点出现在任何大点之前 → 归入虚拟「绪论」大点；
经文出现在任何小点之前 → 归入虚拟「经文」小点。
"""
import re

CN_RE = re.compile(r"[一-鿿]")

TITLE_KEYS = ("题目：", "title：", "title:")
SREF_KEYS = ("经文：", "scripture：", "scripture:")
INTRO_CN = ("绪论", "引言")
INTRO_EN = ("introduction", "preface")

H1_RE = re.compile(r"^[一二三四五六七八九十百]+[、，.]|^\d+[，.](?!\d)")
H2_RE = re.compile(r"^(\d+[)）]|（\d+）|\(\d+\))")

SPLIT_COLON_RE = re.compile(r"[:：]")


def is_zh_line(line):
    return bool(CN_RE.search(line))


def strip_style_markers(line):
    """AI 模式的标记在规则模式里先剥掉再判定（两模式兼容同一份输入）。"""
    # 不改变正文中标记的位置，只用于"行首特征"判断；返回原行 + 去标记视图
    view = line
    view = view.replace("**", "").replace("__", "")
    view = re.sub(r"\{c:[0-9A-Fa-f]{6}\}|\{/c\}", "", view)
    return view


def split_colon_value(text):
    """按第一个冒号切分，取后半段。无冒号返回原文。"""
    parts = SPLIT_COLON_RE.split(text.strip(), 1)
    return parts[-1].strip() if len(parts) == 2 else text.strip()


def pair_lines(raw_lines):
    """把原始行配成中英对。

    raw_lines = [(lineno, text, runs)]，runs 可为 None（纯文本来源）。
    返回 [(lineno, zh_text, en_text, zh_runs, en_runs), ...]；
    无配对的行另一侧为 ""（runs 同步为 None/[]）。
    中文行后紧跟纯英文行 → 配对；中文行连续出现或英文行连续出现时各自独立成对。
    """
    cleaned = []
    for item in raw_lines:
        lineno, line = item[0], item[1]
        runs = item[2] if len(item) > 2 else None
        t = line.strip()
        if t:
            cleaned.append((lineno, t, runs))
    pairs = []
    i, n = 0, len(cleaned)
    while i < n:
        lineno, text, runs = cleaned[i]
        if is_zh_line(strip_style_markers(text)):
            zh, zh_runs = text, runs
            en, en_runs = "", None
            if i + 1 < n and not is_zh_line(strip_style_markers(cleaned[i + 1][1])):
                en, en_runs = cleaned[i + 1][1], cleaned[i + 1][2]
                i += 2
            else:
                i += 1
            pairs.append((lineno, zh, en, zh_runs, en_runs))
        else:
            pairs.append((lineno, "", text, None, runs))
            i += 1
    return pairs


def classify(pairs):
    """把中英对分类为条目流。

    返回 (items, errors)。items 为 [(lineno, kind, zh, en), ...]，
    kind ∈ title/sref/point/sub/verse。中文/英文分别判定，一侧可空。
    """
    items, errors = [], []
    seen_title = seen_sref = False

    def classify_side(text, side):
        """单侧文本 → (kind, value)。kind=None 表示未命中。"""
        if not text:
            return None, ""
        v = strip_style_markers(text)
        low = v.lower()
        # 1) 题目 / 主题经文出处（关键字包含，不限行首）
        if any(k in v or k in low for k in TITLE_KEYS):
            return "title", split_colon_value(v)
        if any(k in v or k in low for k in SREF_KEYS):
            return "sref", split_colon_value(v)
        # 2) 大点
        if H1_RE.match(v) or any(k in v for k in INTRO_CN) or any(k in low for k in INTRO_EN):
            return "point", v
        # 3) 小点
        if H2_RE.match(v):
            return "sub", v
        # 4) 经文正文
        if v.startswith("【"):
            return "verse", v
        return None, v

    for lineno, zh, en, zh_runs, en_runs in pairs:
        kzh, vzh = classify_side(zh, "zh")
        ken, ven = classify_side(en, "en")

        # 中英两侧判定不一致 → 报错（同一条对的两半必须同类）
        if kzh and ken and kzh != ken:
            errors.append(f"第{lineno}行：中英两行类型判定不一致（中文判定为{kzh}，英文判定为{ken}），"
                          f"中文：{zh[:30]}… 英文：{en[:30]}…")
            continue
        kind = kzh or ken
        if kind is None:
            errors.append(f"第{lineno}行：无法识别该行类型（不匹配题目/大点/小点/经文的任何特征），"
                          f"内容：{(zh or en)[:40]}…")
            continue
        if kind == "title":
            if seen_title:
                errors.append(f"第{lineno}行：题目重复出现（只允许一个题目），内容：{(zh or en)[:30]}…")
                continue
            seen_title = True
        if kind == "sref":
            if seen_sref:
                errors.append(f"第{lineno}行：主题经文出处重复出现（只允许一个），内容：{(zh or en)[:30]}…")
                continue
            seen_sref = True
        # 中英各自取值：若该侧没命中关键字值则用另一侧的取值方式
        value_zh = vzh if kzh == kind or kzh is None else vzh
        value_en = ven if ken == kind or ken is None else ven
        items.append((lineno, kind, value_zh, value_en, zh_runs, en_runs))
    return items, errors


def build_sermon(items, errors):
    """把分类条目组装成 Sermon 对象（与 docx_parser.parse_sermon 的产物同构）。

    追加的结构错误也进 errors（严格模式统一报告）。
    """
    from docx_parser import Sermon, Point, Sub, VerseBlock, Pair, Run

    sermon = Sermon()
    cur_point = None
    cur_sub = None
    intro_created = False

    def ensure_intro():
        """经文/小点出现在任何大点之前 → 虚拟绪论大点。"""
        nonlocal cur_point, intro_created
        if cur_point is None:
            cur_point = Point(heading=Pair(zh=[Run("绪论")], en=[Run("Introduction")]))
            sermon.points.append(cur_point)
            intro_created = True

    def ensure_verse_sub(point):
        """经文出现在任何小点之前 → 虚拟「经文」小点。"""
        nonlocal cur_sub
        if not point.subs:
            cur_sub = Sub(heading=Pair(zh=[Run("经文")], en=[Run("Scripture")]))
            point.subs.append(cur_sub)
        return point.subs[-1]

    for item in items:
        lineno, kind, vzh, ven, zh_runs, en_runs = item
        def runs_of(text, runs):
            # docx 来源：用原生 runs（样式保留）；纯文本来源：造无样式 Run
            if runs:
                return list(runs)
            return [Run(text)]

        if kind == "title":
            sermon.title = Pair(zh=runs_of(vzh, zh_runs), en=runs_of(ven, en_runs))
        elif kind == "sref":
            sermon.sref = Pair(zh=runs_of(vzh, zh_runs), en=runs_of(ven, en_runs))
        elif kind == "point":
            cur_point = Point(heading=Pair(zh=runs_of(vzh, zh_runs), en=runs_of(ven, en_runs)))
            sermon.points.append(cur_point)
            cur_sub = None
        elif kind == "sub":
            ensure_intro()
            cur_sub = Sub(heading=Pair(zh=runs_of(vzh, zh_runs), en=runs_of(ven, en_runs)))
            cur_point.subs.append(cur_sub)
        elif kind == "verse":
            ensure_intro()
            block = VerseBlock(ref=Pair(),
                               text=Pair(zh=runs_of(vzh, zh_runs), en=runs_of(ven, en_runs)))
            if cur_sub is None:
                sub = ensure_verse_sub(cur_point)
                sub.verses.append(block)
            else:
                cur_sub.verses.append(block)

    # 结构校验（与 AI 模式同标准）
    if sermon.title is None:
        errors.append("未找到题目行（需要包含「题目：」或「Title:」的行）")
    if sermon.sref is None:
        errors.append("未找到主题经文出处行（需要包含「经文：」或「Scripture:」的行）")
    if not sermon.points:
        errors.append("未找到任何大点（需要行首「一、」/「1.」编号或包含「绪论」/「Introduction」）")
    for pi, p in enumerate(sermon.points, 1):
        if not p.subs:
            errors.append(f"大点{pi}（{p.heading.text('zh')[:20]}）之下没有任何小点或经文")
    return sermon, errors


def parse_by_rules(raw_lines, source_name="输入"):
    """规则模式入口：raw_lines = [(lineno, text), ...]。

    返回 (sermon, errors)。
    """
    pairs = pair_lines(raw_lines)
    items, errors = classify(pairs)
    if errors:
        return None, errors
    sermon, errors2 = build_sermon(items, errors)
    return sermon, errors2
