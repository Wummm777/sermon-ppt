# -*- coding: utf-8 -*-
"""把结构化讲道稿排版成 pptx。

页面规则（设计定案）：
  页1（固定）：题目 + 主题经文出处，别无他物
  大点+其第一小点 = 固定一页（无小点则配第一块经文，皆无则只放大点标题）
  其余内容自动分割：以中英整体为最小单位，绝不拆开；
    页面按基础字号尽量多放整单位，放不下就翻页；
    只有当单个单位在基础字号下都放不下一页时，才无下限缩小字号
  左侧内容区永远中英双语（中文在上英文在下）；
  右侧提纲栏按页轮换单语（第1页中文、第2页英文……），
    当前小点（或大点固定页上的大点）文字变红高亮
"""
import math
import os

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn


# ---------------------------------------------------------------- 测量估算

def text_width_units(text, cfg):
    """估算文本宽度，单位是"字号倍数"：宽度(英寸) = units * size / 72。"""
    lay = cfg["layout"]
    w = 0.0
    for ch in text:
        o = ord(ch)
        if o >= 0x2E80 or ch in "，。！？；：、“”‘’（）《》、—…·【】":
            w += lay["cjk_width"]
        elif ch == " ":
            w += lay["space_width"]
        else:
            w += lay["latin_width"]
    return w


def _run_text(r):
    return r["text"] if isinstance(r, dict) else r.text


def para_height(runs, size, avail_w_in, cfg):
    """估算一个段落渲染后的高度（英寸）。"""
    lay = cfg["layout"]
    units = sum(text_width_units(_run_text(r), cfg) for r in runs)
    w_in = units * size / 72.0
    lines = max(1, math.ceil(w_in * lay["height_safety"] / max(avail_w_in, 0.1)))
    return lines * size * lay["line_spacing"] / 72.0


# ---------------------------------------------------------------- 排版项

def _joined(pair, lang):
    return "".join(r.text for r in (pair.zh if lang == "zh" else pair.en)).strip()


def _item(runs, size, *, bold=False, color=None, indent=0.0, space_after=0.0):
    return {"runs": list(runs), "size": size, "bold": bold, "color": color,
            "indent": indent, "space_after": space_after}


def point_items(point, cfg):
    s = cfg["sizes"]
    c = cfg["colors"]
    zh = _joined(point.heading, "zh")
    en = _joined(point.heading, "en")
    sp = cfg["layout"]["space_after_pt"]
    return [
        _item([{"text": zh}], s["point_zh"], bold=True, color=c["point"], space_after=sp),
        _item([{"text": en}], s["point_en"], bold=True, color=c["point"], space_after=sp),
    ]


def sub_heading_items(sub, cfg):
    s = cfg["sizes"]
    c = cfg["colors"]
    zh = _joined(sub.heading, "zh")
    en = _joined(sub.heading, "en")
    sp = cfg["layout"]["space_after_pt"]
    return [
        _item([{"text": zh}], s["sub_zh"], bold=True, color=c["sub"], space_after=sp),
        _item([{"text": en}], s["sub_en"], bold=True, color=c["sub"], space_after=sp),
    ]


def verse_items(vb, cfg):
    """一块经文 = 中文全文 + 英文全文（出处前缀已含在文本内），作为整体单位。"""
    s = cfg["sizes"]
    c = cfg["colors"]
    lay = cfg["layout"]
    ind = lay["verse_indent_in"]
    sp = lay["space_after_pt"]
    items = [
        _item(list(vb.text.zh), s["verse_zh"], color="000000", indent=ind, space_after=sp),
        _item(list(vb.text.en), s["verse_en"], color=c["verse_en"], indent=ind, space_after=sp),
    ]
    return items


# ---------------------------------------------------------------- 页面规划

class Page:
    def __init__(self, kind):
        self.kind = kind          # title / content
        self.items = []
        self.highlight = None     # ("point", pi) 或 ("sub", pi, si) 或 None
        self.scale = 1.0


def _split_units(units, first_items, avail_w, box_h, cfg):
    """贪心装箱：units 中每个元素是一个"单位"（item 列表），整单位换页绝不拆开。"""
    pages, cur = [], list(first_items)
    for u in units:
        if cur and not (items_height(cur + u, 1.0, avail_w, cfg) <= box_h):
            pages.append(cur)
            cur = list(u)
        else:
            cur = cur + u
    if cur:
        pages.append(cur)
    return pages


def plan_pages(sermon, cfg, avail_w, box_h):
    pages = [Page("title")]
    sp = cfg["layout"]["space_after_pt"]
    pages[0].items = [
        _item([{"text": _joined(sermon.title, "zh")}], cfg["sizes"]["title_zh"],
              bold=True, color=cfg["colors"]["title"], space_after=sp),
        _item([{"text": _joined(sermon.title, "en")}], cfg["sizes"]["title_en"],
              bold=True, color=cfg["colors"]["title"], space_after=sp),
        _item([{"text": _joined(sermon.sref, "zh")}], cfg["sizes"]["sref_zh"],
              color=cfg["colors"]["ref"], space_after=sp),
        _item([{"text": _joined(sermon.sref, "en")}], cfg["sizes"]["sref_en"],
              color=cfg["colors"]["ref"], space_after=sp),
    ]

    def finish(raw_items, highlight):
        pg = Page("content")
        pg.items = raw_items
        pg.highlight = highlight
        pg.scale = page_scale(raw_items, avail_w, box_h, cfg)
        pages.append(pg)

    for pi, point in enumerate(sermon.points):
        if point.subs:
            # 固定页：大点 + 第一小点，别无其他
            finish(point_items(point, cfg) + sub_heading_items(point.subs[0], cfg), ("point", pi))
            # 第一小点的经文从下一页开始
            for pg_items in _split_units([verse_items(v, cfg) for v in point.subs[0].verses],
                                         [], avail_w, box_h, cfg):
                finish(pg_items, ("sub", pi, 0))
            # 其余小点：小点标题 + 其经文，自动分割
            for si, sub in enumerate(point.subs[1:], 1):
                for pg_items in _split_units([verse_items(v, cfg) for v in sub.verses],
                                             sub_heading_items(sub, cfg), avail_w, box_h, cfg):
                    finish(pg_items, ("sub", pi, si))
            # 直接挂在大点下的经文（少见，兜底）
            for pg_items in _split_units([verse_items(v, cfg) for v in point.verses],
                                         [], avail_w, box_h, cfg):
                finish(pg_items, ("point", pi))
        elif point.verses:
            # 兜底：大点没有小点 → 固定页放大点 + 第一块经文
            finish(point_items(point, cfg) + verse_items(point.verses[0], cfg), ("point", pi))
            for pg_items in _split_units([verse_items(v, cfg) for v in point.verses[1:]],
                                         [], avail_w, box_h, cfg):
                finish(pg_items, ("point", pi))
        else:
            finish(point_items(point, cfg), ("point", pi))
    return pages


def items_height(items, scale, avail_w_in, cfg):
    h = 0.0
    for it in items:
        size = it["size"] * scale
        runs = [{"text": r["text"] if isinstance(r, dict) else r.text} for r in it["runs"]]
        w = avail_w_in - it["indent"]
        h += para_height(runs, size, w, cfg) + it["space_after"] / 72.0
    return h


def page_scale(items, avail_w, box_h, cfg):
    """若整页在基础字号下放不下（即单单位超页），无下限缩字号，步进约2%。"""
    s = 1.0
    if items_height(items, s, avail_w, cfg) <= box_h:
        return s
    while s > 0.15:
        s -= 0.02
        if items_height(items, s, avail_w, cfg) <= box_h:
            return s
    return s


# ---------------------------------------------------------------- 渲染

def _set_run_fonts(run, cfg, latin=None, ea=None):
    run.font.name = latin or cfg["fonts"]["latin"]
    rPr = run._r.get_or_add_rPr()
    ea_name = ea or cfg["fonts"]["east_asian"]
    for tagname in ("a:ea", "a:cs"):
        el = rPr.find(qn(tagname))
        if el is None:
            el = rPr.makeelement(qn(tagname), {})
            rPr.append(el)
        el.set("typeface", ea_name)


def _add_para(tf, used_first, runs, size, *, bold=False, color=None,
              indent=0.0, space_after=0.0, line_spacing=1.15,
              latin=None, ea=None):
    p = tf.paragraphs[0] if not used_first else tf.add_paragraph()
    p.alignment = PP_ALIGN.LEFT
    p.line_spacing = line_spacing
    if space_after:
        p.space_after = Pt(space_after)
    if indent:
        pPr = p._p.get_or_add_pPr()
        pPr.set("marL", str(int(Inches(indent))))
        pPr.set("indent", "0")
    for r in runs:
        run = p.add_run()
        run.text = r["text"] if isinstance(r, dict) else r.text
        f = run.font
        f.size = Pt(size)
        f.bold = bold or bool(r.get("bold") if isinstance(r, dict) else r.bold)
        f.italic = bool(r.get("italic") if isinstance(r, dict) else r.italic)
        f.underline = bool(r.get("underline") if isinstance(r, dict) else r.underline)
        c = (r.get("color") if isinstance(r, dict) else r.color) or color
        if c:
            try:
                f.color.rgb = RGBColor.from_string(c)
            except Exception:
                pass
        _set_run_fonts(run, _G_CFG, latin=latin, ea=ea)
    return True  # used_first


_G_CFG = None  # 渲染期间的全局 cfg（供 _set_run_fonts 使用）


def _add_textbox(slide, x, y, w, h):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.03)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    return tf


def compute_boxes(cfg):
    W, H = cfg["slide"]["width_in"], cfg["slide"]["height_in"]
    m, gap = cfg["layout"]["margin_in"], cfg["layout"]["gap_in"]
    cw = W * cfg["slide"]["content_ratio"]
    content = (m, m, cw - 2 * m, H - 2 * m)
    ox = cw + gap
    outline = (ox, m, W - ox - m, H - 2 * m)
    return {"content": content, "outline": outline}


def render_content(slide, page, cfg, box):
    x, y, w, h = box["content"]
    tf = _add_textbox(slide, x, y, w, h)
    avail_w = w - 0.08
    used = False
    for it in page.items:
        if not any((r["text"] if isinstance(r, dict) else r.text).strip() for r in it["runs"]):
            continue
        used = _add_para(tf, used, it["runs"], it["size"] * page.scale,
                         bold=it["bold"], color=it["color"], indent=it["indent"],
                         space_after=it["space_after"],
                         line_spacing=cfg["layout"]["line_spacing"])


def outline_entries(sermon, lang, highlight, cfg):
    """右侧提纲的条目列表（纯结构：题目→主题经文出处→大点→缩进小点）。

    编号完全照抄原文（讲稿自带编号体系），脚本绝不自动添加。
    字号/粗体/间距由 config 的 outline_* 字段统一控制。
    """
    o = cfg["outline"]
    size = o["size_zh"] if lang == "zh" else o["size_en"]
    sp = o["space_after_pt"]
    cc = cfg["colors"]
    entries = [
        {"kind": "title", "text": sermon.title.text(lang), "size": size,
         "bold": o["bold_except_sub"], "indent": 0.0, "space_after": sp, "key": None,
         "color": cc["title"]},
        {"kind": "sref", "text": sermon.sref.text(lang), "size": size,
         "bold": o["bold_except_sub"], "indent": 0.0, "space_after": sp, "key": None,
         "color": cc["ref"]},
    ]
    for pi, p in enumerate(sermon.points):
        entries.append({"kind": "point", "text": p.heading.text(lang), "size": size,
                        "bold": o["bold_except_sub"], "indent": 0.0, "space_after": sp,
                        "key": ("point", pi), "color": cc["point"]})
        for si, sub in enumerate(p.subs):
            entries.append({"kind": "sub", "text": sub.heading.text(lang), "size": size,
                            "bold": o["bold_except_sub"] or o["bold_sub"],
                            "indent": cfg["layout"]["outline_sub_indent_in"],
                            "space_after": sp, "key": ("sub", pi, si), "color": cc["sub"]})
    return entries


def render_outline(slide, sermon, lang, highlight, cfg, box):
    entries = outline_entries(sermon, lang, highlight, cfg)
    x, y, w, h = box["outline"]
    avail_w = w - 0.06

    # 超长则整体缩小，尽量塞进一页
    scale = 1.0
    while scale > 0.3:
        total = sum(para_height([{"text": e["text"]}], e["size"] * scale, avail_w - e["indent"], cfg)
                    + e["space_after"] / 72.0 for e in entries)
        if total <= h:
            break
        scale *= 0.93

    tf = _add_textbox(slide, x, y, w, h)
    hl_color = cfg["colors"]["highlight"]
    o = cfg["outline"]
    ls = o["line_spacing"]
    used = False
    for e in entries:
        used = _add_para(tf, used, [{"text": e["text"]}], max(8, e["size"] * scale),
                         bold=e["bold"],
                         color=hl_color if (highlight and e["key"] == highlight) else e["color"],
                         indent=e["indent"], space_after=e["space_after"],
                         line_spacing=ls,
                         latin=o.get("font_latin"), ea=o.get("font_zh"))


def build_pptx(sermon, cfg, out_path):
    global _G_CFG
    _G_CFG = cfg
    prs = Presentation()
    prs.slide_width = Inches(cfg["slide"]["width_in"])
    prs.slide_height = Inches(cfg["slide"]["height_in"])
    blank = prs.slide_layouts[6]

    box = compute_boxes(cfg)
    avail_w = box["content"][2] - 0.08
    box_h = box["content"][3] - 0.05
    pages = plan_pages(sermon, cfg, avail_w, box_h)

    for no, page in enumerate(pages, 1):
        slide = prs.slides.add_slide(blank)
        # 背景图：先铺满整页，文字叠加在其上
        bg = cfg.get("background", {}).get("resolved")
        if bg:
            slide.shapes.add_picture(bg, 0, 0,
                                     width=prs.slide_width, height=prs.slide_height)
        lang = "zh" if no % 2 == 1 else "en"   # 第1页中文提纲，第2页英文……
        if page.kind == "title":
            render_content(slide, page, cfg, box)
        else:
            render_content(slide, page, cfg, box)
        render_outline(slide, sermon, lang, page.highlight, cfg, box)

    prs.save(out_path)
    return len(pages)
