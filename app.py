# -*- coding: utf-8 -*-
"""讲道稿转PPT · 主界面（单 exe，双模式）。

模式：
  规则匹配（推荐）：docx/txt → 行首/关键字规则自动分类 → pptx（docx 保留原生样式）
  AI 加工：docx/txt → 提取带样式标记的文本 → 贴给 DeepSeek 加 tag → 粘回 → pptx

界面：主窗口（选文件、选模式、生成）+ 排版设置弹窗（config 全字段可视化编辑 +
背景图与透明度）。所有持久化设置写 exe 同目录 config.json。
"""
import json
import os
import re
import sys
import threading
import traceback

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from docx_parser import read_items, read_items_from_text, read_raw_paras, parse_sermon
from rule_parser import parse_by_rules
from pptx_builder import build_pptx
from platform_open import message_box, open_path

HERE = os.path.dirname(os.path.abspath(sys.argv[0]))


# ============================================================ 配置

DEFAULT_CONFIG = {
    "slide": {"width_in": 13.333, "height_in": 7.5, "content_ratio": 0.75},
    "background": {"image": "", "opacity": 100},
    "fonts": {"latin": "Calibri", "east_asian": "Microsoft YaHei"},
    "colors": {
        "title": "1F1F1F", "point": "1F1F1F", "sub": "1F1F1F",
        "verse_en": "595959", "ref": "7F7F7F", "highlight": "C00000",
    },
    "outline": {
        "size_zh": 15, "size_en": 13,
        "font_zh": "", "font_latin": "",
        "line_spacing": 1.15, "space_after_pt": 4,
        "bold_except_sub": True, "bold_sub": False,
    },
    "sizes": {
        "title_zh": 40, "title_en": 32,
        "sref_zh": 22, "sref_en": 18,
        "point_zh": 30, "point_en": 24,
        "sub_zh": 26, "sub_en": 20,
        "verse_zh": 24, "verse_en": 18,
    },
    "layout": {
        "line_spacing": 1.15, "space_after_pt": 8,
        "height_safety": 1.06, "cjk_width": 1.0,
        "latin_width": 0.55, "space_width": 0.33,
        "verse_indent_in": 0.3, "outline_sub_indent_in": 0.22,
        "margin_in": 0.45, "gap_in": 0.15,
    },
}

COLOR_LABELS = [  # (字段, 界面名)
    ("title", "题目"), ("point", "大点"), ("sub", "小点"),
    ("verse_en", "英文经文"), ("ref", "经文出处"), ("highlight", "提纲高亮"),
]
SIZE_LABELS = [
    ("title_zh", "题目中文"), ("title_en", "题目英文"),
    ("sref_zh", "主题经文中文"), ("sref_en", "主题经文英文"),
    ("point_zh", "大点中文"), ("point_en", "大点英文"),
    ("sub_zh", "小点中文"), ("sub_en", "小点英文"),
    ("verse_zh", "经文中文"), ("verse_en", "经文英文"),
]


def migrate_user_config(user_cfg):
    colors = user_cfg.get("colors")
    if isinstance(colors, dict):
        heading = colors.get("heading")
        if heading:
            colors.setdefault("title", heading)
            colors.setdefault("point", heading)
            colors.setdefault("sub", heading)


def normalize_colors(cfg):
    colors = cfg.get("colors", {})
    for k, v in list(colors.items()):
        if isinstance(v, str):
            colors[k] = v.strip().lstrip("#").upper()


def deep_merge(base, over):
    out = dict(base)
    for k, v in over.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def resolve_background(cfg):
    """生成前把 background.image 换算成实际路径（每次生成都重算，杜绝用到旧路径）。

    图片不存在时返回警告（生成继续，以无背景进行）。
    """
    bg = (cfg.get("background", {}).get("image") or "").strip()
    if not bg:
        cfg["background"]["resolved"] = ""
        return None
    p = bg if os.path.isabs(bg) else os.path.join(HERE, "background", bg)
    if os.path.isfile(p):
        cfg["background"]["resolved"] = p
        return None
    cfg["background"]["resolved"] = ""
    return f"背景图已不存在（可能被删除或移动）：{p}\n本次将以无背景生成。"


def config_path():
    return os.path.join(HERE, "config.json")


def load_config():
    cfg = json.loads(json.dumps(DEFAULT_CONFIG))  # deep copy
    note = ""
    if os.path.isfile(config_path()):
        try:
            with open(config_path(), "r", encoding="utf-8-sig") as f:
                user = json.load(f)
            migrate_user_config(user)
            cfg = deep_merge(cfg, user)
        except Exception as e:
            note = f"config.json 读取失败：{e}，已用默认设置"
    normalize_colors(cfg)
    # 背景图路径解析：相对名 → background/ 文件夹；绝对路径直接用
    bg = (cfg.get("background", {}).get("image") or "").strip()
    resolved = ""
    if bg:
        p = bg if os.path.isabs(bg) else os.path.join(HERE, "background", bg)
        if os.path.isfile(p):
            resolved = p
        else:
            note += f"背景图未找到：{p}"
    cfg.setdefault("background", {})["resolved"] = resolved
    return cfg, note


def save_config(cfg):
    out = json.loads(json.dumps(cfg))
    out.get("background", {}).pop("resolved", None)
    with open(config_path(), "w", encoding="utf-8-sig") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)


# ============================================================ 转换核心

def extract_marked_text(path):
    """AI 模式：docx → 带样式标记的文本（复用 sermon2text 核心）。"""
    from sermon2text import extract
    return extract(path)


def run_convert(items, skipped, dst, cfg):
    sermon, errors = parse_sermon(items, skipped)
    if errors:
        return None, errors
    return build_pptx(sermon, cfg, dst), []


def write_error_report(src_name, errors):
    path = os.path.join(HERE, "转换错误报告.txt")
    with open(path, "w", encoding="utf-8-sig") as f:
        f.write(f"来源：{src_name}\n共发现 {len(errors)} 个问题，已停止生成。\n")
        f.write("请修正后重试（AI 模式可把本报告贴回给 AI）。\n\n")
        for e in errors:
            f.write("● " + e + "\n")
    return path


def process_rules(src, dst, cfg):
    """规则模式：docx 保留原生样式；txt 纯文本。"""
    if src.lower().endswith(".docx"):
        raw = read_raw_paras(src)
    else:
        with open(src, "r", encoding="utf-8-sig") as f:
            raw = [(i, ln) for i, ln in enumerate(f, 1)]
    sermon, errors = parse_by_rules(raw)
    if errors:
        return None, errors
    return build_pptx(sermon, cfg, dst), []


def process_ai_text(text, dst, cfg):
    """AI 模式：粘贴的带 tag 文本。"""
    items, skipped = read_items_from_text(text)
    if not items:
        return None, ["粘贴的内容里没有任何 [TAG] 标记行。请确认复制的是 AI 加工后的完整回答。"]
    return run_convert(items, skipped, dst, cfg)


def process_ai_docx(src, dst, cfg):
    """AI 模式：直接喂带 tag 的 docx（兼容旧流程）。"""
    items, skipped = read_items(src)
    return run_convert(items, skipped, dst, cfg)


def process_background_image(src_path):
    """把所选背景图收进 background/ 文件夹（按内容指纹去重）。

    - 与已有图内容相同 → 直接用已有的，不重复存（即使文件名不同）
    - 同名但内容不同 → 存成 图片名_指纹前6位.png，绝不覆盖用户的图
    - 全新图片 → 原名复制
    返回 background/ 里的文件名。
    """
    import hashlib

    bdir = os.path.join(HERE, "background")
    os.makedirs(bdir, exist_ok=True)

    with open(src_path, "rb") as f:
        digest = hashlib.md5(f.read()).hexdigest()
    fingerprint = digest[:6]

    base, ext = os.path.splitext(os.path.basename(src_path))
    ext = ext.lower() or ".png"

    # 1) 内容已在库中（不管叫什么名）→ 直接用
    for name in os.listdir(bdir):
        fp = os.path.join(bdir, name)
        if not os.path.isfile(fp):
            continue
        try:
            with open(fp, "rb") as f:
                if hashlib.md5(f.read()).hexdigest() == digest:
                    return name
        except OSError:
            continue

    # 2) 同名但内容不同 → 加指纹后缀，不覆盖
    candidate = base + ext
    dst = os.path.join(bdir, candidate)
    if os.path.exists(dst):
        candidate = f"{base}_{fingerprint}{ext}"
        dst = os.path.join(bdir, candidate)

    # 3) 全新图片 → 原名（或指纹名）复制
    import shutil
    shutil.copy(src_path, dst)
    return os.path.basename(dst)


def open_background_folder():
    """打开 background/ 文件夹（资源管理器），返回是否成功。"""
    bdir = os.path.join(HERE, "background")
    os.makedirs(bdir, exist_ok=True)
    try:
        open_path(bdir)
        return True
    except Exception:
        return False


# ============================================================ 界面

class App:
    def __init__(self, root):
        self.root = root
        root.title("讲道稿转PPT")
        root.geometry("640x520")
        root.minsize(560, 460)
        self.cfg, _ = load_config()
        self.src_path = tk.StringVar()
        self.mode = tk.StringVar(value="rules")
        self._build()

    # ---------- 主界面 ----------
    def _build(self):
        pad = {"padx": 12, "pady": 6}
        f = ttk.Frame(self.root)
        f.pack(fill="both", expand=True)

        ttk.Label(f, text="① 选择讲道稿").grid(row=0, column=0, sticky="w", **pad)
        row1 = ttk.Frame(f)
        row1.grid(row=1, column=0, sticky="we", **pad)
        ttk.Entry(row1, textvariable=self.src_path).pack(side="left", fill="x", expand=True)
        ttk.Button(row1, text="浏览…", command=self.pick_file).pack(side="left", padx=(6, 0))
        self.root.drag_dest_register = None  # 占位说明：拖拽由命令行/双击替代

        ttk.Label(f, text="② 选择模式").grid(row=2, column=0, sticky="w", **pad)
        row2 = ttk.Frame(f)
        row2.grid(row=3, column=0, sticky="w", **pad)
        ttk.Radiobutton(row2, text="规则匹配（推荐，无需 AI）",
                        variable=self.mode, value="rules",
                        command=self.on_mode_change).pack(side="left")
        ttk.Radiobutton(row2, text="AI 加工（DeepSeek）",
                        variable=self.mode, value="ai",
                        command=self.on_mode_change).pack(side="left", padx=(16, 0))

        # AI 模式区
        self.ai_frame = ttk.Frame(f)
        self.ai_frame.grid(row=4, column=0, sticky="we", **pad)
        ttk.Label(self.ai_frame, text="第 1 步：复制下面的文本（连同提示词）发给 DeepSeek").pack(anchor="w")
        btns = ttk.Frame(self.ai_frame)
        btns.pack(anchor="w", pady=2)
        ttk.Button(btns, text="提取并复制", command=self.extract_copy).pack(side="left")
        ttk.Button(btns, text="复制提示词", command=self.copy_prompt).pack(side="left", padx=(6, 0))
        self.extract_box = tk.Text(self.ai_frame, height=5, wrap="char")
        self.extract_box.pack(fill="x", pady=(2, 4))
        ttk.Label(self.ai_frame, text="第 2 步：把 DeepSeek 的回答粘贴到这里").pack(anchor="w")
        self.ai_box = tk.Text(self.ai_frame, height=6, wrap="char")
        self.ai_box.pack(fill="x")

        ttk.Label(f, text="③ 生成").grid(row=5, column=0, sticky="w", **pad)
        self.gen_btn = ttk.Button(f, text="生成 PPT…", command=self.generate)
        self.gen_btn.grid(row=6, column=0, sticky="we", padx=12, pady=(0, 6))
        self.status = tk.StringVar(value="就绪")
        ttk.Label(f, textvariable=self.status, foreground="#666").grid(row=7, column=0, sticky="w", **pad)

        bottom = ttk.Frame(f)
        bottom.grid(row=8, column=0, sticky="we", **pad)
        ttk.Button(bottom, text="排版设置…", command=self.open_settings).pack(side="left")
        ttk.Button(bottom, text="使用说明", command=self.show_help).pack(side="left", padx=(8, 0))
        ttk.Label(bottom, text="提示：把 docx 拖到本程序图标上 = 规则模式直接转换",
                  foreground="#999").pack(side="right")

        f.columnconfigure(0, weight=1)
        f.rowconfigure(4, weight=0)
        self.on_mode_change()

    def on_mode_change(self):
        if self.mode.get() == "ai":
            self.ai_frame.grid()
            self.root.geometry("640x760")
        else:
            self.ai_frame.grid_remove()
            self.root.geometry("640x520")

    def pick_file(self):
        p = filedialog.askopenfilename(
            title="选择讲道稿",
            filetypes=[("Word/TXT", "*.docx *.txt"), ("所有文件", "*.*")])
        if p:
            self.src_path.set(p)

    # ---------- AI 模式动作 ----------
    def extract_copy(self):
        src = self.src_path.get().strip()
        if not src or not os.path.isfile(src):
            messagebox.showwarning("提示", "请先在上方选择讲道稿文件")
            return
        try:
            text = extract_marked_text(src)
        except Exception as e:
            messagebox.showerror("提取失败", str(e))
            return
        self.extract_box.delete("1.0", "end")
        self.extract_box.insert("1.0", text)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self.status.set("已提取并复制到剪贴板；备用文本在上面框内（可手动全选复制）")

    def copy_prompt(self):
        pf = os.path.join(HERE, "AI提示词模板.md")
        if not os.path.isfile(pf):
            messagebox.showwarning("提示", "找不到 AI提示词模板.md（需与程序同目录）")
            return
        with open(pf, "r", encoding="utf-8-sig") as f:
            content = f.read()
        # 截取提示词正文（第一个分隔线之后）
        marker = "## 提示词正文"
        idx = content.find(marker)
        if idx >= 0:
            content = content[idx:]
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.status.set("提示词已复制到剪贴板，粘贴到 DeepSeek 对话框即可")

    # ---------- 生成 ----------
    def generate(self):
        src = self.src_path.get().strip()
        if not src or not os.path.isfile(src):
            messagebox.showwarning("提示", "请先选择讲道稿文件")
            return
        if self.mode.get() == "rules":
            self._generate_rules(src)
        else:
            text = self.ai_box.get("1.0", "end").strip()
            if not text:
                messagebox.showwarning("提示", "请先把 DeepSeek 的回答粘贴到下方文本框")
                return
            self._generate_ai(text, src)

    def _pick_dst(self, src):
        base = os.path.splitext(os.path.basename(src))[0]
        p = filedialog.asksaveasfilename(
            title="保存 PPT", defaultextension=".pptx",
            initialfile=base + ".pptx",
            filetypes=[("PPT 演示文稿", "*.pptx")])
        return p or None

    def _generate_rules(self, src):
        dst = self._pick_dst(src)
        if not dst:
            return
        warn = resolve_background(self.cfg)  # 每次生成都重算背景路径，防止用到上一次的图
        if warn:
            messagebox.showwarning("背景图缺失", warn)
        self.status.set("正在生成…")
        self.root.update()
        try:
            n, errors = process_rules(src, dst, self.cfg)
        except Exception as e:
            self._crash(e)
            return
        self._after_convert(n, errors, dst, src)

    def _generate_ai(self, text, src):
        dst = self._pick_dst(src)
        if not dst:
            return
        warn = resolve_background(self.cfg)  # 每次生成都重算背景路径，防止用到上一次的图
        if warn:
            messagebox.showwarning("背景图缺失", warn)
        self.status.set("正在生成…")
        self.root.update()
        try:
            n, errors = process_ai_text(text, dst, self.cfg)
        except Exception as e:
            self._crash(e)
            return
        self._after_convert(n, errors, dst, "粘贴的AI回答")

    def _after_convert(self, n, errors, dst, src_name):
        if errors:
            report = write_error_report(src_name, errors)
            try:
                open_path(report)
            except Exception:
                pass
            messagebox.showerror("转换失败",
                                 f"发现 {len(errors)} 个问题，已停止生成。\n错误报告：{report}")
            self.status.set("转换失败，详见错误报告")
            return
        self.status.set(f"完成：{n} 页 → {dst}")
        if messagebox.askyesno("完成", f"已生成 {n} 页。\n现在打开文件吗？"):
            try:
                open_path(dst)
            except Exception:
                pass

    def _crash(self, e):
        log = os.path.join(HERE, "崩溃日志.txt")
        with open(log, "w", encoding="utf-8-sig") as f:
            f.write(traceback.format_exc())
        messagebox.showerror("程序错误", f"发生未预期的错误，详见：{log}")
        self.status.set("出错")

    def show_help(self):
        hp = os.path.join(HERE, "使用说明.md")
        if os.path.isfile(hp):
            open_path(hp)
        else:
            messagebox.showinfo("使用说明", "找不到使用说明.md")

    # ---------- 排版设置弹窗 ----------
    def open_settings(self):
        SettingsDialog(self.root, self.cfg, on_save=lambda: self.status.set("设置已保存"))


class SettingsDialog(tk.Toplevel):
    """config 全字段的可视化编辑器。"""

    def __init__(self, parent, cfg, on_save=None):
        super().__init__(parent)
        self.cfg = cfg
        self.on_save = on_save
        self.title("排版设置")
        self.geometry("560x640")
        self.transient(parent)
        self.grab_set()

        vars = {}
        self.vars = vars

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)

        # ---- 字号页 ----
        f1 = ttk.Frame(nb); nb.add(f1, text="字号")
        for i, (key, label) in enumerate(SIZE_LABELS):
            ttk.Label(f1, text=label).grid(row=i, column=0, sticky="w", padx=10, pady=3)
            v = tk.StringVar(value=str(cfg["sizes"][key]))
            vars[f"sizes.{key}"] = v
            ttk.Spinbox(f1, from_=8, to=96, textvariable=v, width=6).grid(row=i, column=1, padx=6)
        for i, (key, label) in enumerate([("size_zh", "提纲统一中文"), ("size_en", "提纲统一英文")]):
            r = len(SIZE_LABELS) + i
            ttk.Label(f1, text=label).grid(row=r, column=0, sticky="w", padx=10, pady=3)
            v = tk.StringVar(value=str(cfg["outline"][key]))
            vars[f"outline.{key}"] = v
            ttk.Spinbox(f1, from_=8, to=96, textvariable=v, width=6).grid(row=r, column=1, padx=6)

        # ---- 颜色页 ----
        f2 = ttk.Frame(nb); nb.add(f2, text="颜色")
        for i, (key, label) in enumerate(COLOR_LABELS):
            ttk.Label(f2, text=label).grid(row=i, column=0, sticky="w", padx=10, pady=3)
            v = tk.StringVar(value=cfg["colors"][key])
            vars[f"colors.{key}"] = v
            ttk.Entry(f2, textvariable=v, width=10).grid(row=i, column=1, padx=6)
            preview = tk.Label(f2, text="  ", width=6, background="#" + cfg["colors"][key])
            preview.grid(row=i, column=2, padx=6)
            v.trace_add("write", lambda *a, pv=preview, vv=v: self._update_preview(pv, vv))
            ttk.Button(f2, text="选色", width=5,
                       command=lambda k=key, vv=v: self._pick_color(vv)).grid(row=i, column=3)

        # ---- 字体/提纲页 ----
        f3 = ttk.Frame(nb); nb.add(f3, text="字体与提纲")
        rows = [
            ("fonts.latin", "左侧英文字体", cfg["fonts"]["latin"]),
            ("fonts.east_asian", "左侧中文字体", cfg["fonts"]["east_asian"]),
            ("outline.font_latin", "提纲英文字体(空=跟随左侧)", cfg["outline"]["font_latin"]),
            ("outline.font_zh", "提纲中文字体(空=跟随左侧)", cfg["outline"]["font_zh"]),
            ("outline.line_spacing", "提纲行距(倍)", cfg["outline"]["line_spacing"]),
            ("outline.space_after_pt", "提纲段后距(磅)", cfg["outline"]["space_after_pt"]),
            ("layout.line_spacing", "左侧行距(倍)", cfg["layout"]["line_spacing"]),
            ("layout.space_after_pt", "左侧段后距(磅)", cfg["layout"]["space_after_pt"]),
            ("layout.margin_in", "页边距(英寸)", cfg["layout"]["margin_in"]),
            ("layout.gap_in", "左右列间隙(英寸)", cfg["layout"]["gap_in"]),
            ("layout.verse_indent_in", "经文缩进(英寸)", cfg["layout"]["verse_indent_in"]),
            ("slide.content_ratio", "内容区宽度比例", cfg["slide"]["content_ratio"]),
        ]
        for i, (path, label, val) in enumerate(rows):
            ttk.Label(f3, text=label).grid(row=i, column=0, sticky="w", padx=10, pady=3)
            v = tk.StringVar(value=str(val))
            vars[path] = v
            ttk.Entry(f3, textvariable=v, width=12).grid(row=i, column=1, padx=6)
        self.bold_except_sub = tk.BooleanVar(value=cfg["outline"]["bold_except_sub"])
        self.bold_sub = tk.BooleanVar(value=cfg["outline"]["bold_sub"])
        ttk.Checkbutton(f3, text="提纲：小点以外加粗", variable=self.bold_except_sub).grid(
            row=len(rows), column=0, columnspan=2, sticky="w", padx=10, pady=3)
        ttk.Checkbutton(f3, text="提纲：小点也加粗", variable=self.bold_sub).grid(
            row=len(rows) + 1, column=0, columnspan=2, sticky="w", padx=10, pady=3)

        # ---- 背景图页 ----
        f4 = ttk.Frame(nb); nb.add(f4, text="背景图")
        self.bg_status = tk.StringVar(
            value=os.path.basename(cfg["background"].get("image") or "") or "未设置背景图")
        ttk.Label(f4, text="当前背景图：").grid(row=0, column=0, sticky="w", padx=10, pady=8)
        ttk.Label(f4, textvariable=self.bg_status, foreground="#060").grid(row=0, column=1, sticky="w")
        ttk.Button(f4, text="选择图片…", command=self._pick_bg).grid(row=1, column=0, padx=10, pady=4)
        ttk.Button(f4, text="打开背景文件夹", command=open_background_folder).grid(
            row=1, column=1, padx=10, pady=4, sticky="w")
        ttk.Button(f4, text="清除背景", command=self._clear_bg).grid(row=2, column=0, padx=10, pady=4, sticky="w")
        ttk.Label(f4, text="不透明度(%)").grid(row=2, column=0, sticky="w", padx=10, pady=8)
        self.opacity_var = tk.IntVar(value=int(cfg["background"].get("opacity", 100)))
        scale = ttk.Scale(f4, from_=10, to=100, variable=self.opacity_var,
                          command=lambda v: self.opacity_label.config(text=f"{self.opacity_var.get()}%"))
        scale.grid(row=3, column=0, columnspan=2, sticky="we", padx=10)
        self.opacity_label = ttk.Label(f4, text=f"{self.opacity_var.get()}%")
        self.opacity_label.grid(row=3, column=2)
        ttk.Label(f4, text="提示：图片会收进程序同目录 background/ 文件夹（相同内容的图只存一份）；"
                  "透明度越低文字越清晰",
                  foreground="#999", wraplength=420, justify="left").grid(
            row=5, column=0, columnspan=3, sticky="w", padx=10, pady=10)
        f4.columnconfigure(0, weight=1)

        # ---- 底部按钮 ----
        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btns, text="恢复默认", command=self._reset).pack(side="left")
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="保存", command=self._save).pack(side="right", padx=(0, 8))

    # ---- 辅助 ----
    def _update_preview(self, preview, var):
        v = var.get().strip().lstrip("#")
        if re.fullmatch(r"[0-9A-Fa-f]{6}", v):
            preview.config(background="#" + v)

    def _pick_color(self, var):
        from tkinter import colorchooser
        c = colorchooser.askcolor(parent=self)
        if c and c[1]:
            var.set("".join(f"{int(x):02X}" for x in c[0]))

    def _pick_bg(self):
        p = filedialog.askopenfilename(
            parent=self, title="选择背景图",
            filetypes=[("图片", "*.png *.jpg *.jpeg *.bmp"), ("所有文件", "*.*")])
        if not p:
            return
        try:
            name = process_background_image(p)
        except Exception as e:
            messagebox.showerror("复制失败", str(e), parent=self)
            return
        self.cfg["background"]["image"] = name
        self.bg_status.set(name)

    def _clear_bg(self):
        self.cfg["background"]["image"] = ""
        self.bg_status.set("未设置背景图")

    def _set_nested(self, path, value):
        parts = path.split(".")
        d = self.cfg
        for p in parts[:-1]:
            d = d[p]
        d[parts[-1]] = value

    def _save(self):
        try:
            text_fields = {"fonts.latin", "fonts.east_asian",
                           "outline.font_zh", "outline.font_latin"}
            for path, var in self.vars.items():
                raw = var.get().strip()
                if path in text_fields:
                    self._set_nested(path, raw)
                elif path.startswith("colors."):
                    self._set_nested(path, raw.lstrip("#").upper())
                else:
                    self._set_nested(path, float(raw) if "." in raw else int(raw))
            self.cfg["outline"]["bold_except_sub"] = self.bold_except_sub.get()
            self.cfg["outline"]["bold_sub"] = self.bold_sub.get()
            self.cfg["background"]["opacity"] = int(self.opacity_var.get())
            normalize_colors(self.cfg)
            save_config(self.cfg)
        except Exception as e:
            messagebox.showerror("保存失败", f"输入有误：{e}", parent=self)
            return
        if self.on_save:
            self.on_save()
        self.destroy()

    def _reset(self):
        if messagebox.askyesno("恢复默认", "所有设置恢复为默认值？（保存后生效）", parent=self):
            self.cfg = json.loads(json.dumps(DEFAULT_CONFIG))
            messagebox.showinfo("已重置", "界面数值已重置为默认，点「保存」写入 config.json", parent=self)
            self.destroy()
            SettingsDialog(self.root, self.cfg, self.on_save)


# ============================================================ 入口

def main():
    argv = sys.argv[1:]
    if argv:
        # 命令行/拖拽：默认规则模式直接转换
        src = argv[0]
        if not os.path.isfile(src):
            message_box(f"找不到输入文件：\n{src}", error=True)
            return 1
        cfg, note = load_config()
        warn = resolve_background(cfg)
        if warn:
            message_box(warn, error=False)
        dst = argv[1] if len(argv) > 1 else os.path.splitext(src)[0] + ".pptx"
        try:
            n, errors = process_rules(src, dst, cfg)
        except Exception:
            log = os.path.join(HERE, "崩溃日志.txt")
            with open(log, "w", encoding="utf-8-sig") as f:
                f.write(traceback.format_exc())
            message_box(f"发生错误，详见：{log}", error=True)
            return 3
        if errors:
            report = write_error_report(src, errors)
            try:
                open_path(report)
            except Exception:
                pass
            message_box(f"发现 {len(errors)} 个问题，已停止生成。\n{report}", error=True)
            return 2
        message_box(f"已生成 {n} 页：\n{dst}")
        return 0

    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
