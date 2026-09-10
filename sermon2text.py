# -*- coding: utf-8 -*-
"""讲道稿提文本工具：Word docx → 带样式标记的纯文本文件。

用法：
  双击    弹出文件选择框，选 Word 讲道稿 → 在原稿同目录生成
          「原稿名_给AI.txt」，弹窗告知位置；打开它全选复制给 DeepSeek 即可
  命令行  sermon2text.exe 输入.docx [输出.txt]

样式转换规则（供 DeepSeek 等 AI 识别并原样保留）：
  粗体   → **文字**
  下划线 → __文字__
  颜色   → {c:RRGGBB}文字{/c}   如 {c:C00000}…{/c}
"""
import os
import sys
import traceback

from docx import Document
from docx_parser import wrap
from win_dialogs import get_open_filename, message_box


def run_color(r):
    try:
        if r.font.color is not None and r.font.color.type is not None and r.font.color.rgb is not None:
            return str(r.font.color.rgb)
    except Exception:
        pass
    return None


def run_underline(r):
    u = r.underline
    return u is True or (u not in (None, False) and getattr(u, "name", "") != "NONE")


def extract(path):
    """读 docx，输出每段一行的纯文本（相邻同样式 run 合并）。"""
    doc = Document(path)
    lines = []
    for para in doc.paragraphs:
        parts = []
        for r in para.runs:
            if not r.text:
                continue
            seg = (bool(r.font.bold), run_underline(r), run_color(r))
            # 相邻同样式的 run 合并成一个标记段
            if parts and parts[-1][0] == seg:
                parts[-1][1] += r.text
            else:
                parts.append([seg, r.text])
        line = "".join(wrap(t, bold=b, underline=u, color=c)
                       for (b, u, c), t in parts)
        if line.strip():
            lines.append(line)
    return "\r\n".join(lines)


def process(src, dst=None):
    text = extract(src)
    if not text.strip():
        return None, "文档是空的，没有提取到任何内容"
    out = dst or (os.path.splitext(src)[0] + "_给AI.txt")
    with open(out, "w", encoding="utf-8-sig") as f:
        f.write(text)
    return out, ""


def main():
    argv = sys.argv[1:]
    silent = bool(os.environ.get("S2P_SILENT"))
    try:
        if argv:
            src = argv[0]
            if not os.path.isfile(src):
                if not silent:
                    message_box(f"找不到文件：\n{src}", error=True)
                return 1
            backup, _ = process(src, argv[1] if len(argv) > 1 else None)
            print(f"已生成：{backup}")
            return 0 if backup else 2
        src = get_open_filename("选择讲道稿 Word 文档",
                                [("Word 文档", "*.docx"), ("所有文件", "*.*")])
        if not src:
            return 1
        out, note = process(src)
        if out is None:
            if not silent:
                message_box(note, error=True)
            return 2
        if not silent:
            message_box(f"已生成文本文件：\n{out}\n\n"
                        f"下一步：打开它，全选复制，粘贴到 DeepSeek\n"
                        f"（配合「AI提示词模板」使用），\n"
                        f"再把 AI 的回答复制，双击 sermon2ppt.exe 生成 PPT。")
        print(f"已生成：{out}")
        return 0
    except Exception:
        err = traceback.format_exc()
        try:
            log = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])),
                               "sermon2text_崩溃日志.txt")
            with open(log, "w", encoding="utf-8-sig") as f:
                f.write(err)
        except Exception:
            log = None
        if not silent:
            message_box("发生未预期的错误" + (f"，详见：{log}" if log else "：\n" + err),
                        error=True)
        print(err)
        return 3


if __name__ == "__main__":
    sys.exit(main())
