#!/usr/bin/env python3
"""
把操作手册 Markdown 转成格式规范的 Word 文档（封面 + 目录 + 正文，页眉"全称+版本号"，右上角页码）。

用法:
  uv run --with python-docx python3 build_manual_docx.py 草稿/操作手册.md \
      --name "XX系统" --version V1.0 --owner "XX科技有限公司" -o 正式资料/
  可选: --heading-style cn|num  标题编号风格（R011）：num=阿拉伯（默认，1 2 3）；cn=中文数字（一、二、三）
  可选: --static-toc 静态标题列表目录（LibreOffice 直接转 PDF 也能看）

Markdown 约定（只支持这些，够用即可）:
  # 1 系统登录            → 一级标题（Heading 1）
  ## 1.1 访问系统          → 二级标题（Heading 2）
  ### 小标题               → 三级标题
  普通段落                 → 正文，宋体 12pt，**顶格**（无首行缩进），1.5 倍行距
  1. / - 开头的行          → 保留编号/项目符号文字，按正文段落排
  ![图 1-1 登录页面](img/login.png)  → 居中插图（最宽 15cm）+ 下方居中图注（图注即 [] 内文字，必须写）
  【截图预留：请插入「登录页面」截图。】 → 醒目占位段（红色加粗），交付前必须替换
  | a | b |  表格（管道表，第二行为分隔行）→ Word 表格，宋体 10.5pt
  > 引用 → 按正文处理
  空行 → 段落分隔
  正文默认顶格（用户要求：内容顶格写，不要空格）；如旧样式需要首行缩进 2 字符，加 --indent

页面: A4，上下 2.54 / 左右 3.175 cm；每页正文 ≥30 行由 1.5 倍行距 12pt 保证（约 38 行/页）。
目录: 插入 Word TOC 域并设置打开时更新；用 Word 打开一次 → 更新域 → 保存后再导出 PDF。
      加 --static-toc 则改为静态标题列表（无页码，LibreOffice 直接转 PDF 也能看）。
"""
import argparse, re, sys
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx_common import set_font, set_style_font, setup_header, setup_page, add_field, enable_update_fields, remove_doc_grid, check_locked

IMG_RE = re.compile(r"^!\[(?P<cap>[^\]]*)\]\((?P<src>[^)]+)\)\s*$")
PLACEHOLDER_RE = re.compile(r"^【截图预留[:：].*】\s*$")
HEAD_RE = re.compile(r"^(#{1,3})\s+(.*)$")

CN_DIGIT = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
def cn_num(n: int) -> str:
    if n < 10: return CN_DIGIT[n]
    if n < 20: return "十" + CN_DIGIT[n - 10]
    return CN_DIGIT[n // 10] + "十" + (CN_DIGIT[n % 10] if n % 10 else "")

class HeadingStyle:
    """标题编号风格：num=阿拉伯（默认，1 / 1.1）；cn=中文数字一级 + 阿拉伯二级（对齐用户成品样式：一、首页 / 1. 用户管理）。"""
    def __init__(self, style: str):
        self.style = style
        self.h1 = 0

    def convert(self, level: int, text: str) -> str:
        if self.style != "cn": return text
        m = re.match(r"^(\d+)(?:\.(\d+))?(?:\.\d+)*\s+(.*)$", text)
        if not m: return text
        a = int(m.group(1))
        b = int(m.group(2)) if m.group(2) else None
        if b is None:
            self.h1 = a
            return f"{cn_num(a)}、{m.group(3)}"
        return f"{b}. {m.group(3)}"

LINE_SPACING = 1.5  # 对齐用户手册：12pt 正文 1.5 倍行距

def body_para(doc, text, font, size=12, indent=False, align=WD_ALIGN_PARAGRAPH.LEFT):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE; pf.line_spacing = LINE_SPACING
    pf.space_before = Pt(0); pf.space_after = Pt(0)
    if indent: pf.first_line_indent = Pt(size * 2)
    if align is not None: p.alignment = align
    # 简单行内加粗 **x**
    parts = re.split(r"(\*\*[^*]+\*\*)", text)
    for part in parts:
        if not part: continue
        bold = part.startswith("**") and part.endswith("**")
        r = p.add_run(part[2:-2] if bold else part); set_font(r, font, size, bold=bold or None)
    return p

def add_table(doc, rows, font):
    ncol = max(len(r) for r in rows)
    t = doc.add_table(rows=len(rows), cols=ncol); t.style = "Table Grid"
    for i, row in enumerate(rows):
        for j in range(ncol):
            cell = t.cell(i, j); cell.text = ""
            r = cell.paragraphs[0].add_run(row[j] if j < len(row) else ""); set_font(r, font, 10.5, bold=True if i == 0 else None)
    doc.add_paragraph()

def add_image(doc, src: Path, caption: str, font, max_w_cm=15.0):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.add_run().add_picture(str(src), width=Cm(max_w_cm))
    if caption:  # 图注可选：用户手册无图注时（alt 为空）不生成
        c = doc.add_paragraph(); c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.paragraph_format.space_after = Pt(6)
        r = c.add_run(caption); set_font(r, font, 10.5)

def add_placeholder(doc, text, font):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text); set_font(r, font, 12, bold=True); r.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

def add_cover(doc, name, version, owner, font):
    """封面（对齐用户手册样式）：名称 + V1.0 + 用户手册，黑体 24pt，无著作权人行。"""
    def center_line(text, size, bold=False, ascii_name=None):
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.space_before = Pt(0); pf.space_after = Pt(0)
        r = p.add_run(text); set_font(r, font, size, bold=bold or None, ascii_name=ascii_name)
        return p

    def blank(n):
        for _ in range(n):
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
            pf.space_before = Pt(0); pf.space_after = Pt(0)

    blank(6)
    center_line(name, 24, bold=True, ascii_name="黑体")
    center_line(version, 24, ascii_name="Calibri")
    blank(17)
    center_line("用户手册", 24, bold=True, ascii_name="黑体")
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

def add_toc(doc, headings, font, static: bool):
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run("目录"); set_font(r, "宋体", 10.5)
    if static:
        for level, text in headings:
            q = doc.add_paragraph(); q.paragraph_format.left_indent = Pt(24 * (level - 1))
            q.paragraph_format.line_spacing = 1.5
            r = q.add_run(text); set_font(r, font, 12)
    else:
        q = doc.add_paragraph()
        for r in add_field(q, 'TOC \\o "1-3" \\h \\z \\u', "（请在 Word 中右键此处 → 更新域 生成目录）"): set_font(r, font, 12)
        enable_update_fields(doc)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("markdown")
    ap.add_argument("--name", required=True); ap.add_argument("--version", required=True); ap.add_argument("--owner", required=True)
    ap.add_argument("-o", "--outdir", default=".")
    ap.add_argument("--font", default="宋体")
    ap.add_argument("--static-toc", action="store_true")
    ap.add_argument("--heading-style", choices=["num", "cn"], default="num",
                    help="标题编号风格（R011）：num=阿拉伯数字（默认，1 2 3）；cn=中文数字（一、二、三）")
    ap.add_argument("--line-spacing", type=float, default=None, help="正文行距倍数，默认 1.5")
    ap.add_argument("--indent", action="store_true", help="正文首行缩进 2 字符（默认顶格，不加此开关）")
    ap.add_argument("--img-root", default=None, help="图片相对路径的基准目录，默认为 markdown 所在目录")
    a = ap.parse_args()

    md_path = Path(a.markdown); img_root = Path(a.img_root) if a.img_root else md_path.parent
    lines = md_path.read_text(encoding="utf-8").splitlines()
    # 跳过 YAML front matter
    if lines and lines[0].strip() == "---":
        try: lines = lines[lines.index("---", 1) + 1:]
        except ValueError: pass

    global LINE_SPACING
    if a.line_spacing: LINE_SPACING = a.line_spacing
    font = a.font
    doc = Document()
    set_style_font(doc.styles["Normal"], font, 12)
    for lvl, size in ((1, 22), (2, 18), (3, 13.5)):
        st = doc.styles[f"Heading {lvl}"]; set_style_font(st, font, size)
        st.font.bold = True; st.font.color.rgb = RGBColor(0, 0, 0)
        if lvl == 1:
            st.paragraph_format.line_spacing = 1.2
            st.paragraph_format.space_before = Pt(17); st.paragraph_format.space_after = Pt(16)
        else:
            st.paragraph_format.line_spacing = 1.5
            st.paragraph_format.space_before = Pt(6); st.paragraph_format.space_after = Pt(6)
    sec = doc.sections[0]
    setup_page(sec, 2.54, 2.54, 3.175, 3.175, header_dist=1.0, footer_dist=0.8)
    remove_doc_grid(sec)
    setup_header(sec, f"{a.name}{a.version}", font_size=10.5, font=font)

    hs = HeadingStyle(a.heading_style)
    headings = [(len(m.group(1)), hs.convert(len(m.group(1)), m.group(2).strip())) for l in lines if (m := HEAD_RE.match(l))]
    add_cover(doc, a.name, a.version, a.owner, font)
    add_toc(doc, headings, font, a.static_toc)

    stats = {"headings": len(headings), "images": 0, "placeholders": 0, "paragraphs": 0, "tables": 0, "missing_images": []}
    i = 0; n = len(lines)
    while i < n:
        line = lines[i].rstrip()
        if not line.strip():
            i += 1; continue
        if m := HEAD_RE.match(line):
            h = doc.add_heading(hs.convert(len(m.group(1)), m.group(2).strip()), level=len(m.group(1)))
            for r in h.runs: set_font(r, font, {1: 22, 2: 18, 3: 13.5}[len(m.group(1))], bold=True); r.font.color.rgb = RGBColor(0, 0, 0)
            i += 1; continue
        if m := IMG_RE.match(line):
            src = (img_root / m.group("src")).resolve(); cap = m.group("cap").strip()
            if src.exists():
                add_image(doc, src, cap, font); stats["images"] += 1
            else:
                add_placeholder(doc, f"【截图缺失：{m.group('src')} —— {cap}】", font); stats["missing_images"].append(m.group("src"))
            i += 1; continue
        if PLACEHOLDER_RE.match(line):
            add_placeholder(doc, line.strip(), font); stats["placeholders"] += 1
            i += 1; continue
        if line.lstrip().startswith("|"):
            rows = []
            while i < n and lines[i].lstrip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells): rows.append(cells)
                i += 1
            add_table(doc, rows, font); stats["tables"] += 1; continue
        # 普通段落：合并到下一个空行/特殊行
        buf = [line.strip().lstrip("> ").rstrip()]
        i += 1
        while i < n and lines[i].strip() and not HEAD_RE.match(lines[i]) and not IMG_RE.match(lines[i]) \
                and not PLACEHOLDER_RE.match(lines[i]) and not lines[i].lstrip().startswith("|") \
                and not re.match(r"^\s*(\d+[.、)]|[-*•])\s", lines[i]):
            buf.append(lines[i].strip()); i += 1
        text = "".join(buf)
        is_list = bool(re.match(r"^(\d+[.、)]|[-*•])\s", text))
        body_para(doc, re.sub(r"^[-*•]\s+", "", text) if is_list else text, font, indent=a.indent and not is_list)
        stats["paragraphs"] += 1

    outdir = Path(a.outdir); outdir.mkdir(parents=True, exist_ok=True)
    out = outdir / f"{a.name}_操作手册.docx"
    locked = check_locked(out)
    if locked:
        sys.exit(f"⚠️  {out.name} 正被 Word/WPS 打开（{locked}）。请先关闭该文档，再重新运行。"
                  + " 也可先生成到其它目录（-o 临时目录）后再替换正式稿。")
    doc.save(out)
    print(f"✅ {out}")
    print(f"标题 {stats['headings']}，段落 {stats['paragraphs']}，插图 {stats['images']}，表格 {stats['tables']}，截图占位 {stats['placeholders']}（标题风格：{a.heading_style}）")
    if stats["placeholders"]: print(f"⚠️  仍有 {stats['placeholders']} 处【截图预留】，正式稿前必须替换为真实截图")
    if stats["missing_images"]: print("⚠️  图片文件不存在:", ", ".join(stats["missing_images"]))
    if not a.static_toc: print("ℹ️  目录为 Word 域：用 Word 打开 → 全选 → F9 更新域 → 保存，再导出 PDF")

if __name__ == "__main__":
    main()
