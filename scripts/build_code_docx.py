#!/usr/bin/env python3
"""
按格式生成源代码鉴别材料 Word 文档（R006：封面 + 目录页；改进：全量收录 + 目录按手册章节）。

用法:
  uv run --with python-docx python3 build_code_docx.py \
      --root <项目根目录> --files 草稿/代码文件顺序.txt \
      --name "XX系统" --version V1.0 -o 正式资料/ [--owner "XX公司"] \
      [--chapter-map 草稿/章节代码映射.md]

代码文件顺序.txt: 一行一个相对路径，按抽取顺序排列（第 1 个必须是入口文件）。
  - 以 # 开头的行作为模块边界标记（无 --chapter-map 时目录页的依据）：`# 模块名`。
  - 例：
      # 程序入口
      src/main/java/com/example/App.java
      # 登录鉴权模块
      ...

章节代码映射.md（可选，推荐）: 代码目录页按**用户手册章节**组织（用户要求：代码目录按手册目录）。
  - 每行 `章节名-----起始文件相对路径`（点线分隔，AI 按菜单/模块名匹配后由用户确认）。
  - 例：
      首页--------------------src/main/java/com/example/App.java
      系统管理------------------src/main/java/com/example/controller/AuthController.java
  - 脚本按起始文件在流水线中的位置自动计算页码，输出"章节名 … 第 N 页"目录页。

规则（内置，见 references/格式要求.md）:
  - 去空行、去纯注释行（--keep-comments 可保留注释）；Tab 转 4 空格
  - --strip-vue-style 去掉 .vue 的 <style> 块
  - 每页严格 50 行，每页前强制分页；页眉左"全称+版本号"，右"第 N 页"
  - **代码有多少写多少：默认收录全部代码，不限页数上限**；页数下限 60 页（不足时打印警告）
    （原"前 30 + 后 30 截断"已移除；--max-pages 为可选参数，仅特殊场景使用）
  - 目录页列出章节（或模块）与起始页码；正文页码从 1 重新编号（封面/目录不计入）
  - 末页不足 50 行：页底标注"本页为源代码最后一页，共 XX 行"
  - A4，上 1.8 / 下 1.6 / 左右 2.4 cm，代码西文 **Consolas 五号（10.5pt）**、中文宋体，固定行距 12.5pt（50 行×12.5=625pt < 745pt 可用高度，每页严格 50 行）
输出:
  <全称>_软件代码.docx 、 代码文档索引.md（每页对应的文件与原始行号，供人工核对）
"""
import argparse, re, sys
from pathlib import Path
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_LINE_SPACING, WD_ALIGN_PARAGRAPH, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.section import WD_SECTION
from docx_common import (set_font, set_style_font, set_code_font, setup_header, setup_page,
                         page_break_before, remove_doc_grid, set_page_number_start, check_locked)

LINES_PER_PAGE = 50
MIN_PAGES = 60  # 页数下限（用户要求：不得少于 60 页）
BLOCK_COMMENT = {"/*": "*/", "<!--": "-->", '"""': '"""', "'''": "'''"}

CN_NUM = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
def cn(n: int) -> str:
    if n < 10: return CN_NUM[n]
    if n < 20: return "十" + CN_NUM[n - 10]
    return CN_NUM[n // 10] + "十" + (CN_NUM[n % 10] if n % 10 else "")

def clean_lines(text: str, keep_comments: bool, strip_vue_style: bool = False):
    """返回 [(原始行号, 文本)]"""
    out = []; in_block = None; in_style = False
    for i, raw in enumerate(text.splitlines(), 1):
        line = raw.replace("\t", "    ").rstrip()
        s = line.strip()
        if not s: continue
        if strip_vue_style:
            if s.startswith("<style"): in_style = True
            if in_style:
                if s.startswith("</style>"): in_style = False
                continue
        if not keep_comments:
            if in_block:
                if in_block in s: in_block = None
                continue
            if s.startswith(("//", "#", "*", "--")) and not s.startswith("#!"): continue
            opened = False
            for o, c in BLOCK_COMMENT.items():
                if s.startswith(o):
                    if c not in s[len(o):]: in_block = c
                    opened = True; break
            if opened: continue
        out.append((i, line))
    return out

def parse_chapter_map(path: Path) -> list[tuple[int, str, str]]:
    """解析章节代码映射：`章节名-----起始文件路径`。返回 [(层级, 章节名, 起始文件相对路径)]。

    层级判定：行首带缩进（空格/制表符）的为二级条目（对齐手册目录 L2），否则为一级。
    """
    entries = []
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        s = raw.strip()
        if not s or s.startswith("#"): continue
        m = re.match(r"^(.*?)-{3,}\s*(\S+.*)$", s)
        if not m:
            print(f"⚠️  章节映射行无法解析（已跳过）：{s}")
            continue
        level = 2 if len(raw) - len(raw.lstrip()) > 0 else 1
        entries.append((level, m.group(1).strip(), m.group(2).strip()))
    return entries

def add_cover(doc, name, version, owner, font):
    """封面（对齐用户手册样式）：名称 + V1.0 + 源代码，黑体 24pt。

    关键修复：Normal 样式为代码页设置了固定行距（EXACTLY 11pt），封面段落必须显式
    改为单倍行距，否则大字号被裁切、行间重叠。
    """
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
    center_line("源代码", 24, bold=True, ascii_name="黑体")
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

def add_toc(doc, entries, font, text_width):
    """静态目录页（视觉对齐用户手册目录）。entries: [(层级, 名称, 页码)]。

    关键修复：改用右对齐制表位 + 点线前导符（dotted leader），每条严格一行、
    页码右对齐，不再用字面"…"凑长度（中文双宽导致换行错乱）。
    """
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_before = Pt(0); pf.space_after = Pt(6)
    r = p.add_run("目录"); set_font(r, "宋体", 10.5)

    for level, name, page in entries:
        q = doc.add_paragraph()
        qf = q.paragraph_format
        qf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE; qf.line_spacing = 1.5
        qf.space_before = Pt(0); qf.space_after = Pt(0)
        if level == 2:
            qf.left_indent = Pt(21)  # 二级缩进 2 字符（对齐手册 toc2）
        qf.tab_stops.add_tab_stop(text_width, WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
        r = q.add_run(name); set_font(r, "宋体", 10.5)
        r = q.add_run("\t"); set_font(r, "宋体", 10.5)
        r = q.add_run(f"第 {page} 页"); set_font(r, "宋体", 10.5)
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--files", required=True)
    ap.add_argument("--name", required=True, help="软件全称")
    ap.add_argument("--version", required=True, help="版本号，如 V1.0")
    ap.add_argument("--owner", default=None, help="著作权人（可选，用于封面）")
    ap.add_argument("-o", "--outdir", default=".")
    ap.add_argument("--keep-comments", action="store_true")
    ap.add_argument("--strip-vue-style", action="store_true", help="去掉 .vue 文件中的 <style> 块")
    ap.add_argument("--chapter-map", default=None, help="章节代码映射文件（目录页按用户手册章节），每行 `章节名-----起始文件路径`")
    ap.add_argument("--max-pages", type=int, default=None, help="可选：仅输出前 N 页（特殊场景，默认全量收录）")
    ap.add_argument("--font", default="宋体", help="代码中文字体（默认宋体）")
    ap.add_argument("--code-font", default="Consolas", help="代码西文字体（默认 Consolas）")
    ap.add_argument("--font-size", type=float, default=10.5, help="代码字号（默认 10.5pt = 五号）")
    a = ap.parse_args()

    root = Path(a.root).resolve()
    raw_lines = Path(a.files).read_text(encoding="utf-8-sig").splitlines()
    paths = [s.strip() for s in raw_lines if s.strip() and not s.lstrip().startswith("#")]
    if not paths:
        sys.exit("文件清单为空")

    # 拼接所有代码行 + 记录模块边界 + 记录每文件起始行
    all_lines = []
    per_file = []          # (rel, 有效行数)
    file_start = {}        # rel -> 起始行号（用于章节映射页码）
    modules2 = []          # (name, start_line, end_line)
    cur_name = "程序入口"; line_acc = 0; mod_start = 0; mod_files = []
    for l in raw_lines:
        s = l.strip()
        if not s: continue
        if s.startswith("#"):
            if mod_files:
                modules2.append((cur_name, mod_start, line_acc))
            cur_name = s.lstrip("#").strip() or "程序入口"
            mod_start = line_acc; mod_files = []
        else:
            p = root / s
            if not p.exists():
                sys.exit(f"文件不存在: {p}")
            file_start[s] = line_acc
            lines = clean_lines(p.read_text(encoding="utf-8", errors="replace"), a.keep_comments,
                                a.strip_vue_style and p.suffix.lower() == ".vue")
            per_file.append((s, len(lines)))
            all_lines.extend((s, n, t, cur_name) for n, t in lines)
            line_acc += len(lines)
            mod_files.append(s)
    if mod_files:
        modules2.append((cur_name, mod_start, line_acc))
    for i in range(len(modules2) - 1):
        modules2[i] = (modules2[i][0], modules2[i][1], modules2[i + 1][1])
    if modules2:
        modules2[-1] = (modules2[-1][0], modules2[-1][1], line_acc)

    total = len(all_lines)
    pages = [all_lines[i:i + LINES_PER_PAGE] for i in range(0, total, LINES_PER_PAGE)]
    n_pages = len(pages)

    # 默认全量收录（用户要求：代码有多少写多少）；--max-pages 仅特殊场景
    if a.max_pages and a.max_pages < n_pages:
        selected = pages[:a.max_pages]
        mode = f"前 {a.max_pages} 页（原文共 {n_pages} 页，--max-pages 截断）"
    else:
        selected = pages
        mode = f"全部 {n_pages} 页"

    # 目录条目：优先章节映射（按用户手册目录，含一/二级），否则回退模块边界
    toc_entries = []   # (层级, 名称, 起始文件相对路径或 None, 页码)
    if a.chapter_map:
        for level, chapter, rel in parse_chapter_map(Path(a.chapter_map)):
            if rel not in file_start:
                print(f"⚠️  章节「{chapter}」的起始文件不在选中清单中（已跳过）：{rel}")
                continue
            toc_entries.append((level, chapter, rel, file_start[rel] // LINES_PER_PAGE + 1))
        if not toc_entries:
            sys.exit("章节代码映射解析结果为空，请检查映射文件")
    else:
        toc_entries = [(1, name, None, ms // LINES_PER_PAGE + 1) for name, ms, _ in modules2]

    header_text = f"{a.name}{a.version}"
    doc = Document()
    set_style_font(doc.styles["Normal"], a.font, a.font_size)
    pf = doc.styles["Normal"].paragraph_format
    pf.space_before = Pt(0); pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(a.font_size + 2)  # 五号 10.5pt → 固定 12.5pt，保证每页 50 行装得下

    sec0 = doc.sections[0]
    setup_page(sec0, 1.8, 1.6, 2.4, 2.4)
    remove_doc_grid(sec0)
    text_width = sec0.page_width - sec0.left_margin - sec0.right_margin
    add_cover(doc, a.name, a.version, a.owner, a.font)
    add_toc(doc, [(lvl, n, p) for lvl, n, _, p in toc_entries], a.font, text_width)

    # 代码正文节：页眉 + 页码从 1 重新编号（R006）
    sec = doc.add_section(WD_SECTION.NEW_PAGE)
    setup_page(sec, 1.8, 1.6, 2.4, 2.4)
    remove_doc_grid(sec)
    setup_header(sec, header_text, font_size=a.font_size, font=a.font)
    set_page_number_start(sec, 1)

    index = ["# 代码文档索引", "", f"软件：{header_text}", f"抽取方式：{mode}", f"有效代码总行数：{total}", "",
             "## 目录（按用户手册章节 / 模块）与页码", "", "| 层级 | 名称 | 起始文件 | 起始行 | 起始页 |", "|---|---|---|---|---:|"]
    index += [f"| {level} | {name} | {rel or '—'} | {file_start[rel] if rel else ms} | {page} |" for level, name, rel, page in toc_entries]
    index += ["", "## 模块边界（# 注释行）", "", "| 模块 | 起始行 | 起始页 |", "|---|---|---:|"]
    index += [f"| {name} | {ms} | {ms // LINES_PER_PAGE + 1} |" for name, ms, _ in modules2]
    index += ["", "## 文件顺序与有效行数", "", "| 序 | 文件 | 有效行 |", "|---|---|---:|"]
    index += [f"| {i+1} | `{rel}` | {n} |" for i, (rel, n) in enumerate(per_file)]
    index += ["", "## 页 → 文件/原始行号", "", "| 页码 | 起 | 止 |", "|---|---|---|"]

    for pi, page in enumerate(selected, 1):
        first = True
        for rel, n, text, mod in page:
            p = doc.add_paragraph()
            if first and pi > 1:
                page_break_before(p)
            first = False
            r = p.add_run(text); set_code_font(r, a.font, a.font_size, a.code_font)
        s_rel, s_n, _, _ = page[0]; e_rel, e_n, _, _ = page[-1]
        index.append(f"| {pi} | `{s_rel}`:{s_n} | `{e_rel}`:{e_n} |")
        if pi == len(selected) and len(page) < LINES_PER_PAGE:
            p = doc.add_paragraph()
            r = p.add_run(f"本页为源代码最后一页，共 {len(page)} 行；以下无代码。"); set_font(r, a.font, a.font_size, bold=True)
    outdir = Path(a.outdir); outdir.mkdir(parents=True, exist_ok=True)
    out = outdir / f"{a.name}_软件代码.docx"
    locked = check_locked(out)
    if locked:
        sys.exit(f"⚠️  {out.name} 正被 Word/WPS 打开（{locked}）。请先关闭该文档，再重新运行。"
                  + " 也可先生成到其它目录（-o 临时目录）后再替换正式稿。")
    doc.save(out)
    (outdir / "代码文档索引.md").write_text("\n".join(index) + "\n", encoding="utf-8")

    print(f"✅ {out}")
    print(f"✅ {outdir / '代码文档索引.md'}")
    print(f"有效行数 {total}，{mode}，输出 {len(selected)} 页（封面/目录不计入页码）")
    print(f"目录条目 {len(toc_entries)}：" + ("，".join(f"{n}(第{p}页)" for _, n, _, p in toc_entries[:12])) + (" …" if len(toc_entries) > 12 else ""))
    print(f"首页起始：{selected[0][0][0]}:{selected[0][0][1]}  → 请确认是程序入口")
    print(f"末页结束：{selected[-1][-1][0]}:{selected[-1][-1][1]}  → 请确认是完整模块收尾")
    if n_pages < MIN_PAGES:
        print(f"⚠️  仅 {n_pages} 页 < {MIN_PAGES} 页下限：代码总量不足 60 页（真实项目一般远超；若确实如此请确认是否符合版权中心要求）")

if __name__ == "__main__":
    main()
