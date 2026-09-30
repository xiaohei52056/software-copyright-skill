#!/usr/bin/env python3
"""
提交前自检：三份材料的名称/版本/权利人一致性 + 各项硬性指标。

用法:
  uv run --with python-docx --with pyyaml python3 check_consistency.py 正式资料/ \
      [--yaml 草稿/申请表信息.yaml] [--screenshots-dir 草稿/截图/] [--screenshot-name "截图中的系统名称"]

检查项:
  1. 申请表 / 手册封面+页眉 / 代码页眉 中的 全称+版本号 完全一致
  2. 手册封面著作权人 == 申请表权利人名称
  3. 申请表: 主要功能（详版）500–1300 字，技术特点 ≤100 字，必填项非空
  4. 手册: 无【截图预留】/【截图缺失】残留；一级标题数、插图数；正文段落估算页数 ≥20（低于时警告，用户可选精简文案接受补正风险）；正文顶格（无首行缩进）
  5. 手册一级标题（模块名）是否在申请表"主要功能"中被提到
  6. 代码: 每页 50 行、总页数 ≤60、末页标注；代码文档索引中的语言 ⊆ 申请表"编程语言"
  7. 截图对应性（R015，--screenshots-dir）：截图数 ≤ 手册插图数；插图前后有文字；截图名不一致→警告
  8. 截图系统名称（R013，--screenshot-name）：与申报全称不一致时输出**警告**而非错误，不阻断流程

输出分级: ❌ 错误（必须修复） / ⚠️ 警告（人工裁决，不阻断） / ℹ️ 信息
"""
import argparse, re, sys
from pathlib import Path
from docx import Document
import yaml

EXT_LANG = {".java": "Java", ".py": "Python", ".js": "JavaScript", ".jsx": "JavaScript", ".ts": "TypeScript", ".tsx": "TypeScript",
            ".vue": "Vue", ".html": "HTML", ".css": "CSS", ".scss": "CSS", ".less": "CSS", ".go": "Go", ".cs": "C#", ".php": "PHP", ".kt": "Kotlin", ".sql": "SQL"}
VUE_IMPLIES = {"JavaScript", "HTML", "CSS"}
IMG_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}

def header_text(doc):
    """取文档中含页眉内容的节的页眉（R006：代码文档封面/目录节无页眉，正文节才有）"""
    for sec in doc.sections:
        t = "".join(p.text for p in sec.header.paragraphs).strip()
        if t:
            return t.split("\t")[0].strip()
    return ""

def cjk_len(s): return len(re.sub(r"\s", "", s or ""))

def split_pages(doc):
    """按分页符切分段落页。识别 pageBreakBefore 与 run 级 <w:br type=page>。"""
    from docx.oxml.ns import qn
    pages, cur = [], []
    for p in doc.paragraphs:
        is_break = False
        pPr = p._element.pPr
        if pPr is not None and pPr.find(qn("w:pageBreakBefore")) is not None:
            is_break = True
        if not is_break:
            for r in p.runs:
                for br in r._element.findall(qn("w:br")):
                    if br.get(qn("w:type")) == "page":
                        is_break = True; break
                if is_break: break
        if is_break and cur:
            pages.append(cur); cur = []
        if p.text.strip(): cur.append(p.text)
    if cur: pages.append(cur)
    return pages

def looks_like_front(pg, name=None, version=None):
    """封面/目录页特征（不计入代码页统计）。

    需精确区分：封面（名称+版本+文档类型，新封面无著作权人行）、目录（"目录"标题+多条"第 N 页"）。
    不能仅凭某代码页正文中出现"第 N 页"字样就误判为前置页（旧逻辑会漏掉大量代码页）。
    """
    joined = "\n".join(pg)
    n = len(pg)
    # 目录页：含"目录"标题，且至少两条"第 N 页"
    if re.search(r"目\s*录", joined) and len(re.findall(r"第\s*\d+\s*页", joined)) >= 2:
        return True
    # 封面：行数少（≤12），含版本号；或含全称+文档类型；或旧封面的著作权人行
    if n <= 12:
        if version and version in joined:
            return True
        if name and name in joined and re.search(r"(源代码|用户手册|操作手册|说明书)", joined):
            return True
        if "著作权人：" in joined:
            return True
    return False

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("dir")
    ap.add_argument("--yaml", default=None)
    ap.add_argument("--screenshots-dir", default=None, help="截图目录（草稿/截图/），用于截图-章节对应检查（R015）")
    ap.add_argument("--screenshot-name", default=None, help="截图中实际显示的系统名称（用户告知或 OCR），不一致时警告（R013）")
    a = ap.parse_args(); d = Path(a.dir)
    problems, warns, infos = [], [], []
    P = lambda m: problems.append("❌ " + m)
    W = lambda m: warns.append("⚠️  " + m)
    I = lambda m: infos.append("ℹ️  " + m)

    form = next(d.glob("*-申请表.docx"), None)
    manual = next(d.glob("*_操作手册.docx"), None)
    code = next(d.glob("*_软件代码.docx"), None)
    for label, f in (("申请表", form), ("操作手册", manual), ("软件代码", code)):
        if f is None: P(f"缺少 {label} docx")
    data = {}
    if a.yaml: data = yaml.safe_load(Path(a.yaml).read_text(encoding="utf-8")) or {}
    elif form:
        doc = Document(form)
        for t in doc.tables:
            for row in t.rows:
                c = row.cells
                if len(c) >= 2:
                    k = re.split(r"[（(\n]", c[0].text.strip())[0].strip(); v = c[-1].text.strip()
                    if k and k not in data: data[k] = v
    name, ver, owner = data.get("软件全称", ""), data.get("版本号", ""), data.get("权利人名称", "")
    full = f"{name}{ver}"
    I(f"申请表: 全称「{name}」 版本「{ver}」 权利人「{owner}」")

    # 3 申请表指标（R002：只保留详版 500–1300；简版官网填报已不再生成）
    long_ = data.get("软件的主要功能（详版）", "") or data.get("软件的主要功能", "")
    n = cjk_len(long_)
    if not 500 <= n <= 1300: P(f"主要功能（详版）{n} 字，需 500–1300")
    else: I(f"主要功能（详版）{n} 字 ✓")
    t = cjk_len(data.get("软件的技术特点", ""))
    if t > 100: P(f"技术特点 {t} 字，需 ≤100")
    if ver and not re.fullmatch(r"V\d+(\.\d+)*", ver): P(f"版本号「{ver}」应为 V1.0 形式")

    # 1/2 手册
    if manual:
        m = Document(manual); h = header_text(m)
        if h != full: P(f"手册页眉「{h}」≠ 申请表「{full}」")
        paras = [p.text.strip() for p in m.paragraphs]
        # 封面著作权人：新封面无此行（对齐用户手册），仅当存在"著作权人："段落时才校验
        cover_owner = next((p.replace("著作权人：", "") for p in paras if p.startswith("著作权人：")), None)
        if cover_owner is not None and cover_owner != owner:
            P(f"手册封面著作权人「{cover_owner}」≠ 申请表权利人「{owner}」")
        # 封面版本：接受"版本：V1.0"或独立的"V1.0"段落
        cover_ver = next((p.replace("版本：", "") for p in paras if p.startswith("版本：")), None)
        if cover_ver is None:
            cover_ver = next((p for p in paras if re.fullmatch(r"V\d+(\.\d+)*", p)), None)
        if cover_ver is not None and cover_ver != ver:
            P(f"手册封面版本「{cover_ver}」≠「{ver}」")
        if paras and paras[0] == "" and name not in paras[:12]: P("手册封面未见软件全称")
        leftovers = [p for p in paras if p.startswith(("【截图预留", "【截图缺失"))]
        if leftovers: P(f"手册仍有 {len(leftovers)} 处截图占位: " + " / ".join(x[:30] for x in leftovers[:5]))
        h1 = [p.text.strip() for p in m.paragraphs if p.style.name == "Heading 1"]
        imgs = m.element.body.xml.count("<pic:pic")
        body_chars = sum(len(p) for p in paras)
        pdf = manual.with_suffix(".pdf")
        if pdf.exists():
            try:
                from pypdf import PdfReader
                real = len(PdfReader(str(pdf)).pages) - 2  # 去封面、目录
                I(f"手册: 一级标题 {len(h1)} 个，插图 {imgs} 张，正文约 {body_chars} 字，PDF 实际 {real} 页（不含封面目录）")
                if real < 20: W(f"手册 PDF 仅 {real} 页 < 20 页实操建议（用户可选精简文案，接受补正风险）")
            except ImportError:
                I("装 pypdf 可按 PDF 真实页数校验：uv run --with pypdf ...")
        else:
            est_pages = body_chars / 900 + imgs * 0.45
            I(f"手册: 一级标题 {len(h1)} 个，插图 {imgs} 张，正文约 {body_chars} 字，估算 {est_pages:.0f} 页（含图，不含封面目录；转 PDF 后按真实页数校验）")
            if est_pages < 20: W(f"手册估算仅 {est_pages:.0f} 页 < 20 页实操建议（用户可选精简文案，接受补正风险）")
        if imgs == 0: P("手册没有任何截图")
        # 正文顶格检查（用户要求：内容顶格写，不要空格）
        indented = [p.text.strip()[:24] for p in m.paragraphs
                    if p.text.strip() and not p.style.name.startswith("Heading")
                    and "<pic:pic" not in p._element.xml
                    and (p.paragraph_format.first_line_indent or 0) > 0]
        if indented:
            W(f"手册正文存在 {len(indented)} 处首行缩进（应顶格）：" + "、".join(indented[:3]) + (" …" if len(indented) > 3 else ""))
        func = long_
        for title in h1:
            core = re.sub(r"^[\d.、\s一二三四五六七八九十]+", "", title)
            core = re.sub(r"(系统|模块|管理|功能|设置)$", "", core) or core
            if core and core.replace(" ", "") not in func.replace(" ", ""): I(f"手册章节「{title}」在申请表主要功能里未出现，确认是否需要对应")
        for bad in ("import ", "public class", "function(", "SELECT ", "def "):
            if any(bad in p for p in paras): P(f"手册疑似含代码片段「{bad.strip()}」，说明书禁止出现功能函数代码"); break

        # 7 截图对应性（R015）
        if a.screenshots_dir:
            sdir = Path(a.screenshots_dir)
            shots = [p for p in sdir.rglob("*") if p.is_file() and p.suffix.lower() in IMG_EXT] if sdir.exists() else []
            if shots:
                if len(shots) > imgs:
                    P(f"截图目录有 {len(shots)} 张截图，但手册仅 {imgs} 张插图——有截图未找到对应章节："
                      + "、".join(p.name for p in shots[imgs:imgs + 5]) + (" …" if len(shots) - imgs > 5 else ""))
                else:
                    I(f"截图 {len(shots)} 张 ⊆ 手册插图 {imgs} 张 ✓")
                # 截图页面名必须在手册正文出现（有什么截图写什么，截图不可缺少）
                manual_text = "".join(paras)
                for shot in shots:
                    pname = re.sub(r"^\d+[-_]", "", shot.stem).strip()  # 文件名推断页面名（同 scan_screenshots）
                    if pname and pname not in manual_text:
                        P(f"截图「{shot.name}」对应页面「{pname}」未在手册正文中出现——截图不可缺少，请按截图补写对应小节")
                # 裸截图检查：图片段落前后必须有文字（允许中间夹空段落，对齐用户手册排版）
                ps = m.paragraphs
                def neighbor_has_text(idx, step):
                    j = idx + step
                    while 0 <= j < len(ps):
                        if "<pic:pic" in ps[j]._element.xml: return False
                        if re.sub(r"\s", "", ps[j].text): return True
                        j += step
                    return False
                for idx, p in enumerate(ps):
                    if "<pic:pic" in p._element.xml:
                        prev_ok = neighbor_has_text(idx, -1)
                        nxt_ok = neighbor_has_text(idx, 1)
                        if not (prev_ok and nxt_ok):
                            W(f"第 {idx + 1} 段插图前后缺少文字描述（截图与文字必须对应，请补功能描述/操作步骤）")
                # 每个功能二级小节必须有图（截图驱动，无图小节报错；概述/登录/常见问题等非功能小节豁免）
                h2_count, h2_no_img = {}, []
                cur = None
                for p in m.paragraphs:
                    if p.style.name == "Heading 2":
                        cur = p.text.strip(); h2_count.setdefault(cur, 0)
                    elif "<pic:pic" in p._element.xml and cur:
                        h2_count[cur] = h2_count.get(cur, 0) + 1
                for h2, cnt in h2_count.items():
                    if cnt == 0 and not re.search(r"(简介|功能特点|运行环境|常见问题|登录|概述)", h2):
                        h2_no_img.append(h2)
                if h2_no_img:
                    P(f"{len(h2_no_img)} 个功能二级小节没有截图（截图不可缺少）：" + "、".join(h2_no_img[:5]) + (" …" if len(h2_no_img) > 5 else ""))
            else:
                W("未在 --screenshots-dir 找到截图文件，跳过截图对应检查")
        # 8 截图系统名称不一致柔性处理（R013）
        if a.screenshot_name and a.screenshot_name.strip() and a.screenshot_name.strip() != name:
            W(f"截图中系统名称为「{a.screenshot_name}」，与申报全称「{name}」不一致。版权中心可能因此要求补正。"
              + "已柔性放行，请在最终交付说明中再次提醒用户；或选择暂停重新截图。")

    # 6 代码
    if code:
        c = Document(code); h = header_text(c)
        if h != full: P(f"代码页眉「{h}」≠ 申请表「{full}」")
        pages = split_pages(c)
        last_note = any("本页为源代码最后一页" in x for x in pages[-1]) if pages else False
        code_pages = [pg for pg in pages if not looks_like_front(pg, name, ver)]
        # 末页的"最后一页"标注行不计入行数
        code_pages = [pg if i < len(code_pages) - 1 else [x for x in pg if "本页为源代码最后一页" not in x] for i, pg in enumerate(code_pages)]
        bad = [i + 1 for i, pg in enumerate(code_pages[:-1]) if len(pg) != 50]
        I(f"代码: {len(pages)} 页（正文 {len(code_pages)} 页，不含封面/目录），末页 {len(code_pages[-1]) if code_pages else 0} 行" + ("，已标注末页" if last_note else ""))
        if bad: P(f"代码非末页行数≠50: 第 {bad[:10]} 页")
        if len(code_pages) < 60: P(f"代码仅 {len(code_pages)} 页 < 60 页下限（代码有多少写多少，但不得少于 60 页）")
        if pages and len(code_pages[-1]) < 50 and not last_note: P("末页不足 50 行但未标注「本页为源代码最后一页」")
        idx = d / "代码文档索引.md"
        if idx.exists():
            exts = set(re.findall(r"`[^`]+?(\.[a-zA-Z]+)`", idx.read_text(encoding="utf-8")))
            langs = set()
            for e in exts:
                l = EXT_LANG.get(e.lower())
                if l == "Vue": langs |= VUE_IMPLIES
                elif l: langs.add(l)
            declared = data.get("编程语言", "")
            missing = [l for l in langs if l not in declared]
            if missing: P(f"代码涉及 {sorted(langs)}，申请表编程语言「{declared}」缺少 {missing}")
            else: I(f"编程语言 {sorted(langs)} ⊆ 申请表 ✓")

    print("\n".join(infos)); print()
    if warns:
        print("\n".join(warns)); print(f"\n共 {len(warns)} 个警告（人工裁决，不阻断提交，但请逐条确认）。")
    if problems:
        print("\n".join(problems)); print(f"\n共 {len(problems)} 个错误，修正后重跑。"); sys.exit(1)
    print("✅ 自检通过。仍需人工确认：截图完整含标题栏、权属/完成日期真实、Word 目录已更新。")

if __name__ == "__main__":
    main()
