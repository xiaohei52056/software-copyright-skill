#!/usr/bin/env python3
"""
从前端 .vue 页面提取真实表格列名、按钮名、表单字段名（R005），供手册生成引用。

用法:
  python3 extract_fields.py <项目根目录> [-o 草稿/页面字段清单.md] [--json 草稿/fields.json]

提取内容:
  - <el-table-column label="...">         → 表格列名
  - <el-button>新增</el-button> / 文字     → 按钮名
  - <el-form-item label="...">           → 表单字段名
  - <el-select> / <el-option> 文字        → 下拉筛选选项
  - 页面 <title> 或文件名                  → 页面名

手册写作时：每个有对应 .vue 的章节生成"页面展示 XX、XX、XX 等字段信息"，操作步骤使用真实按钮名。
无法解析的页面回退到通用描述，不允许出现空字段。
"""
import argparse, json, re, sys
from pathlib import Path

VUE_EXT = (".vue",)
TAG_RE = {
    "columns": re.compile(r"<el-table-column[^>]*label=\"([^\"]+)\"", re.I),
    "buttons": re.compile(r"<el-button[^>]*>\s*([^<]{1,12}?)\s*</el-button>", re.I),
    "form_items": re.compile(r"<el-form-item[^>]*label=\"([^\"]+)\"", re.I),
    "options": re.compile(r"<el-option[^>]*label=\"([^\"]+)\"", re.I),
}
BUTTON_ATTR = re.compile(r"<el-button[^>]*>\s*([^<]{1,12}?)\s*</el-button>", re.I)
TITLE_RE = re.compile(r"<title>([^<]+)</title>", re.I)

def page_name(path: Path, rel: str) -> str:
    m = TITLE_RE.search(path.read_text(encoding="utf-8", errors="replace"))
    if m: return m.group(1).strip()
    stem = path.stem
    stem = re.sub(r"(index|Index|list|List|edit|Edit|detail|Detail|add|Add)$", "", stem)
    return stem or rel

def extract_vue(path: Path) -> dict:
    text = path.read_text(encoding="utf-8", errors="replace")
    out = {"columns": [], "buttons": [], "form_items": [], "options": []}
    for m in TAG_RE["columns"].finditer(text):
        v = m.group(1).strip()
        if v and v not in out["columns"]: out["columns"].append(v)
    for m in TAG_RE["buttons"].finditer(text):
        v = m.group(1).strip()
        if v and v not in out["buttons"]: out["buttons"].append(v)
    # 按钮也可能以 :text="..." 或 @click 内联
    for m in re.finditer(r"<el-button[^>]*>([^<]{1,12})</el-button>", text, re.I):
        v = m.group(1).strip()
        if v and v not in out["buttons"]: out["buttons"].append(v)
    for m in TAG_RE["form_items"].finditer(text):
        v = m.group(1).strip()
        if v and v not in out["form_items"]: out["form_items"].append(v)
    for m in TAG_RE["options"].finditer(text):
        v = m.group(1).strip()
        if v and v not in out["options"]: out["options"].append(v)
    return out

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--json", default=None)
    ap.add_argument("--min-items", type=int, default=0, help="提取项少于该数的页面也列出（默认全部列出）")
    a = ap.parse_args()
    root = Path(a.root).resolve()
    pages = []
    for p in root.rglob("*.vue"):
        if any(seg in {"node_modules", "dist", "build", ".git"} for seg in p.parts): continue
        rel = p.relative_to(root).as_posix()
        info = extract_vue(p)
        if not any(info.values()): continue
        info["page"] = page_name(p, rel)
        info["file"] = rel
        pages.append(info)

    md = ["# 前端页面字段清单（手册内容素材）", "",
          "> 由 extract_fields.py 从 .vue 文件提取（R005）。手册写作时优先使用真实字段名/按钮名，提升真实感。",
          "", f"共解析 {len(pages)} 个页面。", ""]
    for pg in pages:
        md.append(f"## {pg['page']}  (`{pg['file']}`)")
        if pg["columns"]:
            md.append(f"- 表格列/展示字段：{'、'.join(pg['columns'])}")
        if pg["form_items"]:
            md.append(f"- 表单字段：{'、'.join(pg['form_items'])}")
        if pg["buttons"]:
            md.append(f"- 按钮：{'、'.join(pg['buttons'])}")
        if pg["options"]:
            md.append(f"- 下拉选项：{'、'.join(pg['options'])}")
        md.append("")
    if not pages:
        md += ["⚠️ 未找到 .vue 页面文件，或页面无 el-table/el-button/el-form 元素。", "",
               "手册相关章节回退到通用描述（功能描述 + 常规操作步骤），并在交付说明中标注哪些页面未提取到真实字段。"]
    md += ["## 写作模板", "",
           "- 列表页：\"页面所展示的信息有序号、{列名}、{列名}、操作等字段信息。\"",
           "- 操作步骤：\"点击{按钮名}按钮，即可完成……\"",
           "- 提取失败页面：\"该页面提供{功能}的查看与维护能力。\"（通用描述，禁止空字段）"]

    out = a.out or "页面字段清单.md"
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    Path(out).write_text("\n".join(md) + "\n", encoding="utf-8")
    print(f"✅ {out}")
    if a.json:
        Path(a.json).write_text(json.dumps({"root": str(root), "pages": pages}, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"✅ {a.json}")
    print(f"解析 {len(pages)} 个 .vue 页面")

if __name__ == "__main__":
    main()
