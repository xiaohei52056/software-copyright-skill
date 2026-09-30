#!/usr/bin/env python3
"""
检测项目目录下用户已有的软著材料（R012），供工作流自动参考。

用法:
  python3 detect_existing.py <项目根目录> [-o 草稿/已有材料.md] [--json 草稿/existing.json]

检测对象（含子目录）:
  - *申请表*.txt / *申请表*.xlsx / *申请表*.docx
  - *用户手册*.docx / *操作手册*.docx
  - *源代码*.docx / *软件代码*.docx
  - 软著登记表-内容.xlsx 等

输出: 命中文件清单（路径 + 类型）。SKILL.md 第 0 步据此提示用户"是否参考已有材料的内容和风格"。
"""
import argparse, json, sys
from pathlib import Path

PATTERNS = [
    ("申请表模板", ["*申请表*模板*.docx", "*登记表*模板*.docx", "*模板*.docx"]),
    ("申请表", ["*申请表*.txt", "*申请表*.xlsx", "*申请表*.docx", "*登记表*.xlsx", "*登记表*.docx"]),
    ("操作手册", ["*用户手册*.docx", "*操作手册*.docx", "*说明书*.docx", "*手册*.docx"]),
    ("源代码文档", ["*源代码*.docx", "*软件代码*.docx", "*代码文档*.docx"]),
]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--json", default=None)
    a = ap.parse_args()
    root = Path(a.root).resolve()
    found = []
    for kind, pats in PATTERNS:
        for pat in pats:
            for p in root.rglob(pat):
                if any(seg in {"node_modules", ".git", "软件著作权申请资料"} for seg in p.parts): continue
                found.append({"kind": kind, "path": p.relative_to(root).as_posix()})
    seen = set(); uniq = []
    for f in found:
        if f["path"] in seen: continue
        seen.add(f["path"]); uniq.append(f)

    md = ["# 检测到的已有软著材料", ""]
    if uniq:
        md.append(f"检测到 {len(uniq)} 份已有材料：")
        md.append("| 类型 | 路径 |")
        md.append("|---|---|")
        for f in uniq:
            md.append(f"| {f['kind']} | `{f['path']}` |")
        md += ["", "> 是否参考这些材料的内容和风格来生成新材料？（Y/N）",
               "> 同意后：申请表字段优先从已有材料提取（全称/版本/环境/源程序量等）；",
               "> 手册章节顺序与措辞参考已有手册；参考来源必须在交付说明中标注。"]
    else:
        md.append("未检测到已有材料，从零开始生成。")

    out = a.out or "已有材料.md"
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    Path(out).write_text("\n".join(md) + "\n", encoding="utf-8")
    print(f"✅ {out}")
    if a.json:
        Path(a.json).write_text(json.dumps({"root": str(root), "existing": uniq}, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"✅ {a.json}")
    print(f"检测到 {len(uniq)} 份已有材料" if uniq else "未检测到已有材料")

if __name__ == "__main__":
    main()
