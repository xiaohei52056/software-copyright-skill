#!/usr/bin/env python3
"""
扫描截图目录，生成截图清单（R015：手册以截图为核心驱动）。

用法:
  python3 scan_screenshots.py <截图目录> [-o 草稿/截图清单.md] [--json 草稿/screenshots.json]

输出每张截图: 编号、文件名、推断页面名（从文件名提取，如 07-拧紧数据查询.png → 拧紧数据查询）、
所属一级菜单（由文件名前缀/目录名推断，人工确认后回填）。

手册写作规则（配合 SKILL.md 第 4 步）:
  - 截图里有什么页面，手册就必须有对应二级小节与文字描述
  - 每张截图插入位置前后必须有文字（功能描述 + 操作步骤），不允许"裸截图"
  - 截图页面名无法从文件名识别时，由 Agent 用 OCR 或请用户标注
"""
import argparse, json, re, sys
from pathlib import Path

IMG_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}

def guess_page(filename: str) -> str:
    stem = Path(filename).stem
    # 去掉编号前缀：07-拧紧数据查询 / 07_拧紧数据查询 / 07 拧紧数据查询
    stem = re.sub(r"^\s*\d+\s*[-_—]\s*", "", stem)
    stem = re.sub(r"^\s*\d+\s*", "", stem)
    stem = stem.strip(" _-—")
    return stem or filename

def guess_menu(filename: str, stem: str) -> str:
    # 从编号段或目录名推断一级菜单，无法推断返回空
    return ""

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dir", help="截图目录（通常为 软件著作权申请资料/草稿/截图/）")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--json", default=None)
    a = ap.parse_args()
    d = Path(a.dir).resolve()
    files = sorted([p for p in d.rglob("*") if p.is_file() and p.suffix.lower() in IMG_EXT])
    items = []
    for i, p in enumerate(files, 1):
        page = guess_page(p.name)
        items.append({"no": i, "file": p.name, "page": page, "menu": guess_menu(p.name, page)})

    md = ["# 截图清单（手册章节结构的主要依据）", "",
          "> 由 scan_screenshots.py 生成（R015）。**截图驱动原则**：截图里有什么页面，手册就必须有对应二级小节和文字描述。",
          "", f"共 {len(items)} 张截图。", "",
          "| 编号 | 文件名 | 推断页面名 | 所属一级菜单（人工确认/回填） |", "|---|---|---|---|"]
    for it in items:
        md.append(f"| {it['no']} | {it['file']} | {it['page']} | {it['menu']} |")
    md += ["", "## 使用说明", "",
           "1. 手册二级小节数量必须 ≥ 截图数量；每张截图都要能在手册中找到对应章节。",
           "2. 每张截图插入位置前后必须有文字（功能描述 + 操作步骤），不允许连续多张截图无文字。",
           "3. 页面名无法从文件名识别的，用 OCR 识别截图中的页面标题，或请用户标注后回填上表。",
           "4. 一级菜单列请按系统菜单结构（extract_menu.py 输出）人工确认。"]
    if not items:
        md = ["# 截图清单", "", "⚠️ 未在目录中检测到截图（png/jpg 等）。请将系统截图放入本目录后重跑；",
              "手册相应位置保留【截图预留】占位并列出需要的截图清单给用户。"]
    out = a.out or "截图清单.md"
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    Path(out).write_text("\n".join(md) + "\n", encoding="utf-8")
    print(f"✅ {out}")
    if a.json:
        Path(a.json).write_text(json.dumps({"dir": str(d), "screenshots": items}, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"✅ {a.json}")
    print(f"检测到 {len(items)} 张截图")

if __name__ == "__main__":
    main()
