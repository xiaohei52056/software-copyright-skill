#!/usr/bin/env python3
"""
截图批量重命名（软著手册截图驱动辅助工具）。

用法:
  python3 rename_screenshots.py 草稿/截图/ --names "01-首页,02-用户管理,03-角色管理"
  python3 rename_screenshots.py 草稿/截图/ --map 新名单.txt        # 每行一个页面名
  python3 rename_screenshots.py 草稿/截图/ --names "..." --dry-run # 只预览不执行

规则:
  - 只处理图片文件（.png/.jpg/.jpeg/.bmp/.gif）。
  - 名单按目录内现有图片的字母/数字顺序一一对应；名单数量必须等于图片数。
  - 新名为页面名（如"首页"），脚本自动加两位序号与扩展名 → "01-首页.png"；
    若新名已带序号（如"01-首页"）则原样使用。
  - 采用"先拷到临时目录、清空原名、再移回"两步法，避免 Windows 下批量改名冲突。
  - --dry-run 仅打印计划，不落盘。

说明: 截图命名规范（NN-页面名.png）是 scan_screenshots.py 推断页面名与
      操作手册章节结构的前提，重命名后请重新运行 scan_screenshots.py 生成截图清单。
"""
import argparse, os, re, shutil, sys
from pathlib import Path

IMG_EXT = {".png", ".jpg", ".jpeg", ".bmp", ".gif"}

def numbered(name: str) -> str:
    """新名已带 NN- 前缀则原样用；否则自动加两位序号。"""
    if re.match(r"^\d{2,}[-_]", name):
        return name
    return None

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("dir", help="截图目录")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--names", help="逗号分隔的页面名列表，按目录内顺序一一对应")
    g.add_argument("--map", help="名单文件，每行一个页面名（空行/注释忽略）")
    ap.add_argument("--dry-run", action="store_true", help="只预览重命名计划，不落盘")
    a = ap.parse_args()

    src = Path(a.dir).resolve()
    if not src.is_dir():
        sys.exit(f"目录不存在: {src}")
    files = sorted(p for p in src.iterdir() if p.is_file() and p.suffix.lower() in IMG_EXT)
    if not files:
        sys.exit(f"目录中没有图片文件: {src}")

    if a.names:
        names = [s.strip() for s in a.names.split(",") if s.strip()]
    else:
        names = []
        for raw in Path(a.map).read_text(encoding="utf-8-sig").splitlines():
            s = raw.strip()
            if s and not s.startswith("#"):
                names.append(s)
    if len(names) != len(files):
        sys.exit(f"名单数量 {len(names)} ≠ 图片数量 {len(files)}。请核对后重试。")

    plan = []  # (旧文件, 新文件)
    for i, (old, name) in enumerate(zip(files, names), 1):
        base = numbered(name) or f"{i:02d}-{name}"
        new = old.with_name(base + old.suffix.lower())
        if new == old:
            continue
        plan.append((old, new))

    if not plan:
        print("✅ 所有截图已符合命名规范，无需重命名。")
        return

    print(f"重命名计划（{len(plan)} 个文件）：")
    for old, new in plan:
        print(f"  {old.name}  →  {new.name}")
    if a.dry_run:
        print("（--dry-run：仅预览，未执行）")
        return

    # 两步法：先拷到临时目录（保留扩展名），清空原名，再移回
    tmp = src / ".rename_tmp"
    tmp.mkdir(exist_ok=True)
    try:
        for old, new in plan:
            (tmp / new.name).write_bytes(old.read_bytes())
        for old, _ in plan:
            old.unlink()
        for _, new in plan:
            shutil.move(str(tmp / new.name), str(new))
        print(f"✅ 已重命名 {len(plan)} 个文件。请重新运行 scan_screenshots.py 生成截图清单。")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

if __name__ == "__main__":
    main()
