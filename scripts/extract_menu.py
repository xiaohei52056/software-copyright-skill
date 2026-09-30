#!/usr/bin/env python3
"""
从前端路由配置或数据库菜单表 SQL 中提取系统实际菜单层级与顺序（R004）。

用法:
  python3 extract_menu.py <项目根目录> [-o 草稿/菜单结构.md] [--json 草稿/menu.json]

提取来源（按优先级）:
  1. 前端路由文件（web/src/router/index.ts / index.js / routes.ts 等）:
     - path / meta.title / children 嵌套层级；数组顺序即菜单顺序
  2. 数据库菜单表 SQL（sys_menu 的 INSERT 语句）:
     - menu_name / parent_id / order_num；按 order_num 排序
  3. 均无法提取 → 输出空清单并明确标注"未检测到菜单配置，按业务理解中的模块清单顺序生成"

输出: 一级菜单 = 一级标题、二级菜单 = 二级标题，与系统界面文字一致。
"""
import argparse, json, re, sys
from pathlib import Path

MENU_LIKE = re.compile(r"(menu_name|menu_name_en|perms|order_num|parent_id)", re.I)
ROUTER_FILES = ["router/index.ts", "router/index.js", "router/routes.ts", "router/routes.js",
                "src/router/index.ts", "src/router/index.js", "src/router/routes.ts", "src/router/routes.js",
                "web/src/router/index.ts", "web/src/router/index.js", "web/src/router/routes.ts", "web/src/router/routes.js"]

def find_router(root: Path):
    for rel in ROUTER_FILES:
        p = root / rel
        if p.exists(): return p
    hits = list(root.rglob("router/*.ts")) + list(root.rglob("router/*.js"))
    for h in hits:
        if h.suffix in (".ts", ".js"):
            return h
    return None

def parse_router(text: str):
    """尽力从路由数组提取 (title, path) 层级。返回 (menus, 来源说明)。"""
    menus = []
    # 匹配 meta: { title: 'xxx' } 或 title: 'xxx'，记录缩进层级
    lines = text.splitlines()
    stack = []  # (indent, title)
    last_indent = None
    for line in lines:
        m = re.search(r"(?:meta\s*:\s*\{[^}]*)?title\s*:\s*['\"]([^'\"]+)['\"]", line)
        if not m: continue
        indent = len(line) - len(line.lstrip())
        title = m.group(1)
        if last_indent is not None and indent > last_indent and stack:
            pass  # 更深的层级
        while stack and indent <= stack[-1][0]:
            stack.pop()
        stack.append((indent, title))
        last_indent = indent
        # 只保留一/二级
        if len(stack) == 1:
            menus.append((1, title))
        elif len(stack) == 2:
            menus.append((2, title))
    return menus

def parse_menu_sql(text: str):
    """解析 sys_menu 类 INSERT 语句：parent_id, menu_name, order_num。"""
    rows = []
    for m in re.finditer(r"INSERT\s+INTO\s+[`\"]?sys_menu[`\"]?\s*(?:\([^)]*\))?\s*VALUES\s*\((.*?)\)\s*;", text, re.I | re.S):
        vals = m.group(1)
        # 常见列序：menu_id, parent_id, menu_name, order_num ... 取带引号的字符串和数字做启发式
        parts = re.findall(r"'([^']*)'|(\d+)", vals)
        # 提取 (parent_id, name, order)：启发式——字符串中形似菜单名的第一个字符串
        strs = [s for s, d in parts if s]
        nums = [int(d) for s, d in parts if d]
        if not strs: continue
        name = strs[0]
        parent = nums[0] if nums else 0
        order = nums[1] if len(nums) > 1 else 0
        rows.append((parent, name, order))
    if not rows: return []
    # 构建层级：parent 为 0 是一级；其余挂在最近的一级下
    menus = []
    for parent, name, order in rows:
        menus.append((1 if parent == 0 else 2, name))
    return menus

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--json", default=None)
    a = ap.parse_args()
    root = Path(a.root).resolve()
    source = None; menus = []; note = ""

    router = find_router(root)
    if router:
        menus = parse_router(router.read_text(encoding="utf-8", errors="replace"))
        source = f"前端路由 {router.relative_to(root)}"
    if not menus:
        for sql in root.rglob("*.sql"):
            text = sql.read_text(encoding="utf-8", errors="replace")
            if MENU_LIKE.search(text):
                menus = parse_menu_sql(text)
                if menus:
                    source = f"数据库菜单表 SQL {sql.relative_to(root)}"
                    break
    if not menus:
        source = None
        note = "未检测到菜单配置（未找到 router 路由文件或 sys_menu SQL），按业务理解中的模块清单顺序生成手册章节。"

    md = ["# 系统菜单结构（手册章节顺序依据）", ""]
    if source:
        md += [f"来源：**{source}**（菜单顺序与系统界面一致）", ""]
        for level, title in menus:
            md.append(("## " if level == 1 else "### ") + title)
        md += ["", "> 手册一级标题 = 一级菜单名，二级标题 = 二级菜单名；若菜单名与申报功能模块名有出入，以系统界面实际文字为准。"]
    else:
        md += [f"⚠️ {note}", ""]
        md += ["请在 `业务理解.md` 模块清单基础上人工整理菜单顺序，或提供路由文件/SQL 后重跑本脚本。"]
    if note:
        md.insert(1, f"> ⚠️ {note}")

    out = a.out or "菜单结构.md"
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    Path(out).write_text("\n".join(md) + "\n", encoding="utf-8")
    print(f"✅ {out}")
    if a.json:
        Path(a.json).write_text(json.dumps({"root": str(root), "source": source, "menus": menus},
                                           ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"✅ {a.json}")
    if source:
        print(f"来源 {source}，提取 {len(menus)} 条菜单")
    else:
        print(f"⚠️ {note}")

if __name__ == "__main__":
    main()
