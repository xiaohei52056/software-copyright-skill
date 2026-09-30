#!/usr/bin/env python3
"""
扫描项目源码，输出候选文件清单（供人工勾选）和源程序量多口径统计。

用法:
  python3 collect_sources.py <项目根目录> [-o 草稿/代码文件选择.md] [--json 草稿/sources.json] [--count-mode raw|effective|self-only]

输出:
  - 按目录分组的候选文件表：路径 / 语言 / 总行数 / 有效行数（去空行、去纯注释）/ 归属（自研|框架）/ 建议角色
  - 三种源程序量口径（R001）：
      * 原始总行数（raw）：全部源代码文件（排除依赖目录后）的物理行数，含空行、注释，含前后端
      * 有效行数（effective）：去空行、去纯注释后的行数
      * 自研有效行数（self-only）：有效行中剔除常见开源框架模块（RuoYi 等）后的行数
  - 自动排除 node_modules、构建产物、锁文件、图片、第三方 vendor 等
默认只统计，不抽取；抽取由 build_code_docx.py 按人工确认后的顺序清单执行。

--count-mode 决定"建议填入申请表源程序量"的口径提示，默认 raw（版权中心通常接受物理行数）。
"""
import argparse, json, os, re, sys
from pathlib import Path
from collections import defaultdict

EXT_LANG = {
    ".java": "Java", ".kt": "Kotlin", ".py": "Python", ".js": "JavaScript", ".jsx": "JavaScript",
    ".ts": "TypeScript", ".tsx": "TypeScript", ".vue": "Vue(JavaScript/HTML/CSS)", ".html": "HTML",
    ".htm": "HTML", ".css": "CSS", ".scss": "CSS", ".less": "CSS", ".go": "Go", ".rs": "Rust",
    ".c": "C", ".h": "C", ".cpp": "C++", ".hpp": "C++", ".cs": "C#", ".php": "PHP", ".rb": "Ruby",
    ".swift": "Swift", ".m": "Objective-C", ".sql": "SQL", ".sh": "Shell", ".dart": "Dart",
    ".xml": "XML", ".yaml": "YAML", ".yml": "YAML", ".json": "JSON", ".properties": "Properties",
}
CONFIG_LIKE = {".xml", ".yaml", ".yml", ".json", ".properties", ".sql"}
EXCLUDE_DIRS = {
    "node_modules", ".git", ".svn", ".idea", ".vscode", "dist", "build", "target", "out", "bin", "obj",
    "__pycache__", "__MACOSX", ".venv", "venv", "env", "vendor", "third_party", "thirdparty", "lib", "libs",
    "coverage", ".next", ".nuxt", ".cache", "public/static", "static/js/lib", "assets/lib", "logs", "tmp",
}
EXCLUDE_FILES = re.compile(
    r"(\.min\.(js|css)$|\.lock$|package-lock\.json$|yarn\.lock$|pnpm-lock\.yaml$|\.map$|\.d\.ts$|"
    r"^\.|\.(png|jpg|jpeg|gif|svg|ico|woff2?|ttf|eot|mp4|pdf|docx?|xlsx?)$)", re.I)

ENTRY_HINTS = re.compile(r"(main\.(java|py|js|ts|go)|App\.(vue|jsx|tsx)|index\.(js|ts|html|vue)|Application\.java|app\.py|server\.(js|ts)|router|routes|login)", re.I)
CORE_HINTS = re.compile(r"(service|controller|domain|model|store|api|core|business|logic|engine|handler|views?|pages?|components?)", re.I)
UTIL_HINTS = re.compile(r"(util|helper|common|tool)", re.I)

# R008：常见开源后台框架特征路径（RuoYi / 芋道 yudao / pig 等）。命中即判定为"框架代码"，
# 只用于"自研有效行数"口径与候选表的归属标注，不阻断文件进入"全部代码"选项。
FRAMEWORK_PATTERNS = [
    re.compile(r"(^|/)(ruoyi-[a-z]+|yudao-(framework|module-system|module-infra)|pig-common|pig-register)(/|$)", re.I),
    re.compile(r"(^|/)(system|codegen|generator|framework|monitor|job|scheduler)(/|$)", re.I),
    re.compile(r"(^|/)common/(constant|core|utils?|enums?|exception|config|annotation|result|page|aspect|datasource)(/|$)", re.I),
    re.compile(r"(^|/)(config|security|filter|interceptor|datasource|mybatis|redis)/", re.I),
]

BLOCK_COMMENT = {"/*": "*/", "<!--": "-->", '"""': '"""', "'''": "'''"}

def effective_lines(text: str) -> int:
    """去空行、去纯注释行后的行数（与 build_code_docx.py 默认口径一致）"""
    n = 0; in_block = None
    for raw in text.splitlines():
        s = raw.strip()
        if in_block:
            if in_block in s: in_block = None
            continue
        if not s: continue
        if s.startswith(("//", "#", "*", "--")) and not s.startswith("#!"): continue
        opened = False
        for o, c in BLOCK_COMMENT.items():
            if s.startswith(o):
                if c not in s[len(o):]: in_block = c
                opened = True; break
        if opened: continue
        n += 1
    return n

def role_of(rel: str, ext: str) -> str:
    if ext in CONFIG_LIKE: return "配置(低优先)"
    if ENTRY_HINTS.search(rel): return "入口/路由"
    if CORE_HINTS.search(rel): return "核心业务"
    if UTIL_HINTS.search(rel): return "工具"
    return "其他"

def is_framework(rel: str) -> bool:
    return any(p.search(rel) for p in FRAMEWORK_PATTERNS)

def scan(root: Path):
    rows = []
    for dp, dns, fns in os.walk(root):
        dns[:] = [d for d in dns if d not in EXCLUDE_DIRS and not d.startswith(".")]
        for fn in fns:
            if EXCLUDE_FILES.search(fn): continue
            p = Path(dp) / fn
            ext = p.suffix.lower()
            if ext not in EXT_LANG: continue
            try:
                text = p.read_text(encoding="utf-8", errors="replace")
            except Exception:
                continue
            total = text.count("\n") + (0 if text.endswith("\n") or not text else 1)
            eff = effective_lines(text)
            if eff == 0: continue
            rel = p.relative_to(root).as_posix()
            rows.append({"path": rel, "lang": EXT_LANG[ext], "ext": ext, "total": total, "effective": eff,
                         "role": role_of(rel, ext), "framework": is_framework(rel)})
    return rows

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("-o", "--out", default=None, help="Markdown 清单输出路径")
    ap.add_argument("--json", default=None, help="JSON 清单输出路径")
    ap.add_argument("--min-lines", type=int, default=10, help="有效行数低于此值的文件不列出")
    ap.add_argument("--count-mode", choices=["raw", "effective", "self-only"], default="raw",
                    help="建议填入申请表的源程序量口径（默认 raw：原始总行数，版权中心通常接受物理行数）")
    a = ap.parse_args()
    root = Path(a.root).resolve()
    rows = scan(root)

    # ---- 多口径统计（R001）----
    code_rows = [r for r in rows if r["ext"] not in CONFIG_LIKE]
    self_rows = [r for r in code_rows if not r["framework"]]
    stats = {
        "raw":       {"lines": sum(r["total"] for r in code_rows),     "files": len(code_rows)},
        "effective": {"lines": sum(r["effective"] for r in code_rows), "files": len(code_rows)},
        "self-only": {"lines": sum(r["effective"] for r in self_rows), "files": len(self_rows)},
    }
    by_lang = defaultdict(lambda: [0, 0])
    for r in code_rows:
        by_lang[r["lang"]][0] += r["effective"]; by_lang[r["lang"]][1] += 1

    mode_note = {
        "raw":       "原始总行数（物理行，含空行与注释）",
        "effective": "有效行数（去空行、去纯注释）",
        "self-only": "自研有效行数（剔除框架模块）",
    }
    md = [f"# 代码文件选择 —— {root.name}", "",
          f"扫描根目录：`{root}`", "",
          "## 源程序量统计（不含配置类文件，已排除依赖目录）", "",
          f"- **原始总行数**：{stats['raw']['lines']} 行 —— 全部代码文件物理行数（含空行、注释），含前后端",
          f"- **有效行数**：{stats['effective']['lines']} 行 —— 去空行、去纯注释后的行数",
          f"- **自研有效行数**：{stats['self-only']['lines']} 行 —— 有效行中剔除常见开源框架模块（{len(code_rows)-len(self_rows)} 个文件命中框架特征）",
          "",
          f"> 建议填入申请表\"源程序量\"：**{mode_note[a.count_mode]}（{stats[a.count_mode]['lines']} 行）**。版权中心通常接受物理行数；若项目框架代码占比高可改选自研口径。",
          "",
          "| 语言 | 文件数 | 有效行数 |", "|---|---:|---:|"]
    for lang, (n, c) in sorted(by_lang.items(), key=lambda x: -x[1][0]):
        md.append(f"| {lang} | {c} | {n} |")
    md += ["", "> 申请表\"编程语言\"应填入上表中**最终抽取文件**实际涉及的语言。", "",
           "## 候选文件（在\"选\"列打 ✓，并在\"序\"列写抽取顺序；入口文件必须排第 1）", "",
           "| 选 | 序 | 路径 | 语言 | 归属 | 角色 | 总行 | 有效行 |", "|---|---|---|---|---|---|---|---:|"]
    order = {"入口/路由": 0, "核心业务": 1, "工具": 2, "其他": 3, "配置(低优先)": 4}
    listed = [r for r in rows if r["effective"] >= a.min_lines]
    listed.sort(key=lambda r: (order[r["role"]], r["path"]))
    for r in listed:
        own = "框架" if r["framework"] else "自研"
        md.append(f"|  |  | `{r['path']}` | {r['lang']} | {own} | {r['role']} | {r['total']} | {r['effective']} |")
    md += ["", f"共 {len(listed)} 个候选文件（有效行数 ≥ {a.min_lines}）。全量收录：代码有多少写多少，页数不设上限（下限 60 页）。", "",
           "## 代码范围选项（R008）", "",
           "1. **仅自研业务代码**（推荐）：只选\"归属=自研\"的文件，全量均为自研代码。",
           "2. **全部源代码**：自研+框架都选，但起始页优先放自研入口和核心模块，末页以完整模块收尾。",
           "   ⚠️ 开源框架代码占比过高可能被版权中心质疑原创性，建议起始页优先自研代码。",
           "",
           "## 抽取原则", "",
           "1. 首页必须是程序起始（入口/main/登录/index），尾页必须是完整模块收尾。",
           "2. 从入口开始按业务模块**连续**排列全部代码；末页以某个完整模块结束。",
           "3. 确认后把选中文件按顺序写成一行一个路径的 `代码文件顺序.txt`，交给 build_code_docx.py。",
           "4. 代码量不设上限：满足 50 行/页、全量收录、首页入口、末页完整收尾即可，不强制精简。"]

    out = a.out or "代码文件选择.md"
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    Path(out).write_text("\n".join(md) + "\n", encoding="utf-8")
    print(f"✅ {out}")
    if a.json:
        Path(a.json).write_text(json.dumps({"root": str(root), "stats": stats, "count_mode": a.count_mode,
                                             "by_lang": dict(by_lang), "files": rows}, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"✅ {a.json}")
    print(f"源程序量：原始 {stats['raw']['lines']} 行 / 有效 {stats['effective']['lines']} 行 / 自研有效 {stats['self-only']['lines']} 行；候选 {len(listed)} 个文件")

if __name__ == "__main__":
    main()
