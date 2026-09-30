# 🎯 software-copyright-skill

<h3 align="center">软著登记三件套 · 一键生成</h3>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/badge/License-MIT-yellow.svg" alt="License: MIT"></a>
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.10%2B-blue.svg" alt="Python 3.10+"></a>
  <img src="https://img.shields.io/badge/文档-中文-green.svg" alt="中文文档">
  <a href="https://github.com/xiaohei52056/software-copyright-skill/issues"><img src="https://img.shields.io/badge/Issues-welcome-brightgreen.svg" alt="Issues welcome"></a>
</p>

> 只读真实项目，按中国版权保护中心的交存要求，把软件著作权登记所需的申请表、操作手册、源代码鉴别材料三份 Word 一次生成到位。适合 Claude Code / Codex / Cursor / 豆包 等 Agent 直接调用。

---

## ✨ 亮点特性

与同类软著材料工具最大的不同：**把系统截图当作唯一事实来源**。

| 亮点 | 说明 |
|---|---|
| 📸 截图驱动 · 不无中生有 | 手册章节跟着截图走，截图上没有的功能一个字都不写；每个功能小节必须有图，截图页面名必须在正文出现 |
| 📚 代码全量收录 | 有多少代码写多少，不砍页数、不截断；目录按用户手册章节排版（首页…第 1 页） |
| 🧮 源程序量多口径统计 | 原始总行数 / 有效行数 / 自研有效行数，三种口径一次算清 |
| 🔍 开发环境自动提取 | 从 pom.xml / package.json / Dockerfile / application.yml 自动识别开发环境，不用手填 |
| 🧩 客户模板优先 | 有客户申请表模板就严格照模板填（不无中生有），没有才用内置空白模板 |
| 🔤 字段与菜单自动提取 | 从路由、SQL 与 .vue 中读取真实菜单顺序、表格列名、按钮名，写进手册的都是真字段 |
| ✅ 提交前自动自检 | 名称/版本/权利人三处一字不差、代码页数下限 60、截图-章节一一对应，漏了直接拦住你 |
| 🖥️ 双平台转 PDF | Windows 用 `docx_to_pdf.ps1`（自动探测 soffice），Linux/macOS 用 `docx_to_pdf.sh` |
| 🏷️ 截图批量重命名 | 一键把截图规范成 `NN-页面名.png`，命名乱也不怕 |

---

## 🚀 快速开始

**依赖**：Python 3.10+；LibreOffice（仅转 PDF 用，可选）。

```bash
git clone https://github.com/xiaohei52056/software-copyright-skill.git
```

把目录放进你的 Agent 的 skills 目录（Claude Code `~/.claude/skills/`、Codex `~/.codex/skills/`、Cursor / 豆包 各自的 skills 目录，或直接把 `SKILL.md` 内容贴进项目规则）。

然后在项目里对 Agent 说一句（把尖括号内容替换成你的实际情况）：

> 项目路径：<项目根目录>，目录内应包含：源代码目录、系统功能截图、客户申请表模板 等其他材料模板
> 为当前项目生成软件著作权申请资料

想把关键信息一次说全、让 Agent 少问一轮，可以这样：

> 项目路径：D:\projects\<项目名>（内含 源码/、截图/、申请表模板.docx 等）
> 为当前项目生成软件著作权申请资料，软件全称：<软件全称>，版本号：<版本号>，著作权人：<权利人名称>

Agent 会按 `SKILL.md` 的流程自动推进：**收集信息与业务理解 → 确认代码范围与章节映射 → 截图驱动写手册 → 生成三份 Word 并自检**，全程只在 3 个节点停下来等你确认。产物统一输出到项目根目录 `软件著作权申请资料/`。

> 不依赖 Agent 也行：所有脚本可直接用命令行跑，见 `SKILL.md` 各步命令。

---

## 📂 目录结构

```
SKILL.md                      Agent 读的工作流与硬规则
references/
  格式要求.md                 版权中心交存格式 + 实操防打回要点
  申请表字段说明.md           21 个字段怎么填；主要功能详版（500–1300 字）；源程序量口径
  操作手册写作规范.md         截图驱动规则、菜单顺序、字段提取、措辞要求
templates/申请表模板.docx     内置空白申请表（脚本按单元格填写）
scripts/
  collect_sources.py          扫项目 → 候选文件 + 源程序量多口径 + 自研/框架归属
  extract_env.py              从项目配置自动提取开发环境
  extract_menu.py             从路由/SQL 提取系统菜单顺序
  extract_fields.py           从 .vue 提取表格列名/按钮名/表单字段
  scan_screenshots.py         扫描截图生成截图清单（手册章节唯一依据）
  rename_screenshots.py       截图批量重命名 NN-页面名.png（--dry-run 预览）
  detect_existing.py          检测项目下已有软著材料
  build_code_docx.py          代码 Word（50 行/页、全量收录、封面+目录页按手册章节、索引）
  build_manual_docx.py        手册 Markdown → Word（正文顶格、标题编号风格可选）
  fill_application_form.py    yaml → 申请表 docx + txt（模板优先、主要功能详版）
  check_consistency.py        提交前自检（截图唯一事实源检查、柔性警告）
  docx_to_pdf.sh / .ps1       批量转 PDF（Linux/macOS / Windows）
```

---

---

## ⚠️ 常见问题与限制

- 手册目录是 Word 域：用 Word 打开后全选按 F9 更新域再导出 PDF；LibreOffice 转 PDF 时也会按 updateFields 自动更新目录（实测有效）。
- 手册 Markdown 只支持标题、段落、编号行、管道表、`![图注](路径)` 图片、`【截图预留】` 占位。
- 代码全量收录时页数可能很大（1 万行 ≈ 200 页），docx 体积随之增长，转 PDF 后请核对 `代码文档索引.md`。
- 菜单顺序提取依赖路由文件或 sys_menu SQL，无法提取时回退业务理解模块清单并明确标注。
- Windows 转 PDF 若弹出"正在等待打印机连接"：把默认打印机临时设为 Microsoft Print to PDF（转完再恢复），或弹窗时点"取消"。
- 格式规则以中国版权保护中心官网当期说明为准；本仓库规则整理于 2026 年 8–9 月。

---

## 🙏 致谢与来源

本技能由 [xiaohei52056](https://github.com/xiaohei52056) 创建并持续维护，核心工作流与格式规则**改进自** [pkm365/ruanzhu-skill](https://github.com/pkm365/ruanzhu-skill)（MIT License，© 2026 pkm365），并保留其原 LICENSE 与版权声明。

在原项目基础上，本仓库重点贡献了以下改进：

| 改进点 | 说明 |
|---|---|
| 截图唯一事实源 | 手册内容 100% 由截图决定，截图没有的功能一律不写，杜绝无中生有 |
| 代码全量收录 + 章节目录 | 不再截断（前 30/后 30 页），有多少写多少；目录按用户手册章节生成 |
| 源程序量多口径统计 | 原始 / 有效 / 自研有效三口径 |
| 开发环境自动提取 | pom.xml / package.json / Dockerfile 自动识别 |
| 字段与菜单自动提取 | 从路由、SQL、.vue 读取真实字段名与按钮名 |
| 客户模板优先 | 有客户申请表模板严格照填，无则内置空白模板 |
| 提交前自检增强 | 每个功能小节必须有图、截图页面名必须在正文出现、代码页数下限 60 |
| 双平台转 PDF | Windows `.ps1`（自动探测 soffice）+ Linux/macOS `.sh` |
| 截图批量重命名 | 一键规范为 `NN-页面名.png` |

原项目作者对本衍生项目无任何背书。

---

## 📄 License

MIT License。版权行：`Copyright (c) 2026 xiaohei52056`（本仓库创建者与主要贡献者）+ `Copyright (c) 2026 pkm365`（原项目作者，MIT 要求保留原声明）。详见 [LICENSE](LICENSE)。
