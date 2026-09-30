#!/usr/bin/env python3
"""
从项目真实配置文件提取开发环境信息（R003），替代凭空猜测。

用法:
  python3 extract_env.py <项目根目录> [-o 草稿/环境信息.md] [--json 草稿/environment.json]

提取来源:
  - pom.xml      → Java 版本（maven.compiler.source/release）、Spring Boot 版本、核心依赖版本
  - package.json → Vue / Vite / TypeScript / Element Plus 版本
  - Dockerfile / docker-compose.yml → 基础镜像、容器化部署方式
  - application*.yml / application*.properties → 数据库类型、Redis、MQTT 等中间件
  - .idea/ 或目录线索 → 推断 IDE

输出全部标注"自动提取，待确认"；无法提取的字段留空并提示"待用户补充"，绝不猜测。
"""
import argparse, json, re, sys
from pathlib import Path

def first_existing(root: Path, names):
    for n in names:
        p = root / n
        if p.exists(): return p
    # 递归查找（最多两层）
    for depth in (1, 2):
        for p in root.glob("*" * depth):
            if p.name in names: return p
    return None

def read(p: Path) -> str:
    try:
        return p.read_text(encoding="utf-8", errors="replace")
    except Exception:
        return ""

def extract_pom(text: str) -> dict:
    out = {}
    m = re.search(r"<maven\.compiler\.(source|release)>\s*([^<]+)", text)
    if m: out["java_version"] = m.group(2).strip()
    m = re.search(r"<java\.version>\s*([^<]+)", text)
    if m: out["java_version"] = m.group(2).strip()
    m = re.search(r"<spring-boot[^>]*>\s*<version>\s*([^<]+)", text)
    if m: out["spring_boot_version"] = m.group(1).strip()
    m = re.search(r"<artifactId>([^<]*(?:mybatis|redis|kafka|rabbit|nacos|dubbo|mqtt)[^<]*)</artifactId>\s*<version>\s*([^<]+)", text, re.I)
    if m: out.setdefault("deps", []).append(f"{m.group(1).strip()}:{m.group(2).strip()}")
    m = re.search(r"<artifactId>([^<]+)</artifactId>\s*<version>\s*([^<]+)", text)
    if m: out.setdefault("parent_artifact", f"{m.group(1).strip()}:{m.group(2).strip()}")
    return out

def extract_package_json(text: str) -> dict:
    out = {}
    for key, label in (("vue", "vue_version"), ("vite", "vite_version"),
                       ("typescript", "typescript_version"), ("element-plus", "element_plus_version"),
                       ("pinia", "pinia_version"), ("vue-router", "vue_router_version")):
        m = re.search(rf'"{key}"\s*:\s*"([^"]+)"', text)
        if m: out[label] = m.group(1)
    return out

def extract_docker(text: str) -> dict:
    out = {}
    m = re.search(r"FROM\s+([^\s]+)", text)
    if m: out["base_image"] = m.group(1)
    if re.search(r"docker-compose|services:|image:", text, re.I):
        out["deploy_mode"] = "docker-compose 容器化部署"
    elif re.search(r"FROM\s", text):
        out["deploy_mode"] = "Docker 镜像部署"
    return out

def extract_app_yml(text: str) -> dict:
    out = {}
    m = re.search(r"url:\s*jdbc:(mysql|postgresql|oracle|sqlserver|dm|kingbase|tidb|mariadb)[^:]*", text, re.I)
    if m: out["database"] = m.group(1).upper()
    m = re.search(r"url:\s*jdbc:[^:]+://[^/]+", text)
    if m and "database" not in out: out["database"] = m.group(0).split(":")[2] if ":" in m.group(0) else "DB"
    if re.search(r"(^|\n)\s*redis:", text) or re.search(r"redis:", text, re.I): out["cache"] = "Redis"
    if re.search(r"mqtt", text, re.I): out["mqtt"] = "MQTT"
    if re.search(r"kafka", text, re.I): out["mq"] = "Kafka"
    if re.search(r"rabbit", text, re.I): out["mq"] = "RabbitMQ"
    m = re.search(r"nacos", text, re.I)
    if m: out["registry"] = "Nacos"
    return out

def detect_ide(root: Path, env: dict):
    if (root / ".idea").is_dir(): env["ide"] = "IntelliJ IDEA"
    elif (root / ".vscode").is_dir(): env["ide"] = "Visual Studio Code"
    elif (root / "pom.xml").exists(): env.setdefault("ide", "IntelliJ IDEA（检测到 Maven 工程）")
    elif (root / "package.json").exists(): env.setdefault("ide", "Visual Studio Code（检测到前端工程）")

def build_md(root: Path, env: dict, miss: list) -> str:
    lines = ["# 环境信息（自动提取，待确认）", "",
             f"扫描根目录：`{root}`", "",
             "> 以下信息由脚本从项目配置文件中自动提取（R003），仅供填写申请表参考，**必须人工核对**；无法提取的字段请用户补充。",
             ""]
    def row(k, label, v):
        if v:
            lines.append(f"- **{label}**：{v}")
    row("java_version", "Java 版本", env.get("java_version"))
    row("spring_boot_version", "Spring Boot 版本", env.get("spring_boot_version"))
    row("vue_version", "Vue 版本", env.get("vue_version"))
    row("vite_version", "Vite 版本", env.get("vite_version"))
    row("typescript_version", "TypeScript 版本", env.get("typescript_version"))
    row("element_plus_version", "Element Plus 版本", env.get("element_plus_version"))
    row("pinia_version", "Pinia 版本", env.get("pinia_version"))
    row("vue_router_version", "Vue Router 版本", env.get("vue_router_version"))
    row("base_image", "Docker 基础镜像", env.get("base_image"))
    row("deploy_mode", "部署方式", env.get("deploy_mode"))
    row("database", "数据库", env.get("database"))
    row("cache", "缓存", env.get("cache"))
    row("mqtt", "MQTT 接入", env.get("mqtt"))
    row("mq", "消息队列", env.get("mq"))
    row("registry", "注册中心", env.get("registry"))
    row("ide", "IDE", env.get("ide"))
    if env.get("deps"):
        lines.append(f"- **核心依赖**：{ '；'.join(env['deps']) }")
    if miss:
        lines += ["", "## 未能自动提取（待用户补充）", ""]
        lines += [f"- {x}" for x in miss]
    lines += ["", "## 填写到申请表的建议", "",
              "- 软件开发环境或开发工具：IDE + 构建工具 + 数据库客户端（如 IntelliJ IDEA + Maven + MySQL）",
              "- 软件运行支撑环境或支持软件：数据库 + JDK/Node 运行时 + 浏览器（如 MySQL 8.0、JDK 17、Chrome）",
              "- 运行平台/操作系统：按真实部署环境填（Linux/Windows Server/Docker），不能照抄本文件默认值。"]
    return "\n".join(lines) + "\n"

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("root")
    ap.add_argument("-o", "--out", default=None)
    ap.add_argument("--json", default=None)
    a = ap.parse_args()
    root = Path(a.root).resolve()
    env = {}

    pom = first_existing(root, ["pom.xml"])
    if pom:
        env.update(extract_pom(read(pom)))
        env["build_tool"] = "Maven"
    pkg = first_existing(root, ["package.json"])
    if pkg:
        env.update(extract_package_json(read(pkg)))
        env.setdefault("build_tool", "npm/pnpm/yarn")
    dk = first_existing(root, ["Dockerfile", "docker-compose.yml", "docker-compose.yaml"])
    if dk:
        env.update(extract_docker(read(dk)))
    app_yml = first_existing(root, ["application.yml", "application.yaml", "application-dev.yml",
                                    "application-dev.yaml", "application-prod.yml", "application.properties"])
    if app_yml:
        env.update(extract_app_yml(read(app_yml)))
    detect_ide(root, env)

    miss = []
    for label, key in (("Java 版本（检查 pom.xml maven.compiler.source）", "java_version"),
                       ("构建工具", "build_tool"),
                       ("数据库类型（检查 application.yml 的 jdbc url）", "database"),
                       ("缓存 / 消息中间件（Redis/Kafka/MQTT 等）", "cache"),
                       ("部署方式（Dockerfile / docker-compose）", "deploy_mode")):
        if key not in env or not env.get(key):
            miss.append(label)

    out = a.out or "环境信息.md"
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    Path(out).write_text(build_md(root, env, miss), encoding="utf-8")
    print(f"✅ {out}")
    if a.json:
        Path(a.json).write_text(json.dumps({"root": str(root), "env": env, "missing": miss},
                                           ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"✅ {a.json}")
    for k, v in env.items():
        print(f"  {k} = {v}")
    if miss:
        print("⚠️  未能自动提取：" + "；".join(miss))

if __name__ == "__main__":
    main()
