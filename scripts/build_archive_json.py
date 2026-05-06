#!/usr/bin/env python3
"""Build all_articles.json from raw news_*.txt files in raw directory."""
import json, os, sys, re
from datetime import datetime

date_str = sys.argv[1] if len(sys.argv) > 1 else datetime.now().strftime("%Y-%m-%d")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
raw_dir = os.path.join(BASE, "data", date_str, "raw")
archive_dir = os.path.join(BASE, "data", date_str, "archive")
os.makedirs(archive_dir, exist_ok=True)

# Map source filename to department
DEPT_MAP = {}
# Read all config files to map sources
config_dir = os.path.join(BASE, "config")
if os.path.exists(config_dir):
    for f in os.listdir(config_dir):
        if f.endswith(".json") and f != "email.json":
            try:
                with open(os.path.join(config_dir, f), encoding="utf-8") as fp:
                    cfg = json.load(fp)
                if "sources" in cfg:
                    dept_name = cfg.get("department_name", f.replace(".json", ""))
                    for src in cfg["sources"]:
                        name = src.get("name", src.get("source_name", ""))
                        if name:
                            DEPT_MAP[name] = dept_name
            except:
                pass

# Fallback: manually assign
if not DEPT_MAP:
    DEPT_MAP = {
        "工信部": "工业互联网处",
        "贵州省工信厅": "工业互联网处",
        "国家统计局": "工业互联网处",
        "工信部电子信息司": "工业互联网处",
        "中国电子信息产业发展研究院": "工业互联网处",
        "国家市场监督管理总局": "工业互联网处",
        "中国信息通信研究院": "工业互联网处",
        "中国半导体行业协会": "工业互联网处",
        "中国电子元件行业协会": "工业互联网处",
        "中国电子视像行业协会": "工业互联网处",
        "中国电子学会": "工业互联网处",
        "人民日报": "工业互联网处",
        "CCTV": "工业互联网处",
        "国家卫健": "医药处",
        "中国中医药网": "医药处",
        "医药经济报": "医药处",
    }

articles_by_dept = {}

for fname in sorted(os.listdir(raw_dir)):
    if not fname.startswith("news_"):
        continue
    fpath = os.path.join(raw_dir, fname)
    with open(fpath, encoding="utf-8") as f:
        content = f.read()

    # Parse articles separated by "---" (only present if explicit separator)
    blocks = [b.strip() for b in content.split("\n---\n") if b.strip()]

    for block in blocks:
        item = {"title": "", "date": date_str, "source_name": "", "link": "", "summary": ""}
        lines = block.strip().split("\n")
        for line in lines:
            if line.startswith("标题:"):
                item["title"] = line[3:].strip()
            elif line.startswith("日期:"):
                item["date"] = line[3:].strip()
            elif line.startswith("来源:"):
                src = line[3:].strip()
                item["source_name"] = src
            elif line.startswith("链接:"):
                item["link"] = line[3:].strip()
            elif line.startswith("摘要:"):
                item["summary"] = line[3:].strip()[:500]
            elif line.startswith("信息类型:"):
                item["info_type"] = line[5:].strip()
            # Handle when summary doesn't have explicit prefix (multiline)
            elif not any(line.startswith(p) for p in ["标题:", "日期:", "来源:", "链接:", "摘要:", "信息类型:"]):
                if item.get("summary"):
                    item["summary"] += " " + line.strip()[:200]
                elif item.get("title"):
                    item["summary"] = line.strip()[:500]

        src_name = item.get("source_name", "")
        dept = DEPT_MAP.get(src_name, "工业互联网处")
        articles_by_dept.setdefault(dept, []).append(item)

# Write output
out_path = os.path.join(archive_dir, "all_articles.json")
with open(out_path, "w", encoding="utf-8") as f:
    json.dump(articles_by_dept, f, ensure_ascii=False, indent=2)

total = sum(len(v) for v in articles_by_dept.values())
print(f"✅ all_articles.json written: {len(articles_by_dept)} depts, {total} articles")
for d, arts in articles_by_dept.items():
    print(f"   {d}: {len(arts)} articles")
