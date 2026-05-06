#!/usr/bin/env python3
"""
daily_pipeline.py - 日报全自动编排脚本 v3

由 cron 每日定时触发，按顺序执行:
  1. 三处室综合采集 (enterprise_crawler_v3.py)
  2. 三个处室简报生成 (generate_brief_v2.py)
  3. 邮件发送 (step3_send.py)

用法:
  python3 daily_pipeline.py --date 2026-04-28
  python3 daily_pipeline.py
"""

import json
import os
import sys
import subprocess
import time
import argparse
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")
DATA_DIR = os.path.join(BASE_DIR, "data")


def ensure_date(date_str=None):
    if date_str:
        return date_str
    return datetime.now().strftime("%Y-%m-%d")


def run_cmd(cmd, timeout=300, label="step", env=None):
    """run shell command and stream key output"""
    print("\n{}".format("=" * 60))
    print("  {} {}".format(label, " ".join(cmd)))
    print("{}".format("=" * 60))

    proc_env = env if env else None
    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, env=proc_env)

    if proc.returncode != 0:
        print("  FAILED (rc={})".format(proc.returncode))
        for line in proc.stderr.strip().split("\n")[:10]:
            if line.strip():
                print("  Error: {}".format(line.strip()))
        return False

    for line in proc.stdout.split("\n"):
        s = line.strip()
        if any(m in s for m in ["\u2705", "\u274c", "\u26a0\ufe0f",
                                "\U0001f4e1", "\U0001f4dd", "\U0001f4c4",
                                "\U0001f4cb", "\U0001f389", "\U0001f525",
                                "\u6210\u529f", "\u5931\u8d25",
                                "\u5b8c\u6210", "TODO"]):
            if s:
                print("  {}".format(s))
        elif any(kw in s for kw in ["\u91c7\u96c6\u5b8c\u6210",
                                     "\u5168\u6d41\u7a0b",
                                     "\u7b80\u62a5", "\u8fd0\u884c\u62a5\u544a"]):
            if s:
                print("  {}".format(s))

    return True


def step1_crawl_all(date_str):
    """step1: enterprise_crawler_v3 three depts crawl (news only)"""
    print("\n{}".format("=" * 60))
    print("  Step 1: news crawl (v3 - news only)")
    print("{}".format("=" * 60))

    script = os.path.join(SCRIPTS_DIR, "enterprise_crawler_v3.py")
    cmd = ["python3", script, "--date", date_str, "--departments", "工业互联网处", "医药处", "原材料_金属_非金属处"]
    return run_cmd(cmd, timeout=600, label="enterprise_crawler_v3")


def step1b_sentiment(date_str):
    """step1b: sentiment search via Baidu API + LLM"""
    print("\n{}".format("=" * 60))
    print("  Step 1b: sentiment search (step_sentiment_v3)")
    print("{}".format("=" * 60))

    script = os.path.join(SCRIPTS_DIR, "step_sentiment_v3.py")
    cmd = ["python3", script, "--date", date_str]
    return run_cmd(cmd, timeout=1800, label="step_sentiment_v3")


def step2_generate_brief(date_str):
    """step2: generate template-based .docx briefs"""
    print("\n{}".format("=" * 60))
    print("  Step 2: generate template briefs (v3)")
    print("{}".format("=" * 60))

    # 从 runtime_config.json 加载 API key
    config_path = os.path.join(BASE_DIR, "runtime_config.json")
    env = os.environ.copy()
    if os.path.exists(config_path):
        with open(config_path, encoding="utf-8") as f:
            cfg = json.load(f)
        if "DEEPSEEK_API_KEY" in cfg:
            env["DEEPSEEK_API_KEY"] = cfg["DEEPSEEK_API_KEY"]

    script = os.path.join(SCRIPTS_DIR, "generate_brief_v3.py")
    cmd = ["python3", script, "--date", date_str]
    return run_cmd(cmd, timeout=600, label="generate_brief_v3")


def step3_send_email(config, date_str):
    """step3: send briefs via email"""
    industry = config["industry"]

    script = os.path.join(SCRIPTS_DIR, "step3_send.py")
    email_cfg = os.path.join(CONFIG_DIR, "email.json")

    brief_dir = os.path.join(DATA_DIR, date_str, "brief")
    if os.path.isdir(brief_dir):
        docx_files = sorted([f for f in os.listdir(brief_dir) if f.endswith(".docx")])
        if docx_files:
            print("\n{}".format("=" * 60))
            print("  Step 3: send email")
            print("{}".format("=" * 60))
            cmd = ["python3", script, "--dir", brief_dir,
                   "--config", email_cfg, "--date", date_str]
            return run_cmd(cmd, timeout=60, label="send email")
        else:
            print("  WARNING: no .docx files found in {}".format(brief_dir))
            return False
    else:
        print("  WARNING: brief dir not found: {}".format(brief_dir))
        return False


def run_pipeline_all(date_str):
    """run full pipeline: crawl -> brief -> email"""
    print("\n{}".format("#" * 60))
    print("  # Pipeline v3: crawl + brief + email")
    print("  # date: {}".format(date_str))
    print("{}".format("#" * 60))

    start_time = datetime.now()
    errors = []

    if not step1_crawl_all(date_str):
        errors.append("news crawl failed")
    else:
        print("  news crawl done")

    if not step1b_sentiment(date_str):
        errors.append("sentiment crawl failed")
    else:
        print("  sentiment crawl done")

    if not step2_generate_brief(date_str):
        errors.append("brief generation failed")
    else:
        print("  brief generation done")

    step3_send_email({"industry": "all"}, date_str)

    elapsed = (datetime.now() - start_time).total_seconds()

    print("\n{}".format("#" * 60))
    print("  Pipeline complete! ({}s)".format(elapsed))
    if errors:
        print("  FAILURES: {}".format(", ".join(errors)))
    else:
        print("  All steps succeeded")
    print("  Data: {}".format(os.path.join(DATA_DIR, date_str, "raw")))
    print("  Briefs: {}".format(os.path.join(DATA_DIR, date_str, "brief")))
    print("  Report: {}".format(os.path.join(DATA_DIR, date_str, "archive",
                                             "report.md")))
    print("{}".format("#" * 60))

    return len(errors) == 0


def main():
    parser = argparse.ArgumentParser(description="daily pipeline v3")
    parser.add_argument("--date", help="date YYYY-MM-DD (default: today)")
    parser.add_argument("--step", choices=["1", "1b", "2", "3"],
                        help="single step: 1=news crawl, 1b=sentiment, 2=brief, 3=email")
    parser.add_argument("--skip-llm", action="store_true",
                        help="skip LLM summary in brief generation")
    args = parser.parse_args()

    date_str = ensure_date(args.date)

    if args.skip_llm:
        os.environ["SKIP_LLM"] = "1"

    if args.step == "1":
        step1_crawl_all(date_str)
    elif args.step == "1b":
        step1b_sentiment(date_str)
    elif args.step == "2":
        step2_generate_brief(date_str)
    elif args.step == "3":
        step3_send_email({"industry": "all"}, date_str)
    else:
        run_pipeline_all(date_str)


if __name__ == "__main__":
    main()
