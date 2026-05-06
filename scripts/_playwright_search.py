#!/usr/bin/env python3
"""
_playwright_search.py - playwright 搜索辅助脚本
"""
import argparse, json, sys
from datetime import datetime
from urllib.parse import quote

def search_baidu_news(keyword, max_results=30):
    results = []
    url = f"https://news.baidu.com/ns?word={quote(keyword)}&pn=0&rn={max_results}&cl=2&ct=1&tn=newstitle"
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
            page = context.new_page()
            page.goto(url, timeout=30000, wait_until="domcontentloaded")
            page.wait_for_timeout(3000)
            items = page.evaluate("""
                () => {
                    const r = [];
                    document.querySelectorAll('.result h3 a, h3 a, .c-title a, .news-title a').forEach(a => {
                        const t = a.textContent.trim();
                        if (t && t.length > 3 && t.length < 100) r.push({title: t, link: a.href || ''});
                    });
                    return r.slice(0, 30);
                }
            """)
            browser.close()
            today = datetime.now().strftime("%Y-%m-%d")
            for item in items:
                t = item.get("title","").strip()
                if t and len(t) > 3 and len(t) < 80:
                    results.append({"title": t, "date": today, "source_name": "百度新闻", "link": item.get("link",""), "abstract": ""})
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
    return results

parser = argparse.ArgumentParser()
parser.add_argument("--target", required=True, choices=["baidu_news"])
parser.add_argument("--keyword", required=True)
args = parser.parse_args()
for item in search_baidu_news(args.keyword):
    print(json.dumps(item, ensure_ascii=False))
