#!/usr/bin/env python3
"""
_agent_browser_fetch.py - agent-browser 辅助采集脚本

使用 agent-browser CLI 访问需要交互的页面（翻页、搜索、筛选等），
提取文章列表（标题+链接+日期+摘要），以 JSON Lines 格式输出。

用法：
  python3 _agent_browser_fetch.py --url https://web.csia.net.cn/rdxw --label 半导体协会-热点新闻
  python3 _agent_browser_fetch.py --url https://example.com --label 某网站 --scroll 3

输出格式（每行一条 JSON）：
  {"title": "...", "date": "2026-04-26", "source_name": "...", "link": "...", "abstract": ""}
"""

import argparse
import json
import os
import re
import subprocess
import sys
import time
from datetime import datetime, timedelta
from urllib.parse import urlparse


def run_agent_browser(command_args, timeout=30):
    """执行 agent-browser 命令"""
    cmd = ["agent-browser"] + command_args
    proc = subprocess.run(
        cmd, capture_output=True, text=True, timeout=timeout
    )
    return proc.returncode, proc.stdout, proc.stderr


def fetch_page_list(url, label, scroll_times=0, max_items=40):
    """
    使用 agent-browser 访问页面，提取文章列表
    """
    results = []
    seen_links = set()
    source_name = label or urlparse(url).netloc

    try:
        # 1. 打开页面
        # print(f"  🔍 agent-browser: 打开 {url}", file=sys.stderr)
        rc, out, err = run_agent_browser(["open", url])
        if rc != 0:
            print(f"    ⚠️  打开页面失败: {err[:200]}", file=sys.stderr)
            return results
        time.sleep(2)

        # 2. 获取页面标题（用作日期参考）
        rc, title_out, _ = run_agent_browser(["get", "title"])
        page_title = title_out.strip() if rc == 0 else ""

        # 3. 滚动页面（触发懒加载）
        for s in range(scroll_times):
            print(f"    📜 滚动 #{s+1}", file=sys.stderr)
            run_agent_browser(["scroll", "down", "800"], timeout=10)
            time.sleep(1.5)

        # 4. 获取完整页面 HTML（通过 eval）
        rc, html_out, _ = run_agent_browser([
            "eval", "document.documentElement.outerHTML"
        ], timeout=15)
        if rc != 0:
            print(f"    ⚠️  获取 HTML 失败", file=sys.stderr)
            # 尝试用 snapshot 替代
            rc, snap_out, _ = run_agent_browser(["snapshot", "-c"], timeout=15)
            html_out = snap_out if rc == 0 else ""

        html = html_out.strip()

        # 5. 从 HTML 中提取文章链接
        items = _extract_articles_from_html(html, url)
        for item in items:
            link = item["link"]
            if link and link not in seen_links:
                seen_links.add(link)
                item["source_name"] = source_name
                results.append(item)

        # 6. 如果上面没提取到，尝试用 snapshot 文本分析
        if not results:
            print(f"    🔄 HTML 解析无结果，尝试 snapshot 分析", file=sys.stderr)
            rc, snap_out, _ = run_agent_browser(["snapshot", "-c"], timeout=15)
            if rc == 0:
                items = _extract_from_snapshot(snap_out, url)
                for item in items:
                    link = item["link"]
                    if link and link not in seen_links:
                        seen_links.add(link)
                        item["source_name"] = source_name
                        results.append(item)

        # 7. 尝试翻页（检测"下一页"按钮）
        page_num = 1
        while len(results) < max_items:
            page_num += 1
            has_next = _try_next_page()
            if not has_next:
                break
            print(f"    📄 翻到第 {page_num} 页", file=sys.stderr)
            time.sleep(2)

            rc, html_out, _ = run_agent_browser([
                "eval", "document.documentElement.outerHTML"
            ], timeout=15)
            if rc != 0:
                break

            items = _extract_articles_from_html(html_out.strip(), url)
            new_count = 0
            for item in items:
                link = item["link"]
                if link and link not in seen_links:
                    seen_links.add(link)
                    item["source_name"] = source_name
                    results.append(item)
                    new_count += 1
            if new_count == 0:
                break

    except subprocess.TimeoutExpired:
        pass
        # print(f"    ⚠️  agent-browser 超时", file=sys.stderr)
    except Exception as e:
        print(f"    ❌ agent-browser 异常: {e}", file=sys.stderr)
    finally:
        try:
            run_agent_browser(["close"], timeout=5)
        except Exception:
            pass

    return results[:max_items]


def _extract_articles_from_html(html, base_url):
    """从 HTML 中提取文章列表"""
    items = []
    today = datetime.now().strftime("%Y-%m-%d")

    # 找日期
    dates_found = re.findall(r'(\d{4}[-/]\d{1,2}[-/]\d{1,2})', html)
    dates_found = list(dict.fromkeys(dates_found))

    # 策略1: 数字ID的 .html 链接（常见于新闻列表）
    article_links = re.findall(
        r'<a[^>]*href=["\']([^"\']*?/\d+\.html)["\'][^>]*>\s*(.*?)\s*</a>',
        html, re.IGNORECASE | re.DOTALL
    )

    for href_raw, title_raw in article_links:
        title = re.sub(r'<[^>]+>', '', title_raw).strip()
        if len(title) <= 5:
            continue
        if _is_nav_text(title):
            continue
        href = _normalize_url(href_raw.strip(), base_url)
        if not href:
            continue
        date = _find_nearby_date(html, href_raw, dates_found, today)
        items.append({"title": title, "date": date, "link": href})

    # 策略2: 通用链接
    if not items:
        link_pattern = re.compile(
            r'<a[^>]*href=["\']([^"\']+)["\'][^>]*>\s*([^<]{8,60})\s*</a>',
            re.IGNORECASE | re.DOTALL
        )
        for href_raw, title_raw in link_pattern.findall(html):
            title = re.sub(r'<[^>]+>', '', title_raw).strip()
            if len(title) < 6:
                continue
            if _is_nav_text(title):
                continue
            href = _normalize_url(href_raw.strip(), base_url)
            if not href:
                continue
            date = _find_nearby_date(html, href_raw, dates_found, today)
            items.append({"title": title, "date": date, "link": href})

    return items


def _extract_from_snapshot(snap_text, base_url):
    """从 accessibility tree snapshot 中提取链接"""
    items = []
    today = datetime.now().strftime("%Y-%m-%d")

    # snapshot 格式: "链接 xxx" 或 link "xxx"
    link_pattern = re.findall(r'(?:链接|link)\s+([^\n]+)', snap_text)
    seen = set()
    for text in link_pattern:
        text = text.strip().strip('"').strip("'")
        if len(text) < 6 or text in seen:
            continue
        seen.add(text)
        items.append({"title": text, "date": today, "link": ""})

    # 如果 snaphost 中有 url，优先使用
    url_pattern = re.findall(r'https?://[^\s<>"\']+', snap_text)
    for i, url in enumerate(url_pattern[:len(items)]):
        items[i]["link"] = url

    return items


def _try_next_page():
    """尝试点击"下一页"按钮"""
    try:
        # 先获取页面快照找到翻页按钮
        rc, out, _ = run_agent_browser(["snapshot", "-i", "--json"], timeout=10)
        if rc != 0:
            return False

        # 尝试用 find 语义定位器点击"下一页"
        for text in ["下一页", "下一頁", "Next", "next", "»", "›", ">"]:
            rc, _, _ = run_agent_browser([
                "find", "text", text, "click"
            ], timeout=10)
            if rc == 0:
                time.sleep(2)
                return True
        return False
    except Exception:
        return False


def _normalize_url(href, base_url):
    """标准化URL"""
    if href.startswith('javascript:') or href.startswith('#'):
        return None
    if href.startswith('//'):
        return 'https:' + href
    if href.startswith('/'):
        parsed = urlparse(base_url)
        return f"{parsed.scheme}://{parsed.netloc}{href}"
    if href.startswith('http'):
        return href
    return None


def _is_nav_text(text):
    """判断是否为导航/功能文字"""
    navs = ['首页', '上一页', '下一页', '末页', '加入协会', '关于我们',
            '网站地图', '联系我们', '设为首页', '收藏本站', '登录', '注册',
            'English', '设为首页', '加入收藏', 'RSS', '微博', '微信']
    return text in navs or len(text) > 80


def _find_nearby_date(html, href, dates_found, today):
    """在链接附近找日期"""
    if dates_found:
        return dates_found[0].replace('/', '-')
    return today


def main():
    parser = argparse.ArgumentParser(description="agent-browser 新闻列表采集")
    parser.add_argument("--url", required=True, help="目标页面 URL")
    parser.add_argument("--label", default="", help="来源标签")
    parser.add_argument("--scroll", type=int, default=0, help="滚动次数（触发懒加载）")
    parser.add_argument("--max", type=int, default=40, help="最大条目数")
    args = parser.parse_args()

    results = fetch_page_list(args.url, args.label, args.scroll, args.max)

    for item in results:
        print(json.dumps(item, ensure_ascii=False))

    return 0


if __name__ == "__main__":
    sys.exit(main())
