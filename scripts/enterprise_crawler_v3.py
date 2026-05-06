#!/usr/bin/env python3
"""
企业舆情与新闻采集系统 v6

v6 关键改进：
1. 百度搜索结果后处理：LLM 智能过滤不相关内容（通过 analyze_baidu_results 函数）
2. 舆情搜索关键词优化：结合配置中的 info_type 和 remark 字段，动态构建高级搜索查询
3. 使用 baidu-search 的 freshness 时间过滤确保结果时效性
"""

import sys
import json
import os
import re
import time
import random
import subprocess
import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from typing import List, Dict, Set, Tuple, Optional
from urllib.parse import urlparse, urljoin
from lxml import etree

# ===================== 路径常量 =====================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
DATA_DIR = os.path.join(BASE_DIR, "data")


USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:122.0) Gecko/20100101 Firefox/122.0",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
]

DEFAULT_TIME_RANGE_DAYS = 7
DEPARTMENTS = ["工业互联网处", "医药处", "原材料_金属_非金属处"]

DEPT_KEYWORDS = {
    "工业互联网处": {"电子信息", "半导体", "芯片", "集成电路", "5G", "6G", "通信",
                     "互联网", "大数据", "人工智能", "AI", "云计算", "物联网",
                     "工业互联网", "数字化", "智能制造", "软件", "信息技术",
                     "电子制造", "电子元件", "电子器件", "显示", "光电子",
                     "传感器", "数据中心", "量子", "算力", "信息化", "数字",
                     "光伏", "锂电", "电池", "储能", "电子"},
    "医药处": {"医药", "药品", "生物", "医疗", "健康", "医院", "药监", "医保",
               "集采", "中医药", "中药", "创新药", "仿制药", "疫苗", "诊断",
               "器械", "医疗器械", "卫健", "公共卫生", "临床", "药企", "制药"},
    "原材料_金属_非金属处": {"原材料", "金属", "非金属", "有色金属", "钢铁", "铝",
                             "铜", "锌", "镍", "锡", "铅", "锂", "钴", "稀土",
                             "矿产", "矿石", "采矿", "冶炼", "化工", "磷",
                             "铝土矿", "氧化铝", "电解铝", "钢材", "产能",
                             "大宗商品", "新材料", "电池材料", "碳酸锂",
                             "绿色矿山", "矿权", "黄磷", "磷酸"},
}

SPECIALIZED_SOURCES = [
    "工信部电子信息司", "中国电子信息产业发展研究院", "中国信息通信研究院",
    "中国半导体行业协会", "中国电子元件行业协会", "中国电子视像行业协会",
    "中国电子学会", "中国有色金属工业网", "医药经济报",
    "上海钢联SMM", "生意社", "阿拉丁中营网-铝产业链综合服务平台及生态系统",
    "百川盈孚", "隆众资讯", "电池中国网",
]



# ===================== 公用工具函数 =====================

def normalize_title(title: str) -> str:
    if not title or len(title) < 3:
        return ""
    title = re.sub(r'\s+', ' ', title).strip()
    if len(title) > 30:
        half = len(title) // 2
        for i in range(1, half):
            if title[:i] == title[i:i+i] and len(title) > 2 * i + 5:
                return title[:i].strip()
        third = len(title) // 3
        for i in range(3, third):
            part = title[:i]
            if title[i:i+i] == part and title[2*i:2*i+i] == part:
                return part.strip()
    return title


def parse_date(date_str: str) -> Optional[datetime]:
    if not date_str:
        return None
    try:
        date_str = date_str.strip().replace('/', '-')
        for fmt in ["%Y-%m-%d", "%Y/%m/%d", "%Y%m%d",
                     "%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S"]:
            try:
                return datetime.strptime(date_str[:19], fmt)
            except:
                continue
    except:
        pass
    return None


def extract_date_from_url(url: str) -> str:
    m = re.search(r'/t(20\d{2})(\d{2})(\d{2})_', url)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.search(r'_(20\d{2})(\d{2})(\d{2})[\._]', url)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.search(r'/(20\d{2})-(\d{2})-(\d{2})/', url)
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
    m = re.search(r'/(20\d{2})(\d{2})(\d{2})/', url)
    if m:
        try:
            y, mo, d = m.group(1), m.group(2), m.group(3)
            datetime.strptime(f"{y}-{mo}-{d}", "%Y-%m-%d")
            return f"{y}-{mo}-{d}"
        except:
            pass
    return ""


def extract_date_from_soup(soup, url: str = "") -> str:
    candidates = []
    for tag in ['time', 'span', 'p', 'div', 'li', 'em']:
        for el in soup.find_all(tag):
            text = el.get_text(strip=True)
            m = re.search(r'(20\d{2})[-/](\d{1,2})[-/](\d{1,2})', text)
            if m:
                try:
                    d = f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
                    datetime.strptime(d, "%Y-%m-%d")
                    candidates.append(d)
                except:
                    pass
    if candidates:
        return max(candidates)
    return ""


def extract_keywords_from_info_type(info_type: str) -> Set[str]:
    keywords = set()
    if not info_type:
        return keywords
    clean = re.sub(r'[（(].*?[）)]', '', info_type)
    parts = re.split(r'[、，,;；]', clean)
    for p in parts:
        p = p.strip()
        if len(p) >= 2:
            keywords.add(p)
    return keywords


def get_dept_keywords(dept_name: str) -> Set[str]:
    return DEPT_KEYWORDS.get(dept_name, set())


def is_relevant_for_article(title: str, keywords: Set[str]) -> bool:
    if not keywords:
        return True
    title_lower = title.lower()
    for kw in keywords:
        if kw.lower() in title_lower:
            return True
    return False


# ===================== LLM 过滤函数（v6 新增） =====================

def _call_deepseek(prompt: str, system_prompt: str = "",
                    temperature: float = 0.1, max_tokens: int = 200) -> str:
    """调用 DeepSeek 模型"""
    api_key = os.environ.get("DEEPSEEK_API_KEY", "")
    if not api_key:
        return ""
    try:
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})
        messages.append({"role": "user", "content": prompt})
        payload = {
            "model": "deepseek-chat",
            "messages": messages,
            "temperature": temperature,
            "max_tokens": max_tokens,
        }
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        resp = requests.post("https://api.deepseek.com/v1/chat/completions",
                              json=payload, headers=headers, timeout=15)
        if resp.status_code == 200:
            return resp.json()['choices'][0]['message']['content'].strip()
        return ""
    except Exception:
        return ""


class EnterpriseCrawlerV3:

    def __init__(self, run_date_str: str):
        self.run_date = datetime.strptime(run_date_str, "%Y-%m-%d") if run_date_str else datetime.now()
        self.run_date_str = self.run_date.strftime("%Y-%m-%d")
        self.time_range_days = DEFAULT_TIME_RANGE_DAYS
        self.cutoff_date = self.run_date - timedelta(days=self.time_range_days)
        self.all_articles = {}
        self.failures = []
        self.results_log = []
        self.dept_name = ""
        self.dept_keywords = set()

    def _clean_article_text(self, text: str) -> str:
        """清洗文章摘要，去掉无效信息和类似“该站可能网络问题”的废话"""
        if not text:
            return ""
        # 去除短且无意义的内容
        if len(text) < 15:
            return ""
        # 去除错误提示类文本
        noise_patterns = [
            r'该页面.*无效', r'网络问题', r'请稍后重试',
            r'请检查网络', r'503', r'404', r'502', r'Access Denied',
            r'<script', r'<style', r'您好，欢迎', r' ',
        ]
        for pat in noise_patterns:
            if re.search(pat, text, re.IGNORECASE):
                return ""
        text = re.sub(r'\s+', ' ', text).strip()
        return text[:500]

    def crawl_news_source(self, src: dict, raw_dir: str) -> dict:
        """爬取单个新闻源网站的新闻"""
        name = src.get("name", "")
        url = src.get("url", "")
        src_type = src.get("type", "web")
        info_type = src.get("info_type", "")

        result = {"name": name, "url": url, "type": "web", "count": 0, "articles": []}

        if src_type == "wechat":
            return {"name": name, "url": url, "type": "wechat",
                    "count": 0, "error": "微信公众号内容不可直接爬取"}

        try:
            # print(f"\n  \u250c\u2500 {name} ({url})")
            headers = {"User-Agent": random.choice(USER_AGENTS)}
            resp = requests.get(url, headers=headers, timeout=15, verify=False)
            resp.encoding = resp.apparent_encoding or "utf-8"

            if resp.status_code != 200:
                err = f"HTTP {resp.status_code}"
                self.failures.append({"name": name, "url": url, "error": err})
                result["error"] = err
                return result

            soup = BeautifulSoup(resp.text, "html.parser")
            # 提取文章链接
            articles = self._extract_articles_from_page(soup, url, info_type)

            if not articles:
                # 试用 _agent_browser_fetch 作为备用
                articles = self._crawl_via_agent_browser(src, raw_dir)

            # 保存到文件
            if articles:
                dept_name = self.dept_name
                filepath = os.path.join(raw_dir, f"{dept_name}_news_{name}.txt")
                self._save_articles_to_file(articles, filepath)
                self.all_articles.setdefault(self.dept_name, []).extend(articles)
                result["count"] = len(articles)
                result["articles"] = articles
                # print(f"  \u2514\u2500 \u2705 获取 {len(articles)} 条")
            else:
                print(f"  \u2514\u2500 \u26a0\ufe0f 0条")

        except requests.exceptions.Timeout:
            err = "连接超时"
            self.failures.append({"name": name, "url": url, "error": err})
            result["error"] = err
        except requests.exceptions.SSLError as e:
            err = f"SSL错误: {str(e)[:80]}"
            self.failures.append({"name": name, "url": url, "error": err})
            result["error"] = err
        except Exception as e:
            err = str(e)[:120]
            self.failures.append({"name": name, "url": url, "error": err})
            result["error"] = err

        self.results_log.append(result)
        return result

    def _extract_articles_from_page(self, soup, base_url: str, info_type: str) -> List[dict]:
        """从 HTML 中提取文章列表"""
        articles = []
        seen_links = set()

        # 尝试多种 CSS 选择器
        selectors = [
            'a[href*=".html"]', 'a[href*=".shtml"]', 'a[href*="article"]',
            'a[href*="news"]', 'a[href*="content"]', 'a[href*="detail"]',
            'a[href*="/"]',
        ]

        for sel in selectors:
            for a_tag in soup.select(sel):
                href = a_tag.get("href", "").strip()
                title = a_tag.get_text(strip=True)
                if not title or len(title) < 6:
                    continue
                title = normalize_title(title)
                if not title:
                    continue

                # 完整 URL
                if href.startswith("/"):
                    parsed = urlparse(base_url)
                    full_url = f"{parsed.scheme}://{parsed.netloc}{href}"
                elif href.startswith("http"):
                    full_url = href
                else:
                    full_url = urljoin(base_url, href)

                if full_url in seen_links:
                    continue
                seen_links.add(full_url)

                # 主题匹配
                kw = extract_keywords_from_info_type(info_type)
                kw.update(self.dept_keywords)
                if kw and not is_relevant_for_article(title, kw):
                    continue

                date_str = extract_date_from_url(full_url)
                if not date_str:
                    nearby = soup.find_all(["span", "p", "li", "em", "time"])
                    for el in nearby:
                        d = re.search(r'(20\d{2})[-/](\d{1,2})[-/](\d{1,2})', el.get_text())
                        if d:
                            try:
                                date_str = f"{d.group(1)}-{int(d.group(2)):02d}-{int(d.group(3)):02d}"
                                if self.cutoff_date <= datetime.strptime(date_str, "%Y-%m-%d") <= self.run_date:
                                    break
                            except:
                                pass

                # 时间范围过滤
                if date_str:
                    try:
                        article_date = datetime.strptime(date_str, "%Y-%m-%d")
                        if article_date < self.cutoff_date or article_date > self.run_date:
                            continue
                    except:
                        pass

                articles.append({
                    "title": title,
                    "link": full_url,
                    "date": date_str,
                    "source": self.dept_name,
                    "summary": "",
                    "info_type": info_type,
                })

        # 按时间排序，取前20条
        articles.sort(key=lambda x: x["date"], reverse=True)
        return articles[:20]

    def _crawl_via_agent_browser(self, src: dict, raw_dir: str) -> List[dict]:
        """备用：通过 _agent_browser_fetch 处理动态页面"""
        try:
            sys.path.insert(0, os.path.dirname(__file__))
            from _agent_browser_fetch import fetch_page_list
            articles = fetch_page_list(src.get("url", ""), src.get("name", ""),
                                        scroll_times=2, max_items=30)
            result = []
            for a in articles:
                title = normalize_title(a.get("title", ""))
                if not title:
                    continue
                kw = extract_keywords_from_info_type(src.get("info_type", ""))
                kw.update(self.dept_keywords)
                if kw and not is_relevant_for_article(title, kw):
                    continue
                result.append({
                    "title": title,
                    "link": a.get("link", ""),
                    "date": a.get("date", "") or extract_date_from_url(a.get("link", "")),
                    "source": src["name"],
                    "summary": a.get("summary", ""),
                    "info_type": src.get("info_type", ""),
                })
            # print(f"  \u2514\u2500 \u2705 agent_browser 获取 {len(result)} 条")
            return result[:20]
        except Exception as e:
            print(f"  \u2514\u2500 \u26a0\ufe0f agent_browser 失败: {e}")
            return []

    def _save_articles_to_file(self, articles: List[dict], filepath: str):
        """\u4fdd\u5b58\u6587\u7ae0\u5230\u6587\u4ef6"""
        with open(filepath, "w", encoding="utf-8") as f:
            for i, a in enumerate(articles):
                if i > 0:
                    f.write("---\n")
                f.write("\u6807\u9898: {}\n".format(a.get('title', '')))
                f.write("\u65e5\u671f: {}\n".format(a.get('date', '')))
                f.write("\u6765\u6e90: {}\n".format(a.get('source', a.get('source_name', ''))))
                f.write("\u94fe\u63a5: {}\n".format(a.get('link', '')))
                # 清洗摘要（去掉无效信息）
                raw_summary = a.get('summary', a.get('abstract', ''))
                cleaned_summary = self._clean_article_text(raw_summary)
                f.write("\u6458\u8981: {}\n".format(cleaned_summary))
                if a.get('info_type'):
                    f.write("\u4fe1\u606f\u7c7b\u578b: {}\n".format(a['info_type']))

    def run_dept(self, config: dict, raw_dir: str,
                 do_news: bool = True) -> dict:
        """\u8fd0\u884c\u6574\u4e2a\u5904\u5ba4\u7684\u6240\u6709\u91c7\u96c6 \u2014 v5\u7248"""
        raw_dept = config["department"]
        dept_name = raw_dept.replace('\u3001', '_')
        self.dept_name = dept_name
        self.dept_keywords = get_dept_keywords(dept_name)
        self.all_articles.setdefault(dept_name, [])

        self.time_range_days = config.get("time_range_days", DEFAULT_TIME_RANGE_DAYS)
        self.cutoff_date = self.run_date - timedelta(days=self.time_range_days)

        print(f"\n{'='*60}")
        print(f"  {raw_dept}")
        print(f"{'='*60}")

        # === 新闻采集 ===
        if do_news:
            print(f"\n\u2500\u2500\u2500 \u65b0\u95fb\u91c7\u96c6 ({len(config.get('news_sources', []))}\u4e2a\u6765\u6e90) \u2500\u2500\u2500")
            for src in config.get("news_sources", []):
                result = self.crawl_news_source(src, raw_dir)
                self.results_log.append(result)

        # --- 舆情采集已移至 step_sentiment_v3.py 独立运行 ---

        # --- 百度补充搜索已移除，移至 step_sentiment_v3.py ---

        # === \u4fdd\u5b58\u5168\u91cf\u6587\u7ae0\u5230JSON ===
        articles_file = os.path.join(raw_dir, f"{dept_name}_all_articles.json")
        with open(articles_file, "w", encoding="utf-8") as f:
            json.dump(self.all_articles.get(dept_name, []), f,
                      ensure_ascii=False, indent=2)

        print(f"\n  \u2705 \u6587\u7ae0\u5df2\u4fdd\u5b58: {articles_file}")
        print(f"  \u2705 \u5168\u91cf: {len(self.all_articles.get(dept_name, []))}\u6761")
        print(f"  \u2705 \u5931\u8d25\u8bb0\u5f55: {len(self.failures)}\u6761")

        # === \u6253\u5370\u8be6\u7ec6\u91c7\u96c6\u62a5\u544a ===
        self._print_dept_report(raw_dept, config)

        return {"dept": dept_name, "articles": self.all_articles.get(dept_name, []),
                "failures": self.failures}

    def _print_dept_report(self, raw_dept: str, config: dict):
        """\u6253\u5370\u5904\u5ba4\u7ea7\u522b\u91c7\u96c6\u62a5\u544a\uff1a\u6bcf\u4e2a\u7f51\u5740\u7684\u53ef\u722c\u72b6\u6001\u3001\u83b7\u53d6\u6761\u6570\u3001\u5931\u8d25\u539f\u56e0"""
        # print()
        # print(f"{'='*60}")
        # print(f"  [\u91c7\u96c6\u62a5\u544a] {raw_dept}")
        # print(f"{'='*60}")

        # \u65b0\u95fb\u6e90\u72b6\u6001
        sources = config.get("news_sources", [])
        if sources:
            # print(f"\n  \u250c\u2500 \u7f51\u7ad9\u65b0\u95fb\u6765\u6e90 ({len(sources)}\u4e2a)")
            ok_count = 0
            fail_count = 0
            for src in sources:
                name = src.get("name", "")
                url = src.get("url", "")
                # \u4ece results_log \u4e2d\u67e5\u627e\u5bf9\u5e94\u7ed3\u679c
                result = None
                for r in self.results_log:
                    if r.get("name") == name and r.get("type", "web") != "sentiment":
                        result = r
                        break
                if result:
                    count = result.get("count", 0)
                    if count > 0:
                        ok_count += 1
                        # print(f"  \u251c\u2500 \u2705 {name}\t| url={url}\t| \u83b7\u53d6 {count}\u6761")
                    elif result.get("error", ""):
                        fail_count += 1
                        # print(f"  \u251c\u2500 \u274c {name}\t| url={url}\t| \u5931\u8d25: {result['error']}")
                    else:
                        # print(f"  \u251c\u2500 \u26a0\ufe0f {name}\t| url={url}\t| 0\u6761\uff08\u65e0\u5339\u914d\u5185\u5bb9\uff09")
                        pass
                else:
                    # \u6ca1\u6709\u8f93\u5165\u7ed3\u679c\uff08\u5982\u5fae\u4fe1\u516c\u4f17\u53f7\uff09
                    fail_count += 1
                    # print(f"  \u251c\u2500 \u274c {name}\t| url={url}\t| \u65e0URL\u914d\u7f6e\u6216\u4e0d\u53ef\u76f4\u63a5\u722c\u53d6")
            # print(f"  \u2514\u2500 \u5408\u8ba1: \u2705 {ok_count}\u4e2a\u6210\u529f | \u274c {fail_count}\u4e2a\u5931\u8d25 | \u603b\u83b7\u53d6 {len(self.all_articles.get(self.dept_name, []))}\u6761")

        # \u5fae\u4fe1\u516c\u4f17\u53f7\u6765\u6e90
        wechat_sources = [s for s in sources if s.get("type") == "wechat"]
        if wechat_sources:
            print(f"\n  \u250c\u2500 \u5fae\u4fe1\u516c\u4f17\u53f7\u6765\u6e90 ({len(wechat_sources)}\u4e2a)")
            for ws in wechat_sources:
                print(f"  \u251c\u2500 \u274c {ws.get('name', '')}\t| \u5fae\u4fe1\u516c\u4f17\u53f7\u5185\u5bb9\u4e0d\u53ef\u76f4\u63a5\u722c\u53d6")
            print(f"  \u2514\u2500 \u5fae\u4fe1\u6765\u6e90\u5747\u4e3a\u5931\u8d25\uff0c\u9700\u4e13\u7528\u5de5\u5177")

        # \u8206\u60c5\u91c7\u96c6\u72b6\u6001
        companies = config.get("sentiment_companies", [])
        if companies:
            print(f"\n  \u250c\u2500 \u8206\u60c5\u91c7\u96c6 ({len(companies)}\u5bb6\u4f01\u4e1a)")
            for r in self.results_log:
                if r.get("type") == "sentiment":
                    stats = r.get("stats", {})
                    print(f"  \u251c\u2500 \u641c\u7d22\u6210\u529f: {stats.get('ok', 0)} \u5bb6 | \u5931\u8d25: {stats.get('fail', 0)} \u5bb6 | \u83b7\u53d6\u6761\u6570: {stats.get('articles', 0)}")
                    break

        # \u5931\u8d25\u539f\u56e0\u6c47\u603b
        dept_failures = [f for f in self.failures if f.get("name", "") in [s.get("name","") for s in sources]]
        if dept_failures:
            print(f"\n  \u250c\u2500 \u5931\u8d25\u539f\u56e0\u8be6\u60c5")
            for ff in dept_failures:
                err = ff.get("error", "")
                name = ff.get("name", "")
                print(f"  \u251c\u2500 {name}: {err}")

        print(f"{'='*60}\n")


# ===================== \u751f\u6210\u91c7\u96c6\u8fd0\u884c\u62a5\u544a =====================

def _generate_run_report(run_date_str: str, raw_dir: str, args, crawler):
    """生成 采集运行报告.md，记录每个处室每个网站的可爬状态"""
    import os
    import glob

    archive_dir = os.path.join(os.path.dirname(raw_dir), "archive")
    os.makedirs(archive_dir, exist_ok=True)

    report_lines = []
    report_lines.append("# 采集运行报告")
    report_lines.append("")
    report_lines.append(f"**运行日期：** {run_date_str}")
    report_lines.append(f"**采集范围：** {'新闻' if not args.no_news else ''}{'+舆情' if not args.no_sentiment else ''}")
    report_lines.append("")
    report_lines.append("---")
    report_lines.append("")

    # 自动检测 raw/ 下已经运行过的处室，而不再依赖 args.departments
    raw_json_files = glob.glob(os.path.join(raw_dir, "*_all_articles.json"))
    completed_depts = set()
    for fpath in raw_json_files:
        basename = os.path.basename(fpath)
        dept_name = basename.replace("_all_articles.json", "")
        completed_depts.add(dept_name)

    # 如果 raw 下已有数据，则用已完成的处室；否则回退到 args.departments
    dept_list = sorted(list(completed_depts)) if completed_depts else args.departments

    for dept_name in dept_list:
        config_path = os.path.join(CONFIG_DIR, f"{dept_name}.json")
        if not os.path.exists(config_path):
            report_lines.append(f"## {dept_name}")
            report_lines.append("")
            report_lines.append("> ⚠️ 配置文件不存在，无法生成详细报告")
            report_lines.append("")
            report_lines.append("---")
            report_lines.append("")
            continue

        with open(config_path, encoding="utf-8") as f:
            config = json.load(f)

        raw_dept = config.get("department", dept_name)
        report_lines.append(f"## {raw_dept}")
        report_lines.append("")

        # 新闻源状态
        sources = config.get("news_sources", [])
        if sources:
            report_lines.append("### 网站新闻来源")
            report_lines.append("")
            report_lines.append("| 网站 | URL | 状态 | 获取条数 | 说明 |")
            report_lines.append("|------|-----|--------|----------|------|")

            ok_count = 0
            fail_count = 0
            total_count = 0

            for src in sources:
                name = src.get("name", "")
                url = src.get("url", "")
                src_type = src.get("type", "web")

                if src_type == "wechat":
                    report_lines.append(f"| {name} | 微信公众号 | ❌ 失败 | 0 | 微信内容不可直接爬取 |")
                    fail_count += 1
                    continue

                # 从 raw/ 目录下的 .txt 文件推断结果（不依赖 results_log，支持跨运行合并）
                dept_file_prefix = dept_name.replace('、', '_') + '_news_' + name
                found_files = [f for f in os.listdir(raw_dir) if f.startswith(dept_file_prefix)]
                if found_files:
                    txt_path = os.path.join(raw_dir, found_files[0])
                    with open(txt_path, 'r', encoding='utf-8') as _f:
                        txt_content = _f.read()
                    article_count = txt_content.count('\n---\n') + 1 if '---' in txt_content else (1 if txt_content.strip() else 0)
                    total_count += article_count
                    status = "✅ 成功"
                    note = f"获取 {article_count}条"
                    ok_count += 1
                else:
                    err_msg = ""
                    for ff in crawler.failures:
                        if ff.get("name") == name:
                            err_msg = ff.get("error", "")
                            break
                    if err_msg:
                        status = "❌ 失败"
                        note = err_msg
                        fail_count += 1
                    elif src_type == "wechat":
                        status = "❌ 失败"
                        note = "微信内容不可直接爬取"
                        fail_count += 1
                    else:
                        status = "⚠️ 无匹配"
                        note = "未提取到匹配内容"

                report_lines.append(f"| {name} | {url} | {status} | {article_count if found_files else 0} | {note} |")

            report_lines.append("")
            report_lines.append(f"**合计：** ✅ {ok_count}个成功 | ❌ {fail_count}个失败 | 共获取 {total_count}条")
            report_lines.append("")

        # 微信公众号
        wechat_sources = [s for s in sources if s.get("type") == "wechat"]
        if wechat_sources:
            report_lines.append("### 微信公众号来源")
            report_lines.append("")
            for ws in wechat_sources:
                report_lines.append(f"- ❌ {ws.get('name', '')}: 微信内容不可直接爬取")
            report_lines.append("")

        # 舆情采集
        companies = config.get("sentiment_companies", [])
        if companies and not args.no_sentiment:
            report_lines.append("### 舆情采集")
            report_lines.append("")
            for r in crawler.results_log:
                if r.get("type") == "sentiment":
                    stats = r.get("stats", {})
                    report_lines.append(f"- 搜索成功: {stats.get('ok', 0)} 家")
                    report_lines.append(f"- 搜索失败: {stats.get('fail', 0)} 家")
                    report_lines.append(f"- 获取舆情条数: {stats.get('articles', 0)} 条")
                    break
            report_lines.append("")

        report_lines.append("---")
        report_lines.append("")

    # 失败原因汇总
    if crawler.failures:
        report_lines.append("## 失败原因汇总与排查建议")
        report_lines.append("")
        report_lines.append("| 名称 | 类型 | 错误原因 | 排查建议 |")
        report_lines.append("|------|------|---------|---------|")
        for ff in crawler.failures:
            name = ff.get("name", ff.get("company", ""))
            err = ff.get("error", "")
            # 类型判断
            ftype = "网站" if "http" in err.lower() or "连接" in err else "舆情搜索"
            # 简单建议
            suggestion = err
            if "连接超时" in err or "Timeout" in err:
                suggestion = "服务器响应慢或防火墙拦截，建议：尝试简化URL、增加超时时间"
            elif "NameResolutionError" in err or "Failed to resolve" in err:
                suggestion = "域名解析失败，检查域名拼写或DNS配置"
            elif "SSL" in err or "certificate" in err:
                suggestion = "SSL证书错误，尝试用 http:// 或短时间内跳过"
            elif "412" in err:
                suggestion = "HTTP 412预检失败，网站可能有反爬机制"
            elif "微信" in err:
                suggestion = "微信公众号内容不可直接爬取，需专用微信生态采集工具"

            report_lines.append(f"| {name} | {ftype} | {err} | {suggestion} |")

        report_lines.append("")

    report_lines.append(f"---")
    report_lines.append(f"*报告生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*")

    report_path = os.path.join(archive_dir, "采集运行报告.md")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(report_lines))
    print(f"\n📋 采集报告已生成: {report_path}")

# ===================== Main =====================
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description='\u4f01\u4e1a\u8206\u60c5\u4e0e\u65b0\u95fb\u91c7\u96c6\u7cfb\u7edf v6')
    parser.add_argument('--departments', nargs='+', default=DEPARTMENTS,
                        help='\u8981\u8fd0\u884c\u7684\u5904\u5ba4\u540d\u79f0 (\u9ed8\u8ba4: \u5168\u90e8)')
    parser.add_argument('--date', type=str, default='',
                        help='\u8fd0\u884c\u65e5\u671f (\u9ed8\u8ba4\u4eca\u5929, \u683c\u5f0f: YYYY-MM-DD)')
    parser.add_argument('--raw-dir', type=str, default='',
                        help='\u539f\u59cb\u6570\u636e\u76ee\u5f55 (\u9ed8\u8ba4: data/\u65e5\u671f/raw)')
    parser.add_argument('--no-news', action='store_true',
                        help='\u8df3\u8fc7\u65b0\u95fb\u91c7\u96c6')
    parser.add_argument('--no-sentiment', action='store_true',
                        help='\u8df3\u8fc7\u8206\u60c5\u91c7\u96c6')
    parser.add_argument('--max-companies', type=int, default=0,
                        help='\u6700\u5927\u4f01\u4e1a\u6570 (0=\u5168\u90e8)')
    args = parser.parse_args()

    run_date_str = args.date or datetime.now().strftime("%Y-%m-%d")
    raw_dir = args.raw_dir or os.path.join(DATA_DIR, run_date_str, "raw")
    os.makedirs(raw_dir, exist_ok=True)

    crawler = EnterpriseCrawlerV3(run_date_str=run_date_str)
    for dept_name in args.departments:
        config_path = os.path.join(CONFIG_DIR, f"{dept_name}.json")
        if not os.path.exists(config_path):
            print(f"\u26a0\ufe0f \u914d\u7f6e\u6587\u4ef6\u4e0d\u5b58\u5728: {config_path}")
            continue
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        crawler.run_dept(config, raw_dir,
                         do_news=not args.no_news)
    print(f"\n{'='*60}")
    print(f"  \u2714\ufe0f \u91c7\u96c6\u5b8c\u6210 | \u5931\u8d25: {len(crawler.failures)}\u6761")
    if crawler.failures:
        for f in crawler.failures[:10]:
            print(f"    - {f.get('name', f.get('company',''))}: {f.get('error','')}")
    print(f"{'='*60}")

    # === \u751f\u6210 archive/\u91c7\u96c6\u8fd0\u884c\u62a5\u544a.md ===
    _generate_run_report(run_date_str, raw_dir, args, crawler)
