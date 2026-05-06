#!/usr/bin/env python3
"""
step_sentiment_v3.py - 舆情采集专用脚本

从 enterprise_crawler_v3.py 分离出的舆情搜索模块。
使用百度搜索 API + LLM 过滤，采集配置文件中指定公司的近一周舆情信息。

用法:
  python3 step_sentiment_v3.py --date 2026-04-30
  python3 step_sentiment_v3.py --date 2026-04-30 --departments 医药处
  python3 step_sentiment_v3.py --date 2026-04-30 --max-companies 10
"""

import sys
import argparse
import json
import os
import re
import time
import random
import subprocess
from datetime import datetime, timedelta
from typing import List, Dict, Tuple, Optional

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_DIR = os.path.join(BASE_DIR, "config")
DATA_DIR = os.path.join(BASE_DIR, "data")

BAIDU_API_KEY = os.environ.get("BAIDU_API_KEY", "")
BAIDU_SEARCH_SCRIPT = os.path.expanduser(
    "~/.openclaw/workspace/skills/baidu-search/scripts/search.py"
)

# 从 runtime_config.json 读取 API key（可选）
RUNTIME_CONFIG = os.path.join(BASE_DIR, "runtime_config.json")
if os.path.exists(RUNTIME_CONFIG):
    with open(RUNTIME_CONFIG, encoding="utf-8") as f:
        rc = json.load(f)
        if not BAIDU_API_KEY:
            BAIDU_API_KEY = rc.get("BAIDU_API_KEY", "")
        os.environ.setdefault("DEEPSEEK_API_KEY", rc.get("DEEPSEEK_API_KEY", ""))

if not BAIDU_API_KEY:
    BAIDU_API_KEY = os.environ.get("BAIDU_API_KEY",
        "bce-v3/ALTAK-bBCOizZ2VaB1dVON4z1d5/3b5c35c46c163050cb4b0b08cefa8282a2099c80")

GENERAL_NEGATIVE_KEYWORDS = ["处罚", "违规", "事故", "罚款", "停产", "整改", "诉讼",
                              "查封", "投诉", "安全问题", "环保", "负面", "风险"]

DEFAULT_TIME_RANGE_DAYS = 7
DEPARTMENTS = ["工业互联网处", "医药处", "原材料_金属_非金属处"]


# ===================== 工具函数 =====================

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
                pass
        return None
    except:
        return None


def extract_date_from_url(url: str) -> str:
    if not url:
        return ""
    patterns = [
        r'/(20\d{2}[-/]\d{1,2}[-/]\d{1,2})',
        r'(20\d{2}[-/]\d{1,2}[-/]\d{1,2})',
        r'/(20\d{2})(\d{2})(\d{2})',
    ]
    for pat in patterns:
        m = re.search(pat, url)
        if m:
            try:
                d = m.group(1).replace('/', '-')
                if re.match(r'^\d{4}-\d{2}-\d{2}$', d):
                    return d
                d2 = f"{m.group(1)}-{m.group(2)}-{m.group(3)}"
                if re.match(r'^\d{4}-\d{2}-\d{2}$', d2):
                    return d2
            except:
                pass
    return ""


def _call_deepseek(prompt: str, system_prompt: str = "",
                    temperature: float = 0.1, max_tokens: int = 200) -> str:
    api_key = os.environ.get("DEEPSEEK_API_KEY", "")
    if not api_key:
        return ""
    try:
        import requests
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


# ===================== 舆情搜索相关函数 =====================

def build_search_queries_for_company(company: str, info_type: str, remark: str,
                                      dept_name: str = "") -> List[str]:
    """
    为单个公司构建舆情搜索 query（极简版：每家公司只搜2条，避免百度429限流）
    舆情采集只关注：公司本身的新闻 + 1条行业相关
    """
    queries = []
    if not company:
        return queries

    # 1. 基础搜索：公司名（最重要的一条）
    queries.append(company)

    # 2. 从 info_type 或 remark 选一个最有代表性的行业词
    best_kw = ""
    if info_type:
        for tk in re.split(r'[、，,;；]', info_type):
            tk = tk.strip()
            if len(tk) >= len(best_kw):
                best_kw = tk
    if remark and len(remark) > len(best_kw):
        for rk in re.split(r'[、，,;；]', remark):
            rk = rk.strip()
            if len(rk) >= len(best_kw):
                best_kw = rk
    if best_kw:
        queries.append(f"{company} {best_kw}")

    return queries[:2]  # 最多2条


def _fallback_filter(articles: List[dict], company: str) -> List[dict]:
    """无LLM时的关键词回退过滤"""
    if not articles or not company:
        return articles
    company_parts = set(re.split(r'[（(）)\s]', company))
    company_parts = {p for p in company_parts if len(p) >= 2}
    keep = []
    for a in articles:
        title = a.get('title', '')
        abstract = (a.get('abstract', '') or '')[:300]
        combined = title + ' ' + abstract
        if company in combined:
            keep.append(a)
            continue
        matched_parts = sum(1 for cp in company_parts if cp in title)
        if matched_parts >= 1:
            keep.append(a)
            continue
        event_indicators = ["签约", "投产", "开工", "投产", "融资", "上市", "投资",
                            "收购", "并购", "合作", "中标", "获奖", "认定", "公示",
                            "招标", "项目", "订单", "出口", "获批", "新药", "批准",
                            "注册", "许可", "牌照"]
        has_event = any(e in title for e in event_indicators)
        has_negative = any(kw in title for kw in GENERAL_NEGATIVE_KEYWORDS)
        if has_event and any(cp in combined for cp in company_parts):
            keep.append(a)
            continue
        if has_negative and any(cp in combined for cp in company_parts):
            keep.append(a)
            continue
    return keep




def _fetch_article_content(url: str, max_chars: int = 1000) -> str:
    """获取文章正文内容，用于 LLM 相关性判断"""
    try:
        import urllib.request
        req = urllib.request.Request(url, headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
            "Accept": "text/html,application/xhtml+xml",
            "Accept-Language": "zh-CN,zh;q=0.9",
        })
        with urllib.request.urlopen(req, timeout=10) as resp:
            html = resp.read().decode("utf-8", errors="replace")
        # 提取正文文本
        text = re.sub(r'<script[^>]*>.*?</script>', '', html, flags=re.DOTALL)
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL)
        text = re.sub(r'<[^>]+>', ' ', text)
        text = re.sub(r'\s+', ' ', text).strip()
        return text[:max_chars]
    except Exception:
        return ""


def analyze_baidu_results(articles: List[dict], company: str, info_type: str) -> List[dict]:
    """
    对搜索结果进行智能过滤（舆情优化版）
    先基于标题快速过滤，对 uncertain 的获取正文后再用 LLM 判断
    """
    if not articles:
        return []

    # 提取核心公司名
    core_name = company
    for suffix in ["有限公司", "集团有限公司", "有限责任公司", "股份有限公司", "集团", "科技公司"]:
        if suffix in company:
            core_name = company.split(suffix)[0]
            break

    keep = []
    uncertain = []
    for a in articles:
        title = a.get('title', '')
        abstract = (a.get('abstract', '') or '')[:300]
        combined = title + ' ' + abstract

        # 核心名称出现在标题或摘要 → 直接保留
        if core_name in combined or core_name in title:
            keep.append(a)
            continue

        # 通用行业关键词 → uncertain（待正文验证）
        general_kw = ["医药", "药", "生物", "临床", "新药", "医疗", "研发",
                      "芯片", "半导体", "电子", "信息", "数据", "AI", "人工智能",
                      "矿", "金属", "材料", "能源", "电池", "化工",
                      "汽车", "制造", "科技", "数字", "智能"]
        if any(kw in title for kw in general_kw):
            uncertain.append(a)
            continue

        # 事件/动态类 → uncertain
        event_indicators = ["签约", "投产", "开工", "融资", "上市", "投资",
                            "收购", "并购", "合作", "中标", "获奖", "认定",
                            "招标", "项目", "订单", "出口", "获批",
                            "批准", "注册", "许可", "牌照", "发布", "宣布"]
        if any(e in title for e in event_indicators):
            uncertain.append(a)
            continue

        # 负面关键词 → 直接保留（负面舆情不放过）
        negative_kw = ["负面", "风险", "事故", "处罚", "诉讼", "违规",
                       "停产", "整改", "罚款", "调查", "安全", "环保",
                       "召回", "追责", "曝光"]
        if any(kw in title for kw in negative_kw):
            keep.append(a)
            continue

    # 对 uncertain 的文章：获取正文后 LLM 判断
    if uncertain:
        if os.environ.get("DEEPSEEK_API_KEY") and len(uncertain) <= 3:
            # 获取正文
            contents = []
            for a in uncertain:
                link = a.get('link', '')
                body = ""
                if link and (link.startswith("http://") or link.startswith("https://")):
                    # print(f"        ↪ 获取正文: {link[:50]}...")
                    body = _fetch_article_content(link, max_chars=800)
                contents.append(body or a.get('abstract', '')[:300])

            llm_items = "\n\n".join(
                f"--- 文章{i+1} ---\n标题: {a.get('title', '')[:80]}\n正文(节选): {contents[i][:500]}"
                for i, a in enumerate(uncertain)
            )
            prompt = (
                f"公司名称：{company}\n"
                f"需要关注的资讯类型：{info_type}\n\n"
                f"以下是从搜索结果获取的文章标题和正文节选。\n"
                f"请判断每条文章是否与该公司直接相关（提及公司名、产品、项目、行业动态等）。\n"
                f"输出相关的序号（逗号分隔）。若全部无关，只输出 无。\n\n"
                f"{llm_items}"
            )
            result = _call_deepseek(prompt,
                                    system_prompt='你是一个专业的信息研判助手，根据文章内容准确判断相关性。只输出序号或"无"。',
                                    temperature=0.1, max_tokens=100)
            if result and result.strip() and result.strip() != "无":
                indices = re.findall(r'\d+', result)
                for idx_str in indices:
                    try:
                        idx = int(idx_str) - 1
                        if 0 <= idx < len(uncertain):
                            keep.append(uncertain[idx])
                    except:
                        pass
            else:
                # LLM 认为无关，但保留1条防遗漏
                if not keep and uncertain:
                    keep.append(uncertain[0])
        else:
            # 无 LLM key 或 uncertain 太多：保险保留1条
            if not keep and uncertain:
                keep.append(uncertain[0])

    return keep

# ===================== 舆情采集类 =====================

class SentimentCrawler:
    """舆情采集器：百度搜索 + LLM 过滤"""

    def __init__(self, run_date_str: str = "", dept_name: str = ""):
        self.run_date_str = run_date_str or datetime.now().strftime("%Y-%m-%d")
        self.run_date = self._parse_date(self.run_date_str)
        self.dept_name = dept_name
        self.failures: List[dict] = []
        self.results_log: List[dict] = []
        self._baidu_429_count = 0
        self._baidu_429_use_tavily = False

    def _parse_date(self, date_str: str) -> Optional[datetime]:
        return parse_date(date_str) or datetime.now().replace(
            hour=0, minute=0, second=0, microsecond=0)

    def baidu_search(self, query: str, count: int = 10, freshness: str = "pw",
                     max_retries: int = 3) -> List[dict]:
        """调用百度搜索API"""
        results = []
        for attempt in range(max_retries):
            try:
                time.sleep(random.uniform(1.5, 3.0))
                args = {"query": query, "count": count, "freshness": freshness}
                args_json = json.dumps(args, ensure_ascii=False)
                env = os.environ.copy()
                if BAIDU_API_KEY:
                    env["BAIDU_API_KEY"] = BAIDU_API_KEY

                cmd = ["python3", BAIDU_SEARCH_SCRIPT, args_json]
                proc = subprocess.run(cmd, capture_output=True, text=True,
                                      timeout=30, env=env)

                if proc.returncode != 0:
                    err = proc.stderr.strip().lower()
                    if '429' in err or 'rate' in err or 'limit' in err:
                        self._baidu_429_count += 1
                        if self._baidu_429_count >= 2:
                            self._baidu_429_use_tavily = True
                            print(f"    ⚠️ 百度连续429，降级至Tavily搜索")
                            return results
                        print(f"    ⚠️ 429限频，降级至Tavily")
                        self._baidu_429_use_tavily = True
                        return results
                    return results

                out = proc.stdout.strip()
                # 找到JSON数组的起始位置（去掉 "success parse" 等非JSON前缀）
                json_start = out.find('[')
                if json_start >= 0:
                    data = json.loads(out[json_start:])
                    for item in data:
                        title = normalize_title(item.get('title', '').strip())
                        if len(title) <= 3:
                            continue
                        link = item.get('url', item.get('link', '')).strip()
                        source = item.get('website', item.get('source', ''))
                        # 百度返回的 content 是完整正文，不再截断
                        abstract = item.get('content', item.get('abstract', '')).strip()

                        date = ""
                        for key in ['date', 'publish_time']:
                            val = item.get(key, '')
                            if val:
                                d = str(val)[:10].replace('/', '-')
                                if re.match(r'^\d{4}-\d{2}-\d{2}$', d):
                                    date = d
                                    break
                        if not date:
                            date = extract_date_from_url(link)

                        results.append({
                            "title": title, "date": date,
                            "source_name": source or "百度搜索",
                            "link": link, "abstract": abstract[:1000],
                        })
                    return results
                return results

            except Exception as e:
                e_str = str(e).lower()
                if '429' in e_str or 'rate' in e_str:
                    self._baidu_429_count += 1
                    if self._baidu_429_count >= 2:
                        self._baidu_429_use_tavily = True
                        print(f"    ⚠️ 百度连续429，降级至Tavily搜索")
                        return results
                    self._baidu_429_use_tavily = True
                    print(f"    ⚠️ 429限频，降级至Tavily")
                    return results
                    time.sleep(wait)
                    continue
                print(f"    ❌ 百度搜索失败: {e}")
                return results
        return results

    def tavily_search(self, query: str, count: int = 5, days: int = 7) -> List[dict]:
        """调用 Tavily Search API 搜索舆情（优先使用）"""
        results = []
        try:
            import subprocess
            tavily_script = os.path.expanduser(
                "~/.openclaw/workspace/skills/tavily-search/scripts/search.mjs")

            if not os.path.exists(tavily_script):
                print(f"    ⚠️ Tavily 脚本不存在: {tavily_script}")
                return results

            # 从 runtime_config 获取 key
            tavily_key = ""
            rc_path = os.path.join(BASE_DIR, "runtime_config.json")
            if os.path.exists(rc_path):
                with open(rc_path) as f:
                    rc = json.load(f)
                tavily_key = rc.get("TAVILY_API_KEY", "")

            env = os.environ.copy()
            if tavily_key:
                env["TAVILY_API_KEY"] = tavily_key
            elif "TAVILY_API_KEY" not in env:
                print(f"    ⚠️ TAVILY_API_KEY 未配置")
                return results

            time.sleep(random.uniform(1.0, 2.0))
            cmd = ["node", tavily_script, query,
                   "-n", str(count), "--topic", "news", "--days", str(days)]
            proc = subprocess.run(cmd, capture_output=True, text=True,
                                  timeout=30, env=env)

            if proc.returncode != 0:
                err = proc.stderr.strip()[:200] if proc.stderr else "unknown"
                print(f"    ⚠️ Tavily 搜索失败: {err}")
                return results

            out = proc.stdout.strip()
            # Tavily 输出格式: ## Answer ... ## Sources - **Title** content
            sections = out.split("## Sources")
            if len(sections) < 2:
                return results

            sources_text = sections[1]
            import re as re_module
            source_blocks = re_module.split(r'\n- ', sources_text)
            for block in source_blocks:
                lines = block.strip().split('\n')
                if not lines:
                    continue
                title_line = lines[0].strip().lstrip('- *')
                # 还原可能的加粗标记
                title = title_line.replace('**', '').strip()
                # 去掉 relevance 后缀
                idx_rel = title.rfind('(relevance:')
                if idx_rel > 0:
                    title = title[:idx_rel].strip()

                url_line = ""
                content_line = ""
                for l in lines[1:]:
                    l = l.strip()
                    if l.startswith('http://') or l.startswith('https://'):
                        url_line = l
                    elif l.startswith('|'):
                        content_line = l.lstrip('| ')

                if not title or not url_line:
                    continue
                if len(title) <= 3:
                    continue

                results.append({
                    "title": title,
                    "date": "",
                    "source_name": "Tavily搜索",
                    "link": url_line,
                    "abstract": content_line[:500] if content_line else title[:200],
                })

        except subprocess.TimeoutExpired:
            print(f"    ⚠️ Tavily 搜索超时")
        except Exception as e:
            print(f"    ⚠️ Tavily 异常: {e}")

        return results


    def crawl_dept(self, config: dict, raw_dir: str,
                   max_companies: int = 0) -> dict:
        """采集单个处室的全部舆情"""
        raw_dept = config.get("department", "")
        dept_name = raw_dept.replace('、', '_')
        self.dept_name = dept_name

        companies = config.get("sentiment_companies", [])
        if not companies:
            print(f"  ⚠️ {raw_dept}: 无舆情配置")
            return {"dept": dept_name, "articles": [], "failures": []}

        articles = []
        stats = {"total": len(companies), "ok": 0, "fail": 0, "articles": 0}
        batch_size = 3

        if max_companies > 0:
            companies = companies[:max_companies]

        # print(f"\n{'='*60}")
        # print(f"  [舆情采集] {raw_dept} ({len(companies)}家企业)")
        # print(f"{'='*60}")

        # 为每家公司构建搜索词
        company_queries = {}
        for company in companies:
            cname = company["company"]
            info_type = company.get("info_type", "")
            remark = company.get("remark", "")
            queries = build_search_queries_for_company(cname, info_type, remark, dept_name)
            company_queries[cname] = queries

        for i in range(0, len(companies), batch_size):
            batch = companies[i:i+batch_size]
            total_batches = max(1, (len(companies) - 1) // batch_size + 1)
            print(f"\n    batch {i//batch_size + 1}/{total_batches}")

            for company in batch:
                cname = company["company"]
                info_type = company.get("info_type", "")
                remark = company.get("remark", "")
                queries = company_queries[cname]

                company_articles = []
                seen_links = set()

                for q_idx, q in enumerate(queries):
                    q_label = q if len(q) <= 20 else q[:20] + "..."
                    # print(f"      [{cname[:10]}] 搜索: {q_label}")
                    # 百度搜索为主，Tavily 做429限流自动降级
                    # 舆情用 pm（近一月），企业舆情新闻不是每天都有
                    if self._baidu_429_use_tavily:
                        sr = self.tavily_search(q, count=5, days=7)
                    else:
                        sr = self.baidu_search(q, count=5, freshness="pw")
                        if not sr:
                            sr = self.tavily_search(q, count=5, days=7)

                    if sr:
                        filtered = analyze_baidu_results(sr, cname, info_type)
                        # if len(filtered) < len(sr):
                        #     # print(f"        ↪ LLM过滤: {len(sr)}→{len(filtered)}条")
                        sr = filtered

                    for r in sr:
                        if r['link'] not in seen_links:
                            seen_links.add(r['link'])
                            r['company'] = cname
                            r['info_type'] = info_type
                            r['remark'] = remark
                            r['date'] = r.get('date', '')
                            company_articles.append(r)

                if company_articles:
                    stats["ok"] += 1
                    stats["articles"] += len(company_articles)
                    safe_name = re.sub(r'[^\w\u4e00-\u9fff]', '_', cname)
                    dept_prefix = dept_name.replace('_', '、')
                    fpath = os.path.join(raw_dir, f"{dept_prefix}_sentiment_{safe_name}.txt")
                    with open(fpath, "w", encoding="utf-8") as f:
                        for art in company_articles[:5]:
                            abstract_cleaned = self._clean_text(art.get('abstract', ''))
                            f.write("---\n")
                            f.write("title: {}\n".format(art.get('title', '')))
                            f.write("date: {}\n".format(art.get('date', '')))
                            f.write("source: {}\n".format(art.get('source_name', '')))
                            f.write("link: {}\n".format(art.get('link', '')))
                            f.write("abstract: {}\n".format(abstract_cleaned))
                            f.write("company: {}\n".format(cname))
                            f.write("info_type: {}\n".format(info_type))
                    # print(f"      ✅ {cname[:12]} ({min(5, len(company_articles))}条)")
                    articles.extend(company_articles[:5])
                else:
                    stats["fail"] += 1
                    print(f"      ⚠️ {cname[:12]} (无近一周结果)")
                    search_source = "Tavily" if queries else "百度"
                    self.failures.append({
                        "company": cname, "name": cname,
                        "error": f"{search_source}搜索无近一周舆情结果",
                    })

                time.sleep(random.uniform(1.0, 2.0))

        # 保存处室舆情汇总 JSON
        if articles:
            fpath_json = os.path.join(raw_dir, f"{dept_name}_sentiment_articles.json")
            with open(fpath_json, "w", encoding="utf-8") as f:
                json.dump(articles, f, ensure_ascii=False, indent=2)
            print(f"\n  ✅ 舆情汇总: {fpath_json} ({len(articles)}条)")

        self.results_log.append({"type": "sentiment", "stats": stats})
        return {"dept": dept_name, "articles": articles, "failures": self.failures,
                "stats": stats}

    def _clean_text(self, text: str) -> str:
        """简单清洗文本"""
        if not text:
            return ""
        text = re.sub(r'\s+', ' ', text).strip()
        return text[:500]


# ===================== Main =====================


def _append_sentiment_to_report(archive_dir: str, run_date_str: str, args,
                                 total_stats: dict, all_failures: list):
    """将舆情采集结果追加到现有采集运行报告.md"""
    import os
    report_path = os.path.join(archive_dir, "采集运行报告.md")

    report_lines = []
    report_lines.append("")
    report_lines.append("## 舆情采集")
    report_lines.append("")
    report_lines.append(f"**运行日期：** {run_date_str}")
    report_lines.append(f"**搜索范围：** {', '.join(args.departments)}")
    report_lines.append("")

    raw_dir = os.path.join(os.path.dirname(archive_dir), "raw")
    for dept_name in args.departments:
        config_path = os.path.join(CONFIG_DIR, f"{dept_name}.json")
        if not os.path.exists(config_path):
            continue
        with open(config_path, encoding="utf-8") as f:
            config = json.load(f)
        raw_dept = config.get("department", dept_name)
        companies = config.get("sentiment_companies", [])

        # 统计已生成的舆情.txt文件来确定实际结果
        dept_prefix = dept_name.replace('_', '、')
        sentiment_files = [f for f in os.listdir(raw_dir)
                          if f.startswith(f"{dept_prefix}_sentiment_") and f.endswith('.txt')]

        company_ok = 0
        company_fail = 0
        total_articles = 0
        company_results = []
        for company in companies:
            cname = company["company"]
            safe_name = re.sub(r'[^\w\u4e00-\u9fff]', '_', cname)
            expected_file = f"{dept_prefix}_sentiment_{safe_name}.txt"
            fpath = os.path.join(raw_dir, expected_file)
            if os.path.exists(fpath):
                with open(fpath, 'r', encoding='utf-8') as fr:
                    fc = fr.read()
                cnt = fc.count('\n---\n') + 1 if '---' in fc else (1 if fc.strip() else 0)
                company_ok += 1
                total_articles += cnt
                company_results.append((cname, "✅ 成功", f"获取 {cnt}条"))
            else:
                company_fail += 1
                err = ""
                for ff in all_failures:
                    if ff.get("name") == cname:
                        err = ff.get("error", "")
                        break
                company_results.append((cname, "❌ 失败", err or "无结果"))

        report_lines.append(f"### {raw_dept}")
        report_lines.append("")
        report_lines.append(f"配置企业: {len(companies)}家")
        report_lines.append("")
        report_lines.append("| 企业名称 | 状态 | 结果 |")
        report_lines.append("|---------|--------|------|")
        for cname, status, result_str in company_results:
            report_lines.append(f"| {cname} | {status} | {result_str} |")
        report_lines.append("")
        report_lines.append(f"**合计：** ✅ {company_ok}家成功 | ❌ {company_fail}家失败 | 共获取 {total_articles}条舆情")
        report_lines.append("")

    report_lines.append(f"---")
    report_lines.append(f"*舆情报告生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}*")
    report_lines.append("")

    # 如果现有报告存在，追加；否则创建
    report_content = "\n".join(report_lines)
    if os.path.exists(report_path):
        with open(report_path, "r", encoding="utf-8") as f:
            existing = f.read()
        existing += report_content
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(existing)
    else:
        with open(report_path, "w", encoding="utf-8") as f:
            f.write("# 采集运行报告")
            f.write(report_content)
    print(f"\n📋 舆情结果已追加至: {report_path}")


def _merge_sentiment_to_archive(archive_dir: str, run_date_str: str, dept_names: list):
    """将舆情数据合并到 all_articles.json，供简报生成使用"""
    raw_dir = os.path.join(DATA_DIR, run_date_str, "raw")
    archive_path = os.path.join(archive_dir, "all_articles.json")

    if not os.path.exists(archive_path):
        print(f"  ⚠️ 无archive/all_articles.json，无法合并舆情数据")
        return

    with open(archive_path, "r", encoding="utf-8") as f:
        all_articles = json.load(f)

    merged_total = 0
    for dept_name in dept_names:
        dept_key = dept_name.replace('_', '、')
        sent_json = os.path.join(raw_dir, f"{dept_name}_sentiment_articles.json")
        if not os.path.exists(sent_json):
            sent_json = os.path.join(raw_dir, f"{dept_key}_sentiment_articles.json")
        if not os.path.exists(sent_json):
            continue

        with open(sent_json, "r", encoding="utf-8") as f:
            sent_articles = json.load(f)
        if not sent_articles:
            continue

        target_key = dept_key if dept_key in all_articles else dept_name
        if target_key not in all_articles:
            all_articles[target_key] = []

        existing_links = {a.get('link', '') for a in all_articles[target_key]}
        added = 0
        for sa in sent_articles:
            if sa.get('link') and sa['link'] in existing_links:
                continue
            all_articles[target_key].append({
                'title': sa.get('title', ''),
                'date': sa.get('date', ''),
                'source_name': sa.get('source_name', '百度搜索'),
                'link': sa.get('link', ''),
                'summary': sa.get('abstract', sa.get('summary', '')),
                'company': sa.get('company', ''),
                'info_type': sa.get('info_type', ''),
            })
            if sa.get('link'):
                existing_links.add(sa['link'])
            added += 1

        dept_display = dept_key if dept_key in all_articles else dept_name
        merged_total += added
        print(f"  ➕ 舆情已合并到 {dept_display}: {added}条")

    with open(archive_path, "w", encoding="utf-8") as f:
        json.dump(all_articles, f, ensure_ascii=False, indent=2)
    print(f"  ✅ 舆情数据已集成到 {archive_path} (新增{merged_total}条)")

def main():
    parser = argparse.ArgumentParser(description='舆情采集（百度搜索 + LLM 过滤）')
    parser.add_argument('--departments', nargs='+', default=DEPARTMENTS,
                        help='要运行的处室名称 (默认: 全部)')
    parser.add_argument('--date', type=str, default='',
                        help='运行日期 (格式: YYYY-MM-DD)')
    parser.add_argument('--max-companies', type=int, default=0,
                        help='每处室最大企业数 (0=全部)')
    args = parser.parse_args()

    run_date_str = args.date or datetime.now().strftime("%Y-%m-%d")
    raw_dir = os.path.join(DATA_DIR, run_date_str, "raw")
    os.makedirs(raw_dir, exist_ok=True)

    crawler = SentimentCrawler(run_date_str=run_date_str)
    total_stats = {"ok": 0, "fail": 0, "articles": 0}
    all_failures = []

    for dept_name in args.departments:
        config_path = os.path.join(CONFIG_DIR, f"{dept_name}.json")
        if not os.path.exists(config_path):
            print(f"⚠️ 配置文件不存在: {config_path}")
            continue

        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)

        result = crawler.crawl_dept(config, raw_dir,
                                     max_companies=args.max_companies)
        stats = result.get("stats", {})
        total_stats["ok"] += stats.get("ok", 0)
        total_stats["fail"] += stats.get("fail", 0)
        total_stats["articles"] += stats.get("articles", 0)
        all_failures.extend(result.get("failures", []))

    print(f"\n{'='*60}")
    print(f"  ✔️ 舆情采集完成")
    print(f"  搜索成功: {total_stats['ok']} 家")
    print(f"  搜索失败: {total_stats['fail']} 家")
    print(f"  获取舆情: {total_stats['articles']} 条")
    if all_failures:
        print(f"  失败详情:")
        for ff in all_failures[:5]:
            print(f"    - {ff.get('name', '')}: {ff.get('error', '')}")
    
    # === 生成 archive/采集运行报告.md ===
    archive_dir = os.path.join(DATA_DIR, run_date_str, "archive")
    os.makedirs(archive_dir, exist_ok=True)
    _append_sentiment_to_report(archive_dir, run_date_str, args, total_stats, all_failures)

    # === 将舆情数据集成到 all_articles.json ===
    _merge_sentiment_to_archive(archive_dir, run_date_str, args.departments)
if __name__ == "__main__":
    main()
