# -*- coding: utf-8 -*-
"""Tìm link bài báo qua công cụ tìm kiếm (thư viện ddgs / duckduckgo_search)."""
import random
import time
from urllib.parse import urlparse, urlunparse, parse_qsl, urlencode


try:                                   # package mới tên là "ddgs"
    from ddgs import DDGS
    HAVE_DDGS = True
except ImportError:
    from duckduckgo_search import DDGS
    HAVE_DDGS = False


from .config import RESULTS_PER_QUERY, SLEEP_RANGE, EXCLUDED_DOMAINS, JUNK_URL_PATTERNS


SEARCH_BACKENDS = ["duckduckgo", "bing", "brave", "google"] if HAVE_DDGS else [None]


# ======================= TÌM LINK BẰNG DUCKDUCKGO =======================
def url_matches_site(url, site):
    """Link phải thật sự thuộc đài đang cào (site: của công cụ tìm kiếm đôi khi trả link lạ)."""
    p = urlparse(url)
    host = p.netloc.lower().split(":")[0]
    base, _, path = site.partition("/")
    if not (host == base or host.endswith("." + base)):
        return False
    return p.path.startswith("/" + path) if path else True


def norm_url(u):
    """Chuẩn hóa để so trùng: bỏ #fragment, tham số utm_*, dấu / cuối."""
    p = urlparse(u.strip())
    q = [(k, v) for k, v in parse_qsl(p.query) if not k.lower().startswith("utm_")]
    return urlunparse((p.scheme, p.netloc.lower(), p.path.rstrip("/"), "", urlencode(q), ""))


def _good_urls(results, site):
    urls = []
    for r in results or []:
        u = r.get("href", "")
        if (u.startswith("http")
                and url_matches_site(u, site)
                and not any(bad in u.lower() for bad in EXCLUDED_DOMAINS)
                and not JUNK_URL_PATTERNS.search(u)
                and u not in urls):
            urls.append(u)
    return urls


def search_article_urls(query, site, retries=2):
    """Thử lần lượt từng công cụ tìm kiếm.
    Trả về [] nếu query rỗng thật, None nếu tất cả công cụ đều lỗi kết nối."""
    for attempt in range(retries):
        errors, got_empty = [], False
        for backend in SEARCH_BACKENDS:
            try:
                kw = {"backend": backend} if backend else {}
                results = DDGS().text(query, max_results=RESULTS_PER_QUERY, **kw)
                urls = _good_urls(results, site)
                if urls:
                    time.sleep(random.uniform(*SLEEP_RANGE))
                    return urls
                got_empty = True
            except Exception as e:
                if "no results" in str(e).lower():
                    got_empty = True
                else:
                    errors.append(f"{backend or 'default'}: {str(e)[:80]}")
                    print(f"   ⚠️ Công cụ '{backend or 'default'}' lỗi -> thử công cụ khác...")
            time.sleep(random.uniform(0.5, 1.5))
        if got_empty:
            print("      (không có kết quả)")
            return []
        wait = 20 * (attempt + 1)
        print(f"   ⚠️ Tất cả công cụ tìm kiếm đều lỗi ({errors[-1]}). Chờ {wait}s rồi thử lại "
              f"({attempt + 1}/{retries}). Kiểm tra mạng/VPN nếu lỗi kéo dài.")
        time.sleep(wait)
    return None
