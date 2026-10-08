# -*- coding: utf-8 -*-
"""Tải trang và trích nội dung bài báo (HTML thường + WordPress REST API)."""
import html as htmllib
import json
import re
import time
import urllib.request
from urllib.parse import quote

import trafilatura
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

from .config import MIN_WORDS, WP_PER_PAGE, WP_MAX_PAGES


# ======================= CÀO NỘI DUNG =======================
def clean_text(s):
    if not s:
        return ""
    return ILLEGAL_CHARACTERS_RE.sub("", str(s)).strip()[:32000]   # giới hạn 32767 ký tự/ô của Excel


BROWSER_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/124.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
    "Accept-Encoding": "identity",
}


def download_html(url):
    html = trafilatura.fetch_url(url)
    if html:
        return html
    try:                                              # nhiều báo chặn bot -> thử lại với header giống trình duyệt
        req = urllib.request.Request(url, headers=BROWSER_HEADERS)
        with urllib.request.urlopen(req, timeout=25) as r:
            raw = r.read()
            charset = r.headers.get_content_charset() or "utf-8"
        return raw.decode(charset, errors="replace")
    except Exception:
        return None


def extract_article_content(url):
    """Trả về (bài, lý do_lỗi). lý do: 'tải lỗi' | 'quá ngắn' | 'lỗi đọc'."""
    try:
        downloaded = download_html(url)
        if not downloaded:
            return None, "tải lỗi"
        meta = trafilatura.extract_metadata(downloaded)
        text = trafilatura.extract(downloaded, include_comments=False, include_tables=False)
        if not text or len(text.split()) < MIN_WORDS:
            return None, "quá ngắn"
        return {
            "title": clean_text(meta.title) if meta and meta.title else "No Title",
            "full_text": clean_text(text),
            "date": str(meta.date)[:10] if meta and meta.date else "",
        }, ""
    except Exception:
        return None, "lỗi đọc"


# ---------- Đài chạy WordPress: hỏi thẳng API (lọc theo ngày, có sẵn nội dung) ----------
WP_CACHE = {}


def _http_json(url, timeout=25):
    req = urllib.request.Request(url, headers={**BROWSER_HEADERS, "Accept": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8", errors="replace"))


def wp_available(site):
    if "/" in site:                                    # site có path (vd vietnamnet.vn/en) -> không dùng
        return False
    if site in WP_CACHE:
        return WP_CACHE[site]
    try:
        ok = isinstance(_http_json(f"https://{site}/wp-json/wp/v2/posts?per_page=1&_fields=id"), list)
    except Exception:
        ok = False
    WP_CACHE[site] = ok
    return ok


def _strip_tags(h):
    return re.sub(r"\s+", " ", htmllib.unescape(re.sub(r"<[^>]+>", " ", h or ""))).strip()


def wp_search(site, term, window):
    """Tìm bài theo từ khóa trong khoảng ngày của thảm họa. Trả về [{url,title,html,date}]."""
    q = f"https://{site}/wp-json/wp/v2/posts?search={quote(term)}&per_page={WP_PER_PAGE}&_fields=link,date,title,content"
    if window:
        q += f"&after={window[0]:%Y-%m-%d}T00:00:00&before={window[1]:%Y-%m-%d}T23:59:59"
    out = []
    for page in range(1, WP_MAX_PAGES + 1):
        try:
            data = _http_json(q + f"&page={page}")
        except Exception:
            break
        if not isinstance(data, list):
            break
        for d in data:
            try:
                out.append({"url": d["link"], "title": _strip_tags(d["title"]["rendered"]),
                            "html": d["content"]["rendered"], "date": str(d.get("date", ""))[:10]})
            except Exception:
                continue
        if len(data) < WP_PER_PAGE:
            break
        time.sleep(1)
    return out


def wp_to_content(item):
    """Bài WordPress -> (content, lý do lỗi) giống extract_article_content."""
    text = None
    try:
        text = trafilatura.extract(item["html"], include_comments=False, include_tables=False)
    except Exception:
        pass
    text = text or _strip_tags(item["html"])
    if not text or len(text.split()) < MIN_WORDS:
        return None, "quá ngắn"
    return {"title": clean_text(item["title"]) or "No Title", "full_text": clean_text(text),
            "date": item["date"]}, ""
