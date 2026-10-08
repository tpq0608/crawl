# -*- coding: utf-8 -*-
"""Phân biệt sự kiện: hồ sơ từng thảm họa, chấm điểm bài báo, gán bài cho sự kiện phù hợp nhất."""
import os
import re
from datetime import datetime

import pandas as pd

from .config import NAMED_EVENT_STRICT, STRICT_PLACE, AMBIGUOUS_MARGIN, OTHER_NAMED_PENALTY
from .utils import _cell, _fold
from .vocabulary import HAZARD_RE, TYPE_RE
from .disaster import (disaster_core, disaster_window, date_ok, event_names, term_groups,
                       place_tokens)


def build_profiles(df):
    """Hồ sơ cho MỌI sự kiện của nước (kể cả sự kiện ngoài --year/--limit) để so sánh với nhau.
    rivals[code] = các sự kiện cùng nước có cửa sổ ngày chồng lên nhau (kể cả chính nó)."""
    profiles = {}
    for _, row in df.iterrows():
        code = row["Disaster Name"]
        dtype = _cell(row, "Disaster Type")
        names = event_names(row)
        profiles[code] = {
            "code": code, "row": row,
            "window": disaster_window(row), "core": disaster_core(row),
            "names": names,
            "name_rx": re.compile(r"\b(?:" + "|".join(re.escape(n) for n in names) + r")\b") if names else None,
            "ptoks": place_tokens(row),
            "type_rx": TYPE_RE.get(dtype, HAZARD_RE),
            "sub": term_groups(_cell(row, "Disaster Subtype")),
            "origin": term_groups(_cell(row, "Origin")),
            "assoc": term_groups(_cell(row, "Associated Types", "Associated Disaster Types", "Associated Dis")),
            "rivals": [],
        }
    for code, p in profiles.items():
        w = p["window"]
        p["rivals"] = [code] + [
            c for c, q in profiles.items()
            if c != code and w and q["window"] and w[0] <= q["window"][1] and q["window"][0] <= w[1]]
    return profiles


def prep(content):
    """Tính sẵn bản 'bỏ dấu + lower' của bài (dùng cho mọi phép so khớp)."""
    if "_blob" not in content:
        content["_title"] = _fold(content["title"])
        content["_blob"] = _fold(f"{content['title']} {content['full_text']}")
    return content


def is_relevant(content, prof):
    """Nhận bài nếu thỏa MỘT trong hai nhánh:
    (a) thảm họa có tên bão: bài nhắc tên bão VÀ có từ thiên tai (HAZARD_RE);
    (b) bài nói đúng LOẠI thiên tai VÀ (nhắc một địa danh trong Location HOẶC tiêu đề nói về loại thiên tai đó).
    Nhánh (b) áp dụng cho mọi thảm họa, kể cả bão có tên, nên bài chỉ viết "the typhoon" vẫn được giữ.
    NAMED_EVENT_STRICT = True: bão có tên chỉ xét nhánh (a) (luật cũ)."""
    prep(content)
    blob = content["_blob"]
    name_ok = (bool(prof["name_rx"]) and bool(prof["name_rx"].search(blob))
               and bool(HAZARD_RE.search(blob)))
    if prof["name_rx"] and NAMED_EVENT_STRICT:
        return name_ok
    if name_ok:                                       # nhánh (a)
        return True
    rx = prof["type_rx"]                              # nhánh (b)
    if not rx.search(blob):
        return False
    if not STRICT_PLACE:
        return True
    toks = prof["ptoks"]
    if not toks or any(t in blob for t in toks):
        return True
    return bool(rx.search(content["_title"]))


def date_score(date_str, core):
    """4 điểm nếu bài nằm trong thời gian thảm họa, giảm dần về 0 khi cách xa 60 ngày."""
    if not core or not date_str:
        return 2.0
    try:
        d = datetime.strptime(str(date_str)[:10], "%Y-%m-%d")
    except ValueError:
        return 2.0
    gap = 0 if core[0] <= d <= core[1] else min(abs((d - core[0]).days), abs((d - core[1]).days))
    return 4.0 - 4.0 * min(gap, 60) / 60


def score_article(content, p, named_hits):
    """Điểm 'bài này thuộc sự kiện p đến mức nào'."""
    blob, title = content["_blob"], content["_title"]
    s = 0.0
    if p["name_rx"]:                                              # tên bão: tín hiệu mạnh nhất
        if p["name_rx"].search(title):
            s += 12
        elif p["name_rx"].search(blob):
            s += 10
    if p["type_rx"].search(blob):                                 # Disaster Type
        s += 2
    s += 3.0 * min(2, sum(1 for g in p["sub"] if g.search(blob)))      # Disaster Subtype
    s += 2.0 * min(2, sum(1 for g in p["origin"] if g.search(blob)))   # Origin
    s += 1.5 * min(3, sum(1 for g in p["assoc"] if g.search(blob)))    # Associated Types
    toks = p["ptoks"]
    if toks:                                                      # Location
        s += min(4, sum(1 for t in toks if t in blob))
        if any(t in title for t in toks):
            s += 1.5
    s += date_score(content["date"], p["core"])                   # Start/End date
    if named_hits and p["code"] not in named_hits:                # bài đang nói về bão có tên của sự kiện khác
        s -= OTHER_NAMED_PENALTY
    return s


def pick_event(content, code, profiles):
    """Bài này thuộc sự kiện nào trong nhóm cạnh tranh (cùng nước, cùng thời gian)?
    Trả về (code_thắng | None, thông tin mơ hồ | None). None = không sự kiện nào nhận bài."""
    prep(content)
    rivals = profiles[code]["rivals"]
    named_hits = {rc for rc in rivals
                  if profiles[rc]["name_rx"] and profiles[rc]["name_rx"].search(content["_blob"])}
    cands = []
    for rc in rivals:
        p = profiles[rc]
        if date_ok(content["date"], p["window"]) and is_relevant(content, p):
            cands.append((score_article(content, p, named_hits), rc))
    if not cands:
        return None, None
    cands.sort(key=lambda x: (-x[0], x[1]))
    amb = None
    if len(cands) > 1 and cands[0][0] - cands[1][0] < AMBIGUOUS_MARGIN:
        amb = {"assigned": cands[0][1], "score": round(cands[0][0], 1),
               "runner_up": cands[1][1], "runner_score": round(cands[1][0], 1)}
    return cands[0][1], amb


def log_ambiguous(country, outlet, content, url, amb):
    """Ghi các bài khó phân biệt (2 sự kiện điểm gần nhau) ra file riêng để kiểm tay."""
    path = f"ambiguous_{country}.csv"
    rec = {"URL": url, "OUTLET": outlet, "Tittle": content["title"], "Date": content["date"],
           "Assigned to": amb["assigned"], "Score": amb["score"],
           "Runner-up": amb["runner_up"], "Runner-up score": amb["runner_score"]}
    pd.DataFrame([rec]).to_csv(path, mode="a", header=not os.path.exists(path),
                               index=False, encoding="utf-8-sig")