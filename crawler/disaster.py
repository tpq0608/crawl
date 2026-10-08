# -*- coding: utf-8 -*-
"""Đọc thông tin 1 dòng thảm họa EM-DAT: khoảng ngày, địa danh, tên bão, cụm từ khóa."""
import difflib
import re
from datetime import datetime, timedelta

import pandas as pd

from .config import DAYS_BEFORE, DAYS_AFTER
from .utils import _num, _cell, _fold
from .vocabulary import PLACE_NOISE, PROVINCE_WORDS, DIRECTIONS, GENERIC_WORDS, NOT_A_NAME, HAZARD_KW, KW_STOP


def disaster_core(row):
    """(ngày bắt đầu, ngày kết thúc) thật của thảm họa, không cộng đệm."""
    try:
        sy, sm = _num(row.get("Start Year")), _num(row.get("Start Month"))
        if not sy or not sm or not (1 <= sm <= 12):
            return None
        sd = max(1, min(_num(row.get("Start Day")) or 1, 28))
        ey = _num(row.get("End Year")) or sy
        em = max(1, min(_num(row.get("End Month")) or sm, 12))
        ed = max(1, min(_num(row.get("End Day")) or 28, 28))
        return datetime(sy, sm, sd), datetime(ey, em, ed)
    except Exception:
        return None


def disaster_window(row):
    core = disaster_core(row)
    if core is None:
        return None
    return core[0] - timedelta(days=DAYS_BEFORE), core[1] + timedelta(days=DAYS_AFTER)


def date_ok(date_str, window):
    if window is None or not date_str:
        return True  # Không có ngày -> giữ lại để kiểm tay
    try:
        d = datetime.strptime(str(date_str)[:10], "%Y-%m-%d")
    except ValueError:
        return True
    return window[0] <= d <= window[1]


def places(location, limit=None, include_parens=False):
    """Tách cột Location thành danh sách địa danh."""
    if location is None or (not isinstance(location, str) and pd.isna(location)):
        return []
    s = str(location)
    s = re.sub(r"[()]", ",", s) if include_parens else re.sub(r"\(.*?\)", "", s)
    out = []
    for part in re.split(r"[,;]| and ", s):
        part = PLACE_NOISE.sub("", part).strip(" .-")
        if len(part) >= 3 and part.lower() not in [p.lower() for p in out]:
            out.append(part)
    return out[:limit] if limit else out


def anglicize(name):
    """Tên tỉnh Indonesia -> cách báo tiếng Anh hay viết: 'Jawa Barat' -> 'West Java'."""
    for k, v in PROVINCE_WORDS.items():
        name = re.sub(rf"\b{k}\b", v, name, flags=re.I)
    words = name.split()
    if len(words) >= 2 and words[-1].capitalize() in DIRECTIONS:
        words = [DIRECTIONS[words[-1].capitalize()]] + words[:-1]
    return " ".join(words)


def provinces(location):
    """Lấy tên tỉnh dạng '... (X province)' trong cột Location (định dạng của Indonesia/Philippines)."""
    if location is None or (not isinstance(location, str) and pd.isna(location)):
        return []
    out = []
    for grp in re.findall(r"\(([^()]*province[^()]*)\)", str(location), flags=re.I):
        for part in re.split(r"[,;/]", PLACE_NOISE.sub("", grp)):
            part = part.strip(" .-")
            if len(part) >= 3 and part.lower() not in [x.lower() for x in out]:
                out.append(part)
    return out


def _is_generic(w):
    return w in GENERIC_WORDS or bool(difflib.get_close_matches(w, GENERIC_WORDS, n=1, cutoff=0.82))


def event_names(row):
    """Mọi tên bão (quốc tế + địa phương) trong cột Event Name.
    "Typhoon 'Damrey' / 'Ramil'" -> ['damrey','ramil'];  "Cyclone 'Seroja-21'" -> ['seroja-21','seroja'];
    Sự kiện tên chung chung như 'slide' -> [] (coi như không có tên)."""
    event = _cell(row, "Event Name")
    if not event:
        return []
    names = []
    for piece in re.split(r"[()/&,;]|\band\b", event.replace("\u2019", "'")):
        words = [w for w in re.findall(r"[a-z0-9-]+", _fold(piece).replace("'", " ")) if not _is_generic(w)]
        if not words:
            continue
        full = " ".join(words)
        for n in (full, re.sub(r"-\d+$", "", full)):
            if len(n) >= 3 and not n.isdigit() and n not in NOT_A_NAME and n not in names:
                names.append(n)
    return names


def term_groups(text):
    """Cột EM-DAT (Subtype / Origin / Associated Types) -> list regex; mỗi regex = 1 cụm trong ô.
    VD 'Flash flood, Landslide' -> [regex(flash flood...), regex(landslide...)]."""
    if text is None or (not isinstance(text, str) and pd.isna(text)):
        return []
    groups, seen = [], set()
    for phrase in re.split(r"[,;/&()]|\band\b", _fold(text)):
        phrase = phrase.strip(" .-'\"")
        if len(phrase) < 3 or phrase in KW_STOP:
            continue
        terms = next((kws for key, kws in HAZARD_KW if key in phrase), None)
        if terms is None:
            if len(phrase.split()) > 4:
                continue
            terms = [phrase]
        pat = r"\b(?:" + "|".join(re.escape(_fold(t)) for t in terms) + r")\b"
        if pat not in seen:
            seen.add(pat)
            groups.append(re.compile(pat))
    return groups


def place_tokens(row):
    toks = set()
    for p in places(row.get("Location"), include_parens=True):
        for name in (p, anglicize(p)):
            t = _fold(name)
            if len(t) >= 4:
                toks.add(t)
    return toks
