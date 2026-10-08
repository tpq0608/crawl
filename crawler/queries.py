# -*- coding: utf-8 -*-
"""Tạo các biến thể câu truy vấn tìm kiếm cho 1 thảm họa trên 1 đài."""
import re
from datetime import datetime

from .config import MAX_PLACES, MAX_QUERIES_PER_OUTLET
from .utils import _num, _cell
from .vocabulary import HAZARD_SYNONYMS, KW_STOP
from .disaster import places, anglicize, provinces, event_names


# ======================= TẠO QUERY =======================
def build_queries(row, site, country_name):
    """Nhiều biến thể query cho 1 đài. Ưu tiên: tên bão -> Subtype/Origin/Associated -> tỉnh -> cả nước."""
    dtype = _cell(row, "Disaster Type")
    syns = HAZARD_SYNONYMS.get(dtype, [dtype.lower()])
    year = _num(row.get("Start Year")) or ""
    month = _num(row.get("Start Month"))
    month_name = datetime(2000, month, 1).strftime("%B") if month and 1 <= month <= 12 else ""
    loc = row.get("Location")
    provs = [anglicize(x) for x in provinces(loc)]
    pl = (provs or places(loc))[:MAX_PLACES]

    def first_phrase(*cols):
        txt = re.sub(r"\(.*?\)", "", _cell(row, *cols))
        for part in re.split(r"[,;/&]", txt):
            part = part.strip().lower()
            if len(part) >= 4 and part not in KW_STOP:
                return part
        return ""

    sub = first_phrase("Disaster Subtype")
    origin = first_phrase("Origin")
    assoc = first_phrase("Associated Types", "Associated Disaster Types", "Associated Dis")

    qs = []
    for n in event_names(row)[:3]:                      # tên quốc tế + tên địa phương
        qs.append(f'"{n}" {syns[0]} {year} site:{site}')
    for place in pl:
        qs.append(f"{syns[0]} {place} {month_name} {year} site:{site}")
    if sub and sub != syns[0]:
        for place in pl[:1] or [country_name]:
            qs.append(f"{sub} {place} {month_name} {year} site:{site}")
    qs.append(f"{syns[0]} {country_name} {month_name} {year} site:{site}")
    qs.append(f"{syns[0]}s {country_name} {year} {month_name} deaths displaced site:{site}")
    if origin:
        qs.append(f"{origin} {syns[0]} {pl[0] if pl else country_name} {month_name} {year} site:{site}")
    if assoc and assoc not in (sub, syns[0]):
        qs.append(f"{assoc} {pl[0] if pl else country_name} {month_name} {year} site:{site}")
    for syn in syns[1:]:
        qs.append(f"{syn} {country_name} {month_name} {year} site:{site}")
    if pl:
        qs.append(f"{syns[0]} {pl[0]} {year} damage evacuated site:{site}")

    seen, uniq = set(), []
    for q in qs:
        q = re.sub(r"\s+", " ", q).strip()
        if q not in seen:
            seen.add(q)
            uniq.append(q)
    return uniq[:MAX_QUERIES_PER_OUTLET] if MAX_QUERIES_PER_OUTLET else uniq
