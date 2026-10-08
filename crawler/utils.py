# -*- coding: utf-8 -*-
"""Tiện ích đọc ô dữ liệu và chuẩn hóa chuỗi."""
import re
import unicodedata

import pandas as pd


# ======================= TIỆN ÍCH DỮ LIỆU =======================
def _num(x):
    if x is None or (not isinstance(x, str) and pd.isna(x)):
        return None
    try:
        return int(float(str(x).strip()))
    except (ValueError, TypeError):
        return None


def _cell(row, *names):
    """Lấy giá trị ô (thử nhiều tên cột), trả '' nếu trống/NaN."""
    for n in names:
        v = row.get(n)
        if v is None or (not isinstance(v, str) and pd.isna(v)):
            continue
        s = str(v).strip()
        if s and s.lower() != "nan":
            return s
    return ""


def _fold(s):
    """Bỏ dấu + lower để so khớp (Quảng Bình == quang binh)."""
    s = unicodedata.normalize("NFD", str(s).replace("đ", "d").replace("Đ", "D"))
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


_EN_STOP = {"the", "and", "of", "to", "in", "is", "that", "for", "with", "on", "as", "was", "are", "by",
            "at", "from", "it", "this", "be", "have", "has", "said", "were", "which", "an", "their"}
_ID_STOP = {"yang", "dan", "di", "untuk", "dengan", "dari", "pada", "ini", "itu", "adalah", "dalam",
            "akan", "tidak", "juga", "oleh", "karena", "ke", "sebagai", "telah", "para", "atau",
            "saat", "mereka", "ada", "sebuah", "kepada", "bahwa"}


def is_english(text, min_ratio=0.15, sample=400):
    """Đoán bài có phải tiếng Anh không: đếm từ dừng tiếng Anh vs tiếng Indonesia trong `sample` từ đầu.
    Không cần thư viện ngoài. Bài tiếng Indonesia có ít từ dừng tiếng Anh nên bị loại."""
    words = re.findall(r"[a-z']+", _fold(text))[:sample]
    if not words:
        return False
    en = sum(1 for w in words if w in _EN_STOP)
    idn = sum(1 for w in words if w in _ID_STOP)
    return en / len(words) >= min_ratio and en > 2 * idn