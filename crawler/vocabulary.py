# -*- coding: utf-8 -*-
"""Từ điển từ khóa / regex về thiên tai và địa danh (chỉ là dữ liệu, không có hàm)."""
import re


PLACE_NOISE = re.compile(r"\b(provinces?|districts?|cities|city|towns?|regency|regencies|"
                         r"prefecture|region|areas?|nationwide)\b", re.I)


PROVINCE_WORDS = {"Jawa": "Java", "Sumatera": "Sumatra", "Nangroe Aceh Darussalam": "Aceh",
                  "Dki Jakarta": "Jakarta", "Kepulauan": "Islands"}
DIRECTIONS = {"Barat": "West", "Timur": "East", "Tengah": "Central", "Selatan": "South", "Utara": "North"}


GENERIC_WORDS = ["storm", "typhoon", "tropical", "depression", "cyclone", "super", "severe",
                 "flood", "floods", "the", "no", "number", "slide", "landslide"]
NOT_A_NAME = set(GENERIC_WORDS) | {"rain", "rains", "flooding", "heavy"}


HAZARD_SYNONYMS = {
    "Flood": ["flood", "flooding"],
    "Storm": ["typhoon", "tropical storm"],
    "Mass movement (wet)": ["landslide"],
    "Drought": ["drought"],
    "Extreme temperature": ["heat wave", "cold spell"],
    "Wildfire": ["wildfire"],
}

HAZARD_RE = re.compile(
    r"\b(storms?|typhoons?|cyclones?|depressions?|floods?|flooding|flooded|inundat\w*|landslides?|"
    r"mudslides?|heavy rain\w*|rainfall|drought|heat ?waves?|cold (?:spell|wave|snap)|saltwater|"
    r"salinity|disasters?|damages?d?)\b", re.I)

TYPE_RE = {
    "Flood": re.compile(r"\b(floods?|flooding|flooded|inundat\w*|heavy rain\w*|rainfall|overflow\w*)\b", re.I),
    "Storm": re.compile(r"\b(storms?|typhoons?|cyclones?|depressions?|gales?)\b", re.I),
    "Mass movement (wet)": re.compile(r"\b(landslides?|mudslides?|mass movement|slope failure)\b", re.I),
    "Drought": re.compile(r"\b(drought|dry spell|water shortage|saltwater|salinity)\b", re.I),
    "Extreme temperature": re.compile(r"\b(heat ?waves?|cold (?:spell|wave|snap)|extreme (?:heat|cold)|temperatures?)\b", re.I),
}


# ======================= PHÂN BIỆT SỰ KIỆN: SUBTYPE / ORIGIN / ASSOCIATED TYPES =======================
# (khóa nhận diện trong cột EM-DAT, [từ khóa thường xuất hiện trong bài báo]).
# Khóa CỤ THỂ đặt trước khóa chung; mỗi cụm trong ô chỉ khớp khóa đầu tiên tìm thấy.
HAZARD_KW = [
    ("flash flood",     ["flash flood", "flash floods", "flash-flood", "flash-floods"]),
    ("riverine",        ["riverine", "river flood", "river floods", "river overflow", "river burst",
                         "overflowing river", "rivers overflowed", "burst its banks", "burst their banks"]),
    ("storm surge",     ["storm surge", "tidal surge", "coastal flood", "coastal flooding"]),
    ("coastal",         ["coastal flood", "coastal flooding", "storm surge", "tidal flood", "high tide",
                         "king tide", "sea level"]),
    ("flood",           ["flood", "floods", "flooding", "flooded", "inundated"]),
    ("tropical cyclone", ["typhoon", "tropical storm", "tropical cyclone", "tropical depression",
                          "cyclone", "super typhoon", "low pressure area"]),
    ("convective",      ["thunderstorm", "thunderstorms", "lightning", "tornado", "hailstorm", "squall",
                         "whirlwind", "waterspout", "windstorm", "strong winds", "gusts"]),
    ("extra-tropical",  ["extratropical", "extra-tropical", "northeast monsoon", "southwest monsoon",
                         "cold surge", "monsoon trough"]),
    ("extratropical",   ["extratropical", "extra-tropical", "northeast monsoon", "southwest monsoon",
                         "cold surge", "monsoon trough"]),
    ("storm",           ["storm", "storms", "gale", "gales"]),
    ("mudslide",        ["mudslide", "mudslides", "mudflow", "debris flow", "lahar", "mud flow"]),
    ("landslide",       ["landslide", "landslides", "landslip", "slope failure", "earth slip"]),
    ("mass movement",   ["landslide", "landslides", "mudslide", "slope failure"]),
    ("ground movement", ["landslide", "sinkhole", "ground subsidence"]),
    ("rockfall",        ["rockfall", "rock fall", "falling rocks"]),
    ("avalanche",       ["avalanche", "snow slide"]),
    ("heat",            ["heat wave", "heatwave", "extreme heat", "scorching", "record temperature",
                         "heat index"]),
    ("cold",            ["cold wave", "cold spell", "cold snap", "frost", "freezing", "cold front"]),
    ("winter",          ["severe winter", "cold wave", "cold spell", "frost", "snowfall"]),
    ("drought",         ["drought", "dry spell", "water shortage", "el nino", "saltwater intrusion",
                         "salinity", "crop failure"]),
    ("fire",            ["forest fire", "wildfire", "bushfire", "peatland fire", "haze", "hotspots"]),
    # ---- các giá trị hay gặp ở cột Origin ----
    ("heavy rain",      ["heavy rain", "heavy rains", "heavy rainfall", "torrential rain", "downpour",
                         "downpours", "incessant rain", "prolonged rain"]),
    ("monsoon",         ["monsoon", "monsoon rains", "habagat", "amihan"]),
    ("typhoon",         ["typhoon", "super typhoon"]),
    ("dam",             ["dam", "reservoir", "spillway", "dam release", "dam collapse", "dam break",
                         "released water"]),
    ("tide",            ["high tide", "tidal", "king tide"]),
    ("el nino",         ["el nino"]),
    ("la nina",         ["la nina"]),
    ("snowmelt",        ["snowmelt", "melting snow"]),
    ("rain",            ["rain", "rains", "rainfall", "downpour"]),
]

KW_STOP = {"general", "wet", "dry", "other", "unspecified", "unknown", "nan", "none", "na", "n/a",
           "snow", "debris", "rock", "and", "the", "of"}
