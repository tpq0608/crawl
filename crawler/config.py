# -*- coding: utf-8 -*-
"""Cấu hình: file dữ liệu, giới hạn, cờ bật/tắt, danh sách nước và đài báo."""
import re


# ======================= CẤU HÌNH =======================
DISASTER_FILE = "[8386] DISASTER DATA.xlsx"
DISASTER_SHEET = "(ĐÃ LỌC) THEO QUỐC GIA"
NEWS_TEMPLATE = "[8386] NEWS DATA.xlsx"                  # File mẫu: KHÔNG bị ghi đè
OUTPUT_PATTERN = "[8386] NEWS DATA - {country}.xlsx"     # Mỗi nước 1 file
MERGED_OUTPUT = "[8386] NEWS DATA - MERGED.xlsx"

# ---- KHÔNG GIỚI HẠN số bài: cào hết mọi đài của nước, mọi query ----
RESULTS_PER_QUERY = 30      # số link mỗi query (công cụ tìm kiếm có thể trả ít hơn)
MAX_QUERIES_PER_OUTLET = None   # None = chạy hết mọi biến thể query trên mỗi đài
MAX_PLACES = 3              # số địa danh đầu tiên trong cột Location dùng để tạo query
MIN_WORDS = 100
DAYS_BEFORE, DAYS_AFTER = 7, 45   # cửa sổ ngày hợp lệ quanh thảm họa
FIRST_DATA_ROW = 3

DOWNLOAD_WORKERS = 6        # số bài tải song song
USE_WP_API = True           # Đài dùng WordPress: lấy bài qua API có lọc ngày. Tự phát hiện, không có thì bỏ qua
WP_PER_PAGE = 50
WP_MAX_PAGES = 10           # số trang API tối đa cho mỗi từ khóa
AUTO_ARTICLE_ID = False     # False: để trống cột Article ID (đánh tay). True: tự đánh ID001, VN001...
NAMED_EVENT_STRICT = False  # False: nhận bài nếu nhắc tên bão HOẶC (đúng loại + địa danh). True: bão có tên bắt buộc nhắc tên (luật cũ)
STRICT_PLACE = True         # Bài (nhánh đúng loại thiên tai) phải nhắc ít nhất 1 địa danh trong Location, hoặc tiêu đề nói về loại thiên tai
SLEEP_RANGE = (2, 4)        # nghỉ giữa các lần search (giây), tránh bị chặn


# ---- Chống trùng/lẫn bài giữa các sự kiện cùng nước, cùng khoảng thời gian ----
AMBIGUOUS_MARGIN = 1.5      # chênh điểm giữa sự kiện nhất và nhì < ngưỡng này -> ghi vào ambiguous_<NƯỚC>.csv để kiểm tay
OTHER_NAMED_PENALTY = 8.0   # bài nhắc tên bão của sự kiện khác -> các sự kiện không nhắc tên đó bị trừ điểm

COUNTRY_INFO = {
    "INDONESIA":   {"code": "IDN", "prefix": "ID", "name": "Indonesia"},
    "PHILIPPINES": {"code": "PHL", "prefix": "PH", "name": "Philippines"},
    "THAILAND":    {"code": "THA", "prefix": "TH", "name": "Thailand"},
    "MALAYSIA":    {"code": "MYS", "prefix": "MY", "name": "Malaysia"},
    "VIETNAM":     {"code": "VNM", "prefix": "VN", "name": "Vietnam"},
}

# Danh sách đài theo thứ tự cào của từng nước (đài chính trước). Cào HẾT tất cả.
OUTLETS = {
    "INDONESIA": [
        ("The Jakarta Post", "thejakartapost.com"),
        ("ANTARA News - English", "antaranews.com"),
        ("Tempo English", "en.tempo.co"),
        ("Jakarta Globe", "jakartaglobe.id"),
        ("RRI - Voice of Indonesia", "rri.co.id/voice-of-indonesia"),
        ("Indonesia Expat", "indonesiaexpat.id"),
        ("The Bali Times", "thebalitimes.com"),
        ("Inside Indonesia", "insideindonesia.org"),
        ("The Conversation Indonesia", "theconversation.com"),
        ("Kompas (English)", "kompas.id"),
    ],
    "PHILIPPINES": [
        ("Philippine Daily Inquirer", "inquirer.net"),
        ("GMA News Online", "gmanetwork.com"),
        ("ABS-CBN News", "abs-cbn.com"),
        ("Philstar.com", "philstar.com"),
        ("Rappler", "rappler.com"),
        ("Manila Bulletin", "mb.com.ph"),
        ("BusinessWorld", "bworldonline.com"),
        ("The Manila Times", "manilatimes.net"),
        ("Manila Standard", "manilastandard.net"),
    ],
    "THAILAND": [
        ("Bangkok Post", "bangkokpost.com"),
        ("The Nation Thailand", "nationthailand.com"),
        ("Thai PBS World", "thaipbsworld.com"),
        ("Khaosod English", "khaosodenglish.com"),
        ("Pattaya Mail", "pattayamail.com"),
        ("MCOT", "mcot.net"),
        ("NNT - National News Bureau of Thailand", "thainationalnews.com"),
        ("The Thaiger", "thethaiger.com"),
        ("Thai Enquirer", "thaienquirer.com"),
    ],
    "MALAYSIA": [
        ("Bernama", "bernama.com"),
        ("The Star", "thestar.com.my"),
        ("New Straits Times", "nst.com.my"),
        ("Malay Mail", "malaymail.com"),
        ("Free Malaysia Today (FMT)", "freemalaysiatoday.com"),
        ("The Borneo Post", "theborneopost.com"),
        ("The Edge Malaysia", "theedgemalaysia.com"),
        ("The Sun Daily", "thesun.my"),
        ("Daily Express", "dailyexpress.com.my"),
    ],
    "VIETNAM": [
        ("VietnamPlus", "en.vietnamplus.vn"),
        ("Vietnam News", "vietnamnews.vn"),
        ("VnExpress International", "e.vnexpress.net"),
        ("Tuoi Tre News", "tuoitrenews.vn"),
        ("VietNamNet Global", "vietnamnet.vn/en"),
        ("VOV English", "english.vov.vn"),
        ("Nhan Dan English", "en.nhandan.vn"),
        ("Vietnam Investment Review (VIR)", "vir.com.vn"),
        ("Saigon Times English", "english.thesaigontimes.vn"),
    ],
}

# Các domain chỉ nhận bài TIẾNG ANH (bài tiếng Indonesia bị loại sau khi tải, dựa vào nội dung bài).
ENGLISH_ONLY_SITES = {"theconversation.com", "kompas.id"}
MIN_EN_RATIO = 0.15         # tỉ lệ từ dừng tiếng Anh tối thiểu trong 400 từ đầu của bài

COLS = ["OUTLET", "Disaster Name", "Article ID", "Tittle", "Full text", "Date", "URL"]

EXCLUDED_DOMAINS = ["login.yahoo", "facebook.com", "youtube.com", "twitter.com",
                    "instagram.com", "accounts.google"]
# Trang không phải bài báo (thẻ, chuyên mục, tìm kiếm, phân trang...)
JUNK_URL_PATTERNS = re.compile(
    r"/(tag|tags|topic|topics|category|categories|search|author|page|video|photo)s?(/|\?|$)", re.I)