# News Crawler — Cào tin thiên tai Đông Nam Á

Công cụ đọc danh sách thảm họa (dữ liệu EM-DAT) của từng nước, tự tìm các bài báo tiếng Anh
trên những đài báo lớn của nước đó, tải nội dung bài, **gán mỗi bài cho đúng thảm họa** rồi ghi
vào form Excel `NEWS DATA`.

Hỗ trợ 5 nước: **INDONESIA, PHILIPPINES, THAILAND, MALAYSIA, VIETNAM** 

---

## 1. Cài đặt

```bash
python -m pip install pandas trafilatura openpyxl ddgs
# (nếu không cài được ddgs thì dùng bản cũ: pip install duckduckgo_search)
```

Cần Python 3.8 trở lên.

## 2. Chuẩn bị thư mục chạy

Đặt `main_crawler.py`, thư mục `crawler/` và 2 file Excel đầu vào **chung một thư mục**:

```
<thư mục chạy>/
├── main_crawler.py
├── crawler/                     ← phải nằm cạnh main_crawler.py
├── [8386] DISASTER DATA.xlsx
└── [8386] NEWS DATA.xlsx
```

| File | Vai trò |
|---|---|
| `[8386] DISASTER DATA.xlsx`, sheet `(ĐÃ LỌC) THEO QUỐC GIA` | Danh sách thảm họa. Các cột dùng: `Disaster Name` (mã, kết thúc bằng `-VNM`, `-IDN`…), `Disaster Type`, `Disaster Subtype`, `Event Name`, `Origin`, `Associated Types`, `Location`, `Start Year/Month/Day`, `End Year/Month/Day` |
| `[8386] NEWS DATA.xlsx` | File **mẫu**, có 1 sheet cho mỗi nước. Script chỉ đọc, **không ghi đè** file này |

## 3. Cách chạy

Mở terminal tại thư mục chạy:

```bash
python main_crawler.py VIETNAM                  # cào toàn bộ thảm họa của Việt Nam
python main_crawler.py VIETNAM --year 2016-2018 # chỉ thảm họa bắt đầu 2016–2018
python main_crawler.py VIETNAM --limit 3        # chạy thử 3 thảm họa đầu
python main_crawler.py VIETNAM --check          # kiểm tra nhanh các đài (không cào thật)
python main_crawler.py VIETNAM --reassign       # chấm lại bài trong checkpoint, chuyển về đúng sự kiện
python main_crawler.py --merge                  # gộp file của các nước thành 1 file
```

- Có thể truyền nhiều nước một lúc: `... main_crawler.py VIETNAM THAILAND`.
- **Dừng giữa chừng**: bấm `Ctrl+C`. Phần đã cào được lưu vào checkpoint. Chạy lại **đúng lệnh cũ** để làm tiếp.
- Sau khi `--reassign`, chạy lại lệnh cào bình thường để xuất lại file Excel.

## 4. File đầu ra (trong thư mục chạy)

| File | Nội dung |
|---|---|
| `[8386] NEWS DATA - <NƯỚC>.xlsx` | Kết quả cuối, đúng form mẫu: OUTLET, Disaster Name, Article ID, Tittle, Full text, Date, URL. Nếu file đang mở trong Excel, script lưu sang tên có thêm giờ, vd `... (153012).xlsx` |
| `[8386] NEWS DATA - MERGED.xlsx` | Gộp các nước (lệnh `--merge`) |
| `checkpoint_<NƯỚC>.csv` | Toàn bộ bài đã cào, dùng để chạy tiếp |
| `checkpoint_<NƯỚC>_done.txt` | Các cặp (thảm họa, đài) đã cào xong |
| `checkpoint_<NƯỚC>.csv.bak` | Bản sao lưu được tạo trước khi `--reassign` |
| `report_<NƯỚC>.csv` | Số bài của mỗi thảm họa, chia theo đài; ghi `KHÔNG CÓ BÀI` nếu = 0 |
| `ambiguous_<NƯỚC>.csv` | Những bài mà 2 thảm họa có điểm sát nhau, **cần kiểm tra tay** |


---

## 5. Cấu trúc thư mục

```
news_crawler/
├── main_crawler.py        # Điểm chạy (CLI): đọc tham số dòng lệnh, gọi lệnh tương ứng
├── README.md
└── crawler/
    ├── config.py          # ★ CẤU HÌNH: tên file, giới hạn, cờ bật/tắt, danh sách nước & đài báo
    ├── vocabulary.py      # Từ điển từ khóa/regex về thiên tai, tên tỉnh (chỉ là dữ liệu)
    ├── utils.py           # _num, _cell, _fold, is_english: đọc ô an toàn, bỏ dấu tiếng Việt, nhận diện bài tiếng Anh
    ├── search.py          # Tìm link bài báo qua DuckDuckGo/Bing/Brave/Google (thư viện ddgs)
    ├── fetch.py           # Tải trang, trích nội dung (trafilatura) + lấy bài qua WordPress API
    ├── disaster.py        # Đọc 1 dòng thảm họa: khoảng ngày, địa danh, tên bão, cụm từ khóa
    ├── queries.py         # Tạo các câu truy vấn tìm kiếm cho 1 thảm họa trên 1 đài
    ├── matching.py        # Chấm điểm bài báo và chọn thảm họa phù hợp nhất
    ├── excel_io.py        # Ghi dữ liệu vào form Excel mẫu (giữ nguyên font/màu)
    ├── pipeline.py        # Quy trình cào 1 nước (vòng lặp chính + checkpoint)
    └── commands.py        # 4 lệnh: run_country, reassign_country, merge_outputs, check_country
```

Quan hệ phụ thuộc giữa các module (mũi tên nghĩa là "dùng"):

```
main_crawler.py ─► commands ─► pipeline ─► queries ─► disaster ─► utils, vocabulary, config
                       │           │        matching ─► disaster
                       │           ├──────► search, fetch
                       └─► excel_io
```

Đọc code theo thứ tự này sẽ dễ hiểu nhất: `config` → `disaster` → `matching` → `queries` → `pipeline` → `commands`.

---

## 6. Script hoạt động thế nào

### 6.1 Luồng chính (`pipeline.process_country`)

```
Với mỗi thảm họa của nước (lọc theo --year / --limit):
  Với mỗi đài báo trong OUTLETS[nước]  (cào HẾT mọi đài, không dừng sớm):
     Nếu cặp (thảm họa, đài) đã xong trong checkpoint → bỏ qua
     Nguồn 1: nếu đài chạy WordPress → gọi /wp-json/wp/v2/posts?search=... lọc sẵn theo ngày
     Nguồn 2: build_queries() → tìm "… site:<đài>" → lọc link → tải 6 bài song song
     Với mỗi bài tải được:
        - bỏ nếu đài thuộc ENGLISH_ONLY_SITES mà bài không phải tiếng Anh (is_english)
        - bỏ nếu trùng URL hoặc trùng tiêu đề (bài đăng lại)
        - pick_event(): chọn thảm họa phù hợp nhất (có thể là thảm họa KHÁC thảm họa đang chạy)
        - lưu vào rows; nếu điểm sát nhau → ghi ambiguous_<NƯỚC>.csv
     Lưu checkpoint, rồi mới đánh dấu "đài xong" (mất điện giữa chừng vẫn an toàn)
Cuối cùng: ghi report, sắp xếp theo thứ tự thảm họa → thứ tự đài, ghi vào form Excel
```

### 6.2 Một bài được giữ lại khi nào? (`matching.is_relevant` + `disaster.date_ok`)

1. **Ngày đăng** nằm trong khoảng `[ngày bắt đầu − 7, ngày kết thúc + 45]` (`DAYS_BEFORE`, `DAYS_AFTER`). Bài không có ngày vẫn được giữ để kiểm tra tay.
2. **Bài dài** ít nhất `MIN_WORDS` = 100 từ.
3. **Nội dung** phải thỏa **một trong hai nhánh** (`is_relevant`):
   - **(a) Tên bão**: thảm họa có tên bão (vd `Typhoon 'Damrey' / 'Ramil'`) và bài nhắc tên bão **cùng** có từ chỉ thiên tai.
   - **(b) Đúng loại + địa danh**: bài nói đúng **loại** thiên tai, và nhắc ít nhất 1 địa danh trong cột `Location` (hoặc tiêu đề nói về loại thiên tai đó).
4. Nhánh (b) áp dụng cho **cả thảm họa có tên bão**, nên bài chỉ viết "the typhoon" hoặc "Typhoon hits Luzon" vẫn được giữ. Bài chỉ có tên bão mà thiếu địa danh (vd "Alpha weakens") vẫn được giữ nhờ nhánh (a). Muốn quay về luật cũ (bão có tên bắt buộc nhắc tên) thì đặt `NAMED_EVENT_STRICT = True`.
5. Với các đài trong `ENGLISH_ONLY_SITES` (Indonesia: `theconversation.com`, `kompas.id`, có đăng cả bài tiếng Indonesia): bài phải qua `is_english` mới được xét tiếp. Hàm này đếm từ dừng tiếng Anh và tiếng Indonesia trong 400 từ đầu của bài.

### 6.3 Phân biệt các thảm họa cùng nước, cùng thời gian (`matching.score_article`)

Thảm họa nào có khoảng ngày chồng lên nhau được xem là "đối thủ" (`rivals`). Mỗi bài được chấm điểm
với từng đối thủ, và **thảm họa có điểm cao nhất nhận bài**:

| Tín hiệu | Điểm |
|---|---|
| Tên bão xuất hiện trong tiêu đề / trong thân bài | +12 / +10 |
| Đúng loại (`Disaster Type`) | +2 |
| Khớp `Disaster Subtype` | +3 mỗi cụm (tối đa 2 cụm) |
| Khớp `Origin` | +2 mỗi cụm (tối đa 2 cụm) |
| Khớp `Associated Types` | +1.5 mỗi cụm (tối đa 3 cụm) |
| Địa danh trong `Location` | +1 mỗi địa danh (tối đa 4); +1.5 nếu có trong tiêu đề |
| Ngày đăng | 0–4 (4 nếu nằm trong thời gian thảm họa, giảm dần về 0 khi cách xa 60 ngày) |
| Bài nhắc tên bão của thảm họa **khác** | −8 (`OTHER_NAMED_PENALTY`) |

Nếu điểm nhất và nhì chênh nhau dưới `AMBIGUOUS_MARGIN` = 1.5, bài vẫn được gán cho thảm họa
nhất nhưng cũng được ghi vào `ambiguous_<NƯỚC>.csv`.

Cột EM-DAT (`Subtype`, `Origin`, `Associated Types`) được đổi thành từ khóa báo chí qua bảng
`HAZARD_KW` trong `vocabulary.py`. Ví dụ `Flash flood` → `flash flood, flash floods, flash-flood…`.

### 6.4 Truy vấn tìm kiếm (`queries.build_queries`)

Thứ tự ưu tiên: tên bão → từng địa danh (tối đa `MAX_PLACES` = 3) → Subtype → cả nước → Origin →
Associated Types → từ đồng nghĩa. Tất cả đều kèm tháng/năm và `site:<đài>`.
Với Indonesia, tên tỉnh được dịch sang cách báo tiếng Anh thường viết (`Jawa Barat` → `West Java`).

---

## 7. Những thay đổi hay gặp

| Muốn… | Sửa ở đâu |
|---|---|
| Thêm/bớt/đổi thứ tự đài báo | `OUTLETS` trong `crawler/config.py` (đài đứng trước được cào trước) |
| Thêm nước mới | `COUNTRY_INFO` + `OUTLETS` trong `config.py`, và thêm sheet cùng tên vào file mẫu |
| Nới/siết khoảng ngày | `DAYS_BEFORE`, `DAYS_AFTER` trong `config.py` |
| Tự đánh Article ID (`VN001`…) | `AUTO_ARTICLE_ID = True` |
| Bị công cụ tìm kiếm chặn | tăng `SLEEP_RANGE`, đổi `SEARCH_BACKENDS` (biến này nằm trong `crawler/search.py`) |
| Thêm từ khóa thiên tai | `HAZARD_KW`, `TYPE_RE`, `HAZARD_SYNONYMS` trong `crawler/vocabulary.py` |
| Bỏ yêu cầu phải nhắc địa danh (nhánh đúng loại) | `STRICT_PLACE` trong `config.py` |
| Bắt buộc bão có tên phải nhắc tên (luật cũ) | `NAMED_EVENT_STRICT = True` trong `config.py` (mặc định `False`: nhận bài nếu nhắc tên bão HOẶC đúng loại + địa danh) |
| Đài chỉ nhận bài tiếng Anh | thêm domain vào `ENGLISH_ONLY_SITES`; chỉnh ngưỡng bằng `MIN_EN_RATIO` (mặc định 0.15) trong `config.py` |

Khi một đài báo cho kết quả lạ, chạy `--check` trước. Kết quả hiện theo các cột sau:
- `LINK = 0`: công cụ tìm kiếm không trả kết quả.
- `tải lỗi`: báo chặn bot.
- `quá ngắn`: thường do paywall.
- `WP-API có`: đài lấy bài được qua API, rất nhanh.