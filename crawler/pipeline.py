# -*- coding: utf-8 -*-
"""Quy trình cào cho 1 nước: duyệt thảm họa -> duyệt đài -> tìm link -> tải bài -> gán sự kiện -> checkpoint."""
import os
from collections import Counter
from concurrent.futures import ThreadPoolExecutor

import pandas as pd

from .config import (COUNTRY_INFO, OUTLETS, COLS, USE_WP_API, DOWNLOAD_WORKERS, AUTO_ARTICLE_ID,
                     ENGLISH_ONLY_SITES, MIN_EN_RATIO)
from .utils import _cell, _fold, is_english
from .vocabulary import HAZARD_SYNONYMS
from .search import norm_url, search_article_urls
from .fetch import extract_article_content, wp_available, wp_search, wp_to_content
from .disaster import disaster_window, date_ok, event_names
from .matching import build_profiles, pick_event, log_ambiguous
from .queries import build_queries


def get_outlets(country):
    """[(tên đài, site)] theo thứ tự trong OUTLETS, bỏ đài trùng domain."""
    outlets, used = [], set()
    for name, site in OUTLETS[country]:
        if site in used:
            continue
        used.add(site)
        outlets.append((name, site))
    return outlets


def load_country_disasters(df_all, code, years, limit):
    df = df_all.dropna(subset=["Disaster Name"]).copy()
    df["Disaster Name"] = df["Disaster Name"].astype(str).str.strip()
    df = df[df["Disaster Name"].str.endswith("-" + code)]
    if years:
        sy = pd.to_numeric(df["Start Year"], errors="coerce")
        df = df[(sy >= years[0]) & (sy <= years[1])]
    return df.head(limit) if limit else df


def process_country(country, df_all, outlets, years, limit):
    info = COUNTRY_INFO[country]
    ckpt = f"checkpoint_{country}.csv"
    combos_file = f"checkpoint_{country}_done.txt"

    profiles = build_profiles(load_country_disasters(df_all, info["code"], None, None))   # TẤT CẢ sự kiện của nước
    df = load_country_disasters(df_all, info["code"], years, limit)                       # sự kiện sẽ chạy
    print(f"\n📂 {country}: {len(df)} thảm họa | cào HẾT {len(outlets)} đài, không giới hạn số bài")
    print("   Thứ tự đài: " + " → ".join(n for n, _ in outlets))

    # ---- nạp checkpoint (chạy lại sẽ tiếp tục, không cào lại) ----
    rows = (pd.read_csv(ckpt, dtype=str, keep_default_na=False).to_dict("records")
            if os.path.exists(ckpt) else [])
    seen_urls = {norm_url(r["URL"]) for r in rows if r.get("URL")}
    seen_titles = {_fold(r["Tittle"]) for r in rows if r.get("Tittle") and r["Tittle"] != "No Title"}
    done_combos = {(r["Disaster Name"], r["OUTLET"]) for r in rows}
    if os.path.exists(combos_file):
        with open(combos_file, encoding="utf-8") as f:
            done_combos |= {tuple(l.rstrip("\n").split("\t")) for l in f if "\t" in l}

    if rows or done_combos:
        n_ev = len({r["Disaster Name"] for r in rows})
        print(f"   ♻️ Tiếp tục từ checkpoint: đã có {len(rows)} bài của {n_ev} thảm họa, "
              f"{len(done_combos)} cặp (thảm họa-đài) đã xong -> sẽ bỏ qua phần này.")

    def count(code):
        return sum(1 for r in rows if r["Disaster Name"] == code)

    def crawl_outlet(row, code, window, outlet, site):
        """Cào HẾT một đài. Mỗi bài tìm được sẽ được gán cho sự kiện PHÙ HỢP NHẤT trong nhóm
        sự kiện cùng nước/cùng thời gian (không nhất thiết là sự kiện đang chạy)."""
        if (code, outlet) in done_combos:
            return
        tried, st, net_fail = set(), Counter(), False

        def consider(url, c):
            if site in ENGLISH_ONLY_SITES and not is_english(c["full_text"], MIN_EN_RATIO):
                st["không phải tiếng Anh"] += 1             # đài chỉ nhận bài tiếng Anh
                return
            tkey = _fold(c["title"])
            if c["title"] != "No Title" and tkey in seen_titles:
                st["đã có/trùng"] += 1                        # bài đăng lại
                return
            best, amb = pick_event(c, code, profiles)
            if best is None:
                st["ngoài khoảng ngày" if not date_ok(c["date"], window) else "không liên quan"] += 1
                return
            seen_urls.add(norm_url(url))
            seen_titles.add(tkey)
            rows.append({"OUTLET": outlet, "Disaster Name": best, "Article ID": "",
                         "Tittle": c["title"], "Full text": c["full_text"],
                         "Date": c["date"], "URL": url})
            if amb:
                log_ambiguous(country, outlet, c, url, amb)
            if best == code:
                st["LẤY ĐƯỢC"] += 1
                print(f"      ✅ {c['title'][:70]}" + ("  ⚠️ khó phân biệt" if amb else ""))
            else:
                st["chuyển sang sự kiện khác"] += 1
                print(f"      ↪️ [{best}] {c['title'][:60]}")

        # ---- Nguồn 1: API WordPress (nếu đài dùng WordPress) ----
        if USE_WP_API and wp_available(site):
            dtype = _cell(row, "Disaster Type")
            terms = event_names(row)[:3] or HAZARD_SYNONYMS.get(dtype, [dtype.lower()])
            for term in terms:
                print(f"   🧩 [{outlet}] WordPress API: '{term}' trong khoảng ngày của thảm họa")
                for item in wp_search(site, term, window):
                    key = norm_url(item["url"])
                    st["link tìm thấy"] += 1
                    if key in seen_urls or key in tried:
                        st["đã có/trùng"] += 1
                        continue
                    tried.add(key)
                    c, why = wp_to_content(item)
                    if c:
                        consider(item["url"], c)
                    else:
                        st[why] += 1

        # ---- Nguồn 2: công cụ tìm kiếm (site:) ----
        for query in build_queries(row, site, info["name"]):
            print(f"   🔍 [{outlet}] {query}")
            urls = search_article_urls(query, site)
            if urls is None:                          # lỗi mạng: query này chưa được tìm
                net_fail = True
                continue
            st["link tìm thấy"] += len(urls)
            todo = []
            for url in urls:
                key = norm_url(url)
                if key in seen_urls or key in tried:
                    st["đã có/trùng"] += 1
                    continue
                tried.add(key)
                todo.append(url)
            if not todo:
                continue
            with ThreadPoolExecutor(max_workers=DOWNLOAD_WORKERS) as pool:   # tải song song
                fetched = list(pool.map(extract_article_content, todo))
            for url, (c, why) in zip(todo, fetched):
                if c:
                    consider(url, c)
                else:
                    st[why] += 1
        got = st["LẤY ĐƯỢC"]
        print(f"   📈 [{outlet}] " + (" | ".join(f"{k}: {v}" for k, v in st.items()) or "không có link nào"))
        # lưu bài trước, đánh dấu "xong đài" sau -> mất điện giữa chừng vẫn an toàn
        pd.DataFrame(rows, columns=COLS).to_csv(ckpt, index=False)
        if net_fail:                                  # có query bị lỗi mạng -> KHÔNG đánh dấu xong
            print(f"   → {outlet}: +{got} bài (⚠️ có query lỗi mạng, đài này sẽ được chạy lại ở lần sau)")
            return
        done_combos.add((code, outlet))
        with open(combos_file, "a", encoding="utf-8") as f:
            f.write(f"{code}\t{outlet}\n")
        print(f"   → {outlet}: +{got} bài")

    report = []

    def _run_event(row):
        code = row["Disaster Name"]
        window = disaster_window(row)
        n_rivals = len(profiles[code]["rivals"]) - 1
        print(f"\n📌 [{code}] {row.get('Disaster Type')} | {row.get('Disaster Subtype')} | "
              f"{row.get('Event Name')} | {row.get('Location')}"
              + (f"  (cạnh tranh với {n_rivals} sự kiện cùng kỳ)" if n_rivals else ""))

        for outlet, site in outlets:                       # cào HẾT mọi đài, không dừng sớm
            crawl_outlet(row, code, window, outlet, site)

        n = count(code)
        by_outlet = "; ".join(f"{o}: {sum(1 for r in rows if r['Disaster Name'] == code and r['OUTLET'] == o)}"
                              for o, _ in outlets
                              if any(r["Disaster Name"] == code and r["OUTLET"] == o for r in rows))
        print(f"   📊 {code}: {n} bài ({by_outlet})")
        report.append({"Disaster Name": code, "Type": row.get("Disaster Type"),
                       "Subtype": row.get("Disaster Subtype"), "Event Name": row.get("Event Name"),
                       "Articles": n, "Status": "KHÔNG CÓ BÀI" if n == 0 else "OK",
                       "By outlet": by_outlet})

    for _, row in df.iterrows():
        try:
            _run_event(row)
        except KeyboardInterrupt:
            print("\n⏹️  Đã dừng theo yêu cầu. Đang lưu phần đã cào... (chạy lại đúng lệnh này để làm tiếp)")
            pd.DataFrame(rows, columns=COLS).to_csv(ckpt, index=False)
            break

    pd.DataFrame(report).to_csv(f"report_{country}.csv", index=False, encoding="utf-8-sig")

    # ---- sắp xếp theo thứ tự thảm họa -> thứ tự đài, rồi đánh Article ID ----
    ev_order = {c: i for i, c in enumerate(df["Disaster Name"])}
    ou_order = {o: i for i, (o, _) in enumerate(outlets)}
    out = pd.DataFrame(rows, columns=COLS)
    out = out[out["Disaster Name"].isin(ev_order)]         # chỉ giữ thảm họa đang chạy (--year/--limit)
    out = (out.assign(_e=out["Disaster Name"].map(ev_order), _o=out["OUTLET"].map(ou_order))
              .sort_values(["_e", "_o"], kind="stable").drop(columns=["_e", "_o"]).reset_index(drop=True))
    if AUTO_ARTICLE_ID:
        out["Article ID"] = [f"{info['prefix']}{i:03d}" for i in range(1, len(out) + 1)]
    else:
        out["Article ID"] = ""                            # để trống, đánh tay sau
    return out