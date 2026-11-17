import re
import time
from io import BytesIO
from urllib.parse import urlparse, urlunparse, parse_qs, urlencode

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

st.set_page_config(page_title="Screener.in to Excel", page_icon="📈", layout="wide")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}


def clean_base_url(url: str) -> str:
    """Remove any existing page= parameter, keep the rest of the URL."""
    p = urlparse(url.strip())
    q = parse_qs(p.query)
    q.pop("page", None)
    return urlunparse((p.scheme, p.netloc, p.path, "", urlencode(q, doseq=True), ""))


def page_url(base: str, page: int, limit: int | None) -> str:
    p = urlparse(base)
    q = parse_qs(p.query)
    if limit:
        q["limit"] = [str(limit)]
    if page > 1:
        q["page"] = [str(page)]
    return urlunparse((p.scheme, p.netloc, p.path, "", urlencode(q, doseq=True), ""))


def to_number(v):
    if v is None:
        return None
    s = str(v).replace(",", "").replace("%", "").strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return v


def parse_page(html: str):
    """Return (columns, rows, total_pages) from one screener page."""
    soup = BeautifulSoup(html, "html.parser")

    total_pages = None
    pag = soup.select_one("div.pagination, .pagination")
    if pag:
        m = re.search(r"of\s+(\d+)", pag.get_text(" ", strip=True))
        if m:
            total_pages = int(m.group(1))

    table = soup.select_one("table.data-table")
    if table is None:
        return [], [], total_pages

    columns = None
    rows = []
    for tr in table.find_all("tr"):
        ths = tr.find_all("th")
        tds = tr.find_all("td")
        if ths and not tds:
            if columns is None:
                columns = [re.sub(r"\s+", " ", th.get_text(" ", strip=True)) for th in ths]
            continue  # repeated header rows are skipped
        if not tds:
            continue
        values = [re.sub(r"\s+", " ", td.get_text(" ", strip=True)) for td in tds]
        link = tds[1].find("a") if len(tds) > 1 else None
        href = ("https://www.screener.in" + link["href"]) if link and link.get("href") else ""
        rows.append(values + [href])

    if columns is None:
        return [], [], total_pages
    columns = columns + ["Link"]
    rows = [r for r in rows if len(r) == len(columns)]
    return columns, rows, total_pages


def scrape(base_url, max_pages, limit, session_id, delay, progress, status):
    sess = requests.Session()
    sess.headers.update(HEADERS)
    if session_id:
        sess.cookies.set("sessionid", session_id.strip(), domain="www.screener.in")

    all_frames = []
    columns = None
    seen_first = set()
    total_pages = None

    for page in range(1, max_pages + 1):
        url = page_url(base_url, page, limit)
        status.write(f"Fetching page {page}: {url}")
        try:
            r = sess.get(url, timeout=30)
        except requests.RequestException as e:
            st.warning(f"Page {page}: request failed ({e}). Stopping.")
            break

        if r.status_code == 429:
            st.warning("Screener is rate-limiting (429). Increase the delay and retry.")
            break
        if r.status_code != 200:
            st.warning(f"Page {page}: HTTP {r.status_code}. Stopping.")
            break

        cols, rows, tp = parse_page(r.text)
        if tp and not total_pages:
            total_pages = tp
        if not rows:
            status.write(f"Page {page} has no rows. Finished.")
            break

        first_key = tuple(rows[0][:3])
        if first_key in seen_first:  # page repeated -> past the last page
            break
        seen_first.add(first_key)

        if columns is None:
            columns = cols
        df = pd.DataFrame(rows, columns=columns)
        df.insert(0, "Page", page)
        all_frames.append(df)

        last = min(max_pages, total_pages) if total_pages else max_pages
        progress.progress(min(page / last, 1.0))

        if total_pages and page >= total_pages:
            break
        time.sleep(delay)

    if not all_frames:
        return pd.DataFrame()

    data = pd.concat(all_frames, ignore_index=True)
    for c in data.columns:
        if c not in ("Page", "Company", "Link"):
            data[c] = data[c].map(to_number)
    # Re-number S.No. continuously
    if "S.No." in data.columns:
        data["S.No."] = range(1, len(data) + 1)
    return data


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Screener Data")
        ws = writer.sheets["Screener Data"]
        ws.freeze_panes = "A2"
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
            ws.column_dimensions[col[0].column_letter].width = min(width + 3, 40)
    return buf.getvalue()


# ---------------------------- UI ----------------------------
st.title("📈 Screener.in → Excel")
st.caption("All pages of a screen are combined into one sheet.")

url = st.text_input(
    "Screener URL (paste the page-1 URL; page=2,3,... is added automatically)",
    value="https://www.screener.in/screens/4101182/up-from-52w-low-15/",
)

c1, c2, c3 = st.columns(3)
max_pages = c1.number_input("Max pages", min_value=1, max_value=500, value=100)
use_limit = c2.checkbox("Ask for 50 rows per page (limit=50)", value=True)
delay = c3.number_input("Delay between pages (sec)", min_value=0.0, max_value=10.0, value=1.0, step=0.5)

session_id = st.text_input(
    "Screener 'sessionid' cookie (optional, needed for private screens / extra columns)",
    type="password",
    help="Chrome → F12 → Application → Cookies → www.screener.in → sessionid",
)

if st.button("Create Excel", type="primary"):
    if not url.strip():
        st.error("Please enter a URL.")
    else:
        base = clean_base_url(url)
        progress = st.progress(0.0)
        status = st.empty()
        with st.spinner("Scraping..."):
            df = scrape(base, int(max_pages), 50 if use_limit else None, session_id, delay, progress, status)
        if df.empty:
            st.error("No data found. If the screen is private, add your sessionid cookie.")
        else:
            progress.progress(1.0)
            st.success(f"Done: {len(df)} shares from {df['Page'].nunique()} page(s).")
            st.dataframe(df, use_container_width=True)
            st.download_button(
                "⬇️ Download Excel",
                data=to_excel_bytes(df),
                file_name="screener_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
