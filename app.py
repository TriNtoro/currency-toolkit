import calendar
import io
import re
import xml.etree.ElementTree as ET
from datetime import date

import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ============================================================
# KONFIGURASI
# ============================================================

BI_ENDPOINT = (
    "https://www.bi.go.id/biwebservice/wskursbi.asmx/getSubKursLokal3"
)

REQUEST_TIMEOUT = 45

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0 Safari/537.36"
    ),
    "Accept": "text/xml, application/xml, */*;q=0.8",
    "Accept-Language": "id-ID,id;q=0.9,en;q=0.8",
}

BULAN_ID = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]


# ============================================================
# AMBIL DATA DARI BI
# ============================================================

def fetch_bi(currency, start_date, end_date):
    """Ambil kurs dari web service BI. Coba GET lalu POST."""

    params = {
        "mts": currency,
        "startdate": start_date.strftime("%Y-%m-%d"),
        "enddate": end_date.strftime("%Y-%m-%d"),
    }

    last_error = None

    for method in ("GET", "POST"):
        try:
            if method == "GET":
                resp = requests.get(
                    BI_ENDPOINT, params=params,
                    headers=HEADERS, timeout=REQUEST_TIMEOUT,
                )
            else:
                resp = requests.post(
                    BI_ENDPOINT, data=params,
                    headers=HEADERS, timeout=REQUEST_TIMEOUT,
                )

            resp.raise_for_status()
            rows = parse_xml(resp.content)

            if rows:
                return rows

            last_error = (
                f"{method}: BI merespons tetapi tidak ada "
                f"baris data untuk periode ini."
            )

        except Exception as exc:
            last_error = f"{method}: {type(exc).__name__}: {exc}"

    raise RuntimeError(
        last_error or "Gagal mengambil data dari Bank Indonesia."
    )


def parse_xml(content):
    """Tag BI: jual_subkurslokal, beli_subkurslokal, tgl_subkurslokal."""

    def local(tag):
        return tag.rsplit("}", 1)[-1].lower()

    root = ET.fromstring(content)
    rows = []

    for table in root.iter():
        if local(table.tag) != "table":
            continue

        sell = buy = tgl = None
        for child in table:
            name = local(child.tag)
            val = (child.text or "").strip()
            if name == "jual_subkurslokal":
                sell = to_number(val)
            elif name == "beli_subkurslokal":
                buy = to_number(val)
            elif name == "tgl_subkurslokal":
                tgl = to_date(val)

        if tgl and sell is not None and buy is not None:
            rows.append({"date": tgl, "sell": sell, "buy": buy})

    return rows


def to_number(s):
    if not s:
        return None
    s = re.sub(r"[^0-9,.\-]", "", s)
    if not s:
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return None


def to_date(s):
    if not s:
        return None
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", s)
    if m:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
    return None


# ============================================================
# TABEL HARIAN (hari libur tetap muncul sebagai baris kosong)
# ============================================================

def build_table(rows, start_date, end_date):
    by_date = {r["date"]: r for r in rows}
    records = []
    cur = start_date

    while cur <= end_date:
        r = by_date.get(cur)
        if r:
            sell, buy = r["sell"], r["buy"]
            avg = round((sell + buy) / 2, 2)
        else:
            sell = buy = avg = None

        records.append(
            {"Tanggal": cur, "Sell": sell, "Buy": buy, "Average": avg}
        )
        cur = date.fromordinal(cur.toordinal() + 1)

    return pd.DataFrame(records)


# ============================================================
# EXCEL (satu bulan per file)
# ============================================================

def _tulis_sheet(ws, currency, table, year, month):
    """Isi satu worksheet dengan tabel satu bulan."""
    ws.title = f"{BULAN_ID[month - 1][:3]} {year}"

    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="B0B0B0")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:D1")
    c = ws["A1"]
    c.value = f"Kurs Transaksi {currency} - {BULAN_ID[month - 1]} {year}"
    c.font = Font(bold=True, size=13)
    c.alignment = center

    for col, h in enumerate(["Tanggal", "Sell", "Buy", "Average"], start=1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    r = 4
    for _, row in table.iterrows():
        ws.cell(row=r, column=1, value=row["Tanggal"].strftime("%d/%m/%Y"))
        for col, key in enumerate(["Sell", "Buy", "Average"], start=2):
            cell = ws.cell(row=r, column=col, value=row[key])
            cell.number_format = "#,##0.00"
            cell.alignment = center
        for col in range(1, 5):
            ws.cell(row=r, column=col).border = border
        r += 1

    for col in range(1, 5):
        ws.column_dimensions[get_column_letter(col)].width = 16


def make_excel(currency, table, year, month):
    """Satu file, satu bulan."""
    wb = Workbook()
    _tulis_sheet(wb.active, currency, table, year, month)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def make_excel_tahun(currency, table, year):
    """Satu file, 12 sheet (satu sheet per bulan)."""
    wb = Workbook()
    first = True
    for m in range(1, 13):
        month_table = table[
            table["Tanggal"].apply(lambda d: d.month == m)
        ].reset_index(drop=True)

        if first:
            ws = wb.active
            first = False
        else:
            ws = wb.create_sheet()
        _tulis_sheet(ws, currency, month_table, year, m)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ============================================================
# UI
# ============================================================

st.set_page_config(
    page_title="Currency Download Toolkit",
    page_icon="💱",
    layout="wide",
)

st.title("💱 Currency Download Toolkit")
st.caption(
    "Kurs Transaksi Bank Indonesia · USD & JPY · oleh Tri Noviantoro"
)

with st.sidebar:
    st.header("Pengaturan")

    currency = st.radio("Mata Uang", ["USD", "JPY"], horizontal=True)
    mode = st.radio("Periode", ["Satu Bulan", "Satu Tahun"])

    this_year = date.today().year
    year = st.selectbox("Tahun", list(range(this_year, this_year - 8, -1)))

    if mode == "Satu Bulan":
        month = st.selectbox(
            "Bulan",
            list(range(1, 13)),
            format_func=lambda m: BULAN_ID[m - 1],
            index=date.today().month - 1,
        )
    else:
        month = None

    st.divider()
    fetch = st.button(
        "🔄 Ambil Data", type="primary", use_container_width=True
    )
    st.caption("JPY ditampilkan per 100 JPY.")
    st.caption("Sumber: Bank Indonesia")


if fetch:
    if mode == "Satu Bulan":
        start_date = date(year, month, 1)
        end_date = date(year, month, calendar.monthrange(year, month)[1])
    else:
        start_date = date(year, 1, 1)
        end_date = date(year, 12, 31)

    with st.spinner(f"Mengambil {currency} dari Bank Indonesia..."):
        try:
            rows = fetch_bi(currency, start_date, end_date)
            table = build_table(rows, start_date, end_date)
            st.session_state["result"] = {
                "currency": currency,
                "mode": mode,
                "year": year,
                "month": month,
                "table": table,
            }
        except Exception as exc:
            st.error("Gagal mengambil data dari Bank Indonesia.")
            with st.expander("Detail error"):
                st.code(str(exc))
            st.stop()


if "result" not in st.session_state:
    st.info("Atur mata uang & periode di kiri, lalu klik 🔄 Ambil Data.")
    st.stop()

res = st.session_state["result"]
currency = res["currency"]
mode = res["mode"]
year = res["year"]
table = res["table"]

filled = table.dropna(subset=["Average"])
if filled.empty:
    st.warning("Tidak ada data kurs pada periode ini.")
    st.stop()

def fmt_id(x, decimals):
    """Format angka gaya Indonesia: titik ribuan, koma desimal."""
    s = f"{x:,.{decimals}f}"            # 17,743.00 (gaya en)
    return s.replace(",", "·").replace(".", ",").replace("·", ".")


mean_sell = filled["Sell"].mean()
mean_buy = filled["Buy"].mean()

if currency == "JPY":
    m_sell = fmt_id(mean_sell / 100, 2)
    m_buy = fmt_id(mean_buy / 100, 2)
    satuan = " (per 1 JPY, dibagi 100)"
else:
    m_sell = fmt_id(mean_sell, 0)
    m_buy = fmt_id(mean_buy, 0)
    satuan = ""

c1, c2, c3 = st.columns(3)
c1.metric(f"Rata-rata Sell{satuan}", m_sell)
c2.metric("Rata-rata Buy", m_buy)
c3.metric("Jumlah Hari Data", f"{len(filled)}")

# ---- Rata-rata per bulan (dihitung sekali, ditampilkan di atas) ----
monthly = filled.copy()
monthly["Bulan"] = monthly["Tanggal"].apply(lambda d: (d.year, d.month))

agg = (
    monthly.groupby("Bulan")
    .agg(
        Sell=("Sell", "mean"),
        Buy=("Buy", "mean"),
        Average=("Average", "mean"),
        Hari=("Average", "size"),
    )
    .reset_index()
    .sort_values("Bulan")
)
agg["Bulan"] = agg["Bulan"].apply(lambda ym: f"{BULAN_ID[ym[1] - 1]} {ym[0]}")

# Format khusus tabel rata-rata bulanan:
#   USD -> bilangan bulat, titik ribuan (mis. 17.743)
#   JPY -> dibagi 100 lalu 2 desimal (mis. 113,27)
if currency == "JPY":
    for col in ["Sell", "Buy", "Average"]:
        agg[col] = (agg[col] / 100).round(2)
    decimals = 2
    judul = "📊 Rata-rata per Bulan (JPY per 1 yen, dibagi 100)"
else:
    for col in ["Sell", "Buy", "Average"]:
        agg[col] = agg[col].round(0)
    decimals = 0
    judul = "📊 Rata-rata per Bulan"

# Versi tampilan: teks dengan titik ribuan & koma desimal (gaya Indonesia)
agg_show = agg.copy()
for col in ["Sell", "Buy", "Average"]:
    agg_show[col] = agg_show[col].apply(lambda v: fmt_id(v, decimals))

st.subheader(judul)
st.dataframe(
    agg_show,
    use_container_width=True,
    hide_index=True,
)

# Download / copy data tabel rata-rata bulanan
csv_bulanan = agg_show.to_csv(index=False).encode("utf-8")
st.download_button(
    "⬇️ Download Rata-rata per Bulan (CSV)",
    data=csv_bulanan,
    file_name=f"Rata2_Bulanan_{currency}_{year}.csv",
    mime="text/csv",
)
with st.expander("📋 Salin data (copy-paste)"):
    st.code(agg_show.to_csv(index=False), language="text")

label = (
    f"{BULAN_ID[res['month'] - 1]} {year}"
    if mode == "Satu Bulan"
    else f"Tahun {year}"
)
st.subheader(f"Kurs {currency} · {label}")

show = table.copy()
show["Tanggal"] = show["Tanggal"].apply(lambda d: d.strftime("%d/%m/%Y"))
st.dataframe(show, use_container_width=True, hide_index=True, height=520)


st.subheader("📥 Download Excel")

if mode == "Satu Bulan":
    excel = make_excel(currency, table, year, res["month"])
    st.download_button(
        "Download Excel",
        data=excel,
        file_name=f"Kurs_{currency}_{year}_{res['month']:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
else:
    # --- Download setahun penuh (1 file, 12 sheet) ---
    excel_tahun = make_excel_tahun(currency, table, year)
    st.download_button(
        f"⬇️ Download Setahun {year} (1 file, 12 sheet)",
        data=excel_tahun,
        file_name=f"Kurs_{currency}_{year}_SETAHUN.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    # --- Atau download per bulan ---
    st.caption("Atau unduh satu bulan saja:")
    avail = sorted({d.month for d in table["Tanggal"]})
    dl_month = st.selectbox(
        "Pilih bulan untuk diunduh",
        avail,
        format_func=lambda m: BULAN_ID[m - 1],
    )
    month_table = table[
        table["Tanggal"].apply(lambda d: d.month == dl_month)
    ].reset_index(drop=True)
    excel = make_excel(currency, month_table, year, dl_month)
    st.download_button(
        "Download Bulan Ini",
        data=excel,
        file_name=f"Kurs_{currency}_{year}_{dl_month:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
