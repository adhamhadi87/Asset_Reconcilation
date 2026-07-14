import io
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Dashboard Aset SAP vs eAsset",
    page_icon="🏢",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    [data-testid="stMetric"] {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 14px;
        padding: 14px 16px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.05);
    }
    .title-card {
        padding: 1.25rem 1.5rem;
        border-radius: 18px;
        background: linear-gradient(120deg, #064e3b, #0f766e);
        color: white;
        margin-bottom: 1rem;
        box-shadow: 0 10px 28px rgba(6, 78, 59, .18);
    }
    .title-card h1 {margin: 0; font-size: 2rem;}
    .title-card p {margin: .35rem 0 0 0; opacity: .9;}
    .status-ok {color:#047857;font-weight:700;}
    .status-warn {color:#b45309;font-weight:700;}
    .status-bad {color:#b91c1c;font-weight:700;}
    </style>
    """,
    unsafe_allow_html=True,
)

DEFAULT_FILE = Path(__file__).with_name("Asset.xlsx")


def clean_asset_no(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    )


def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype("string").str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0.0)


@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    required = {"SAP", "EAsset", "DIM Eva grp 1"}
    missing = required.difference(xls.sheet_names)
    if missing:
        raise ValueError(f"Sheet tidak ditemui: {', '.join(sorted(missing))}")

    sap = pd.read_excel(xls, sheet_name="SAP", dtype=str)
    easset = pd.read_excel(xls, sheet_name="EAsset", dtype=str)
    dim = pd.read_excel(xls, sheet_name="DIM Eva grp 1", dtype=str)

    sap = sap.rename(
        columns={
            "Eval. Grp 1-Block": "Eval Group",
            "Asset": "No Aset SAP",
            "Sub-number": "Sub No",
            "Asset Description": "Deskripsi Ringkas",
            "Asset Description.1": "Deskripsi Aset",
            "Acquis.val.": "Nilai Perolehan SAP",
            "Book val.": "Nilai Buku SAP",
        }
    )
    easset = easset.rename(
        columns={
            "No. Aset SAP": "No Aset SAP",
            "No. Siri Pendaftaran": "No Siri Pendaftaran",
            "Harga (RM)": "Nilai eAsset",
            "Eval Group 1": "Eval Group",
        }
    )
    dim = dim.rename(columns={"Eval. Grp 1-Block": "Eval Group"})

    sap["No Aset SAP"] = clean_asset_no(sap["No Aset SAP"])
    easset["No Aset SAP"] = clean_asset_no(easset["No Aset SAP"])
    sap = sap[sap["No Aset SAP"].notna()].copy()
    easset = easset[easset["No Aset SAP"].notna()].copy()

    sap["Nilai Perolehan SAP"] = to_number(sap["Nilai Perolehan SAP"])
    sap["Nilai Buku SAP"] = to_number(sap["Nilai Buku SAP"])
    easset["Nilai eAsset"] = to_number(easset["Nilai eAsset"])
    easset["Tarikh Beli"] = pd.to_datetime(easset.get("Tarikh Beli"), errors="coerce")
    easset["Tarikh Invois"] = pd.to_datetime(easset.get("Tarikh Invois"), errors="coerce")

    # Nilai perolehan SAP bagi baris subtotal tanpa nombor aset telah dibuang.
    # Data diagregat mengikut nombor aset supaya sub-number/duplicate dapat dibandingkan secara adil.
    sap_agg = (
        sap.groupby("No Aset SAP", as_index=False)
        .agg(
            {
                "Eval Group": "first",
                "Deskripsi Ringkas": lambda x: " | ".join(x.dropna().astype(str).unique()[:3]),
                "Deskripsi Aset": lambda x: " | ".join(x.dropna().astype(str).unique()[:3]),
                "Nilai Perolehan SAP": "sum",
                "Nilai Buku SAP": "sum",
                "Sub No": "count",
            }
        )
        .rename(columns={"Sub No": "Bil Rekod SAP"})
    )

    easset_agg = (
        easset.groupby("No Aset SAP", as_index=False)
        .agg(
            {
                "No Siri Pendaftaran": lambda x: " | ".join(x.dropna().astype(str).unique()[:5]),
                "Jenis": lambda x: " | ".join(x.dropna().astype(str).unique()[:3]),
                "Jenama": lambda x: " | ".join(x.dropna().astype(str).unique()[:3]),
                "Lokasi": lambda x: " | ".join(x.dropna().astype(str).unique()[:3]),
                "Pegawai Penempatan": lambda x: " | ".join(x.dropna().astype(str).unique()[:3]),
                "No. Pesanan": lambda x: " | ".join(x.dropna().astype(str).unique()[:3]),
                "Tarikh Beli": "min",
                "Tarikh Invois": "min",
                "Nilai eAsset": "sum",
                "Eval Group": "first",
                "Bil.": "count",
            }
        )
        .rename(columns={"Bil.": "Bil Rekod eAsset", "Eval Group": "Eval Group eAsset"})
    )

    compare = sap_agg.merge(easset_agg, on="No Aset SAP", how="outer", indicator=True)
    compare["Sumber Rekod"] = compare["_merge"].map(
        {"both": "SAP & eAsset", "left_only": "SAP Sahaja", "right_only": "eAsset Sahaja"}
    )
    compare.drop(columns="_merge", inplace=True)

    # Kekalkan Eval Group daripada kedua-dua sistem secara berasingan.
    compare = compare.rename(columns={"Eval Group": "Eval Group SAP"})

    dim_map = dim[["Eval Group", "Detail", "Detail 2", "Detail 3"]].drop_duplicates("Eval Group")
    sap_dim = dim_map.rename(
        columns={
            "Eval Group": "Eval Group SAP",
            "Detail": "Klasifikasi SAP",
            "Detail 2": "Zon SAP",
            "Detail 3": "PTJ SAP",
        }
    )
    ea_dim = dim_map.rename(
        columns={
            "Eval Group": "Eval Group eAsset",
            "Detail": "Klasifikasi eAsset",
            "Detail 2": "Zon eAsset",
            "Detail 3": "PTJ eAsset",
        }
    )
    compare = compare.merge(sap_dim, on="Eval Group SAP", how="left")
    compare = compare.merge(ea_dim, on="Eval Group eAsset", how="left")

    # Kolum gabungan digunakan untuk paparan/filter PTJ sahaja.
    compare["Eval Group"] = compare["Eval Group SAP"].fillna(compare["Eval Group eAsset"])
    compare["Detail"] = compare["Klasifikasi SAP"].fillna(compare["Klasifikasi eAsset"])
    compare["Detail 3"] = compare["PTJ SAP"].fillna(compare["PTJ eAsset"])

    both = compare["Sumber Rekod"].eq("SAP & eAsset")
    valid_ptj = compare["PTJ SAP"].notna() & compare["PTJ eAsset"].notna()
    valid_eval = compare["Eval Group SAP"].notna() & compare["Eval Group eAsset"].notna()

    compare["Berlainan Lokasi"] = both & valid_ptj & (
        compare["PTJ SAP"].astype("string").str.strip()
        != compare["PTJ eAsset"].astype("string").str.strip()
    )
    compare["Salah Klasifikasi"] = both & valid_eval & (
        compare["Eval Group SAP"].astype("string").str.strip()
        != compare["Eval Group eAsset"].astype("string").str.strip()
    )

    compare["Perbezaan Nilai"] = compare["Nilai Perolehan SAP"].fillna(0) - compare["Nilai eAsset"].fillna(0)
    compare["Perbezaan Mutlak"] = compare["Perbezaan Nilai"].abs()

    tolerance = 1.00
    compare["Status Padanan"] = np.select(
        [
            compare["Sumber Rekod"].eq("SAP Sahaja"),
            compare["Sumber Rekod"].eq("eAsset Sahaja"),
            compare["Perbezaan Mutlak"].le(tolerance),
        ],
        ["Tiada dalam eAsset", "Tiada dalam SAP", "Padan"],
        default="Nilai Tidak Padan",
    )

    # Tahun aset untuk analisis umur rekod eAsset.
    compare["Tahun Beli"] = compare["Tarikh Beli"].dt.year.astype("Int64")

    # Detail berlainan lokasi dikira pada peringkat rekod/baris asal.
    # Setiap rekod SAP dipadankan dengan setiap rekod eAsset yang mempunyai No. Aset sama.
    # Kaedah ini menghasilkan 60 rekod bagi keseluruhan fail semasa.
    sap_location = sap[[
        "No Aset SAP", "Eval Group", "Deskripsi Aset", "Deskripsi Ringkas"
    ]].copy()
    sap_location = sap_location.rename(columns={"Eval Group": "Eval Group SAP"})

    easset_location_cols = [
        "No Aset SAP", "Eval Group", "No Siri Pendaftaran", "Jenis",
        "Jenama", "Lokasi", "Pegawai Penempatan"
    ]
    easset_location = easset[easset_location_cols].copy()
    easset_location = easset_location.rename(columns={"Eval Group": "Eval Group eAsset"})

    location_detail = sap_location.merge(
        easset_location, on="No Aset SAP", how="inner"
    )
    location_detail["Eval Group SAP"] = (
        location_detail["Eval Group SAP"].astype("string").str.strip()
    )
    location_detail["Eval Group eAsset"] = (
        location_detail["Eval Group eAsset"].astype("string").str.strip()
    )
    location_detail = location_detail[
        location_detail["Eval Group SAP"].notna()
        & location_detail["Eval Group eAsset"].notna()
        & location_detail["Eval Group SAP"].ne(location_detail["Eval Group eAsset"])
    ].copy()

    location_detail = location_detail.merge(sap_dim, on="Eval Group SAP", how="left")
    location_detail = location_detail.merge(ea_dim, on="Eval Group eAsset", how="left")
    location_detail["No Aset Numerik"] = pd.to_numeric(
        location_detail["No Aset SAP"], errors="coerce"
    )
    location_detail["PTJ Penapis"] = (
        location_detail["PTJ SAP"].fillna(location_detail["PTJ eAsset"])
    )

    return sap, easset, dim, compare, location_detail


def fmt_rm(value):
    return f"RM {value:,.2f}"


def csv_bytes(df):
    return df.to_csv(index=False).encode("utf-8-sig")


st.markdown(
    """
    <div class="title-card">
        <h1>🏢 Dashboard Aset SAP vs eAsset</h1>
        <p>Semakan padanan nombor aset, nilai perolehan, nilai buku dan maklumat penempatan.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if DEFAULT_FILE.exists():
    file_bytes = DEFAULT_FILE.read_bytes()
else:
    st.error("Fail Asset.xlsx tidak ditemui dalam folder aplikasi.")
    st.stop()

try:
    sap, easset, dim, compare, location_detail = load_data(file_bytes)
except Exception as exc:
    st.error(f"Fail tidak dapat dibaca: {exc}")
    st.stop()

# Nombor aset SAP dalam fail ini menggunakan format 9 digit.
# Aset Alih: julat bermula 1 hingga 6. Aset Tak Ketara: julat bermula 7 hingga 9.
compare["No Aset Numerik"] = pd.to_numeric(compare["No Aset SAP"], errors="coerce")

with st.sidebar:
    st.header("🔎 Penapis")

    selected_category = st.selectbox(
        "Kategori",
        [
            "Semua",
            "Aset Alih",
            "Aset Tak Ketara",
        ],
        index=0,
    )
    ptj_options = sorted(
        compare.get("Detail 3", pd.Series(dtype="string"))
        .dropna()
        .astype("string")
        .str.strip()
        .loc[lambda x: x.ne("")]
        .unique()
        .tolist()
    )
    selected_ptj = st.selectbox("PTJ", ["Semua"] + ptj_options, index=0)

filtered = compare[
    compare["No Aset Numerik"].between(100000000, 999999999, inclusive="both")
].copy()

if selected_category == "Aset Alih":
    filtered = filtered[
        filtered["No Aset Numerik"].between(100000000, 699999999, inclusive="both")
    ]
elif selected_category == "Aset Tak Ketara":
    filtered = filtered[
        filtered["No Aset Numerik"].between(700000000, 999999999, inclusive="both")
    ]

if selected_ptj != "Semua" and "Detail 3" in filtered.columns:
    filtered = filtered[filtered["Detail 3"].astype("string").str.strip().eq(selected_ptj)]

matched = filtered.loc[filtered["Sumber Rekod"].eq("SAP & eAsset"), "No Aset SAP"].nunique()
sap_only = filtered.loc[filtered["Sumber Rekod"].eq("SAP Sahaja"), "No Aset SAP"].nunique()
ea_only = filtered.loc[filtered["Sumber Rekod"].eq("eAsset Sahaja"), "No Aset SAP"].nunique()
filtered_location = location_detail[
    location_detail["No Aset Numerik"].between(100000000, 999999999, inclusive="both")
].copy()
if selected_category == "Aset Alih":
    filtered_location = filtered_location[
        filtered_location["No Aset Numerik"].between(100000000, 699999999, inclusive="both")
    ]
elif selected_category == "Aset Tak Ketara":
    filtered_location = filtered_location[
        filtered_location["No Aset Numerik"].between(700000000, 999999999, inclusive="both")
    ]
if selected_ptj != "Semua":
    filtered_location = filtered_location[
        filtered_location["PTJ Penapis"].astype("string").str.strip().eq(selected_ptj)
    ]

# KPI lokasi mengira bilangan rekod/baris, selaras dengan working (60 rekod bagi Semua).
location_mismatch = len(filtered_location)
classification_mismatch = filtered.loc[filtered["Salah Klasifikasi"], "No Aset SAP"].nunique()

m1, m2, m3, m4 = st.columns(4)
m1.metric("Jumlah Aset di SAP tetapi Tiada di eAsset", f"{sap_only:,}")
m2.metric("Jumlah Aset di eAsset tetapi Tiada di SAP", f"{ea_only:,}")
m3.metric("Jumlah Aset Berlainan Lokasi", f"{location_mismatch:,}")
m4.metric("Aset Salah Klasifikasi", f"{classification_mismatch:,}")

st.caption(
    f"Jumlah nilai perolehan SAP: **{fmt_rm(filtered['Nilai Perolehan SAP'].sum())}**  ·  "
    f"Jumlah nilai eAsset: **{fmt_rm(filtered['Nilai eAsset'].sum())}**  ·  "
    f"Jumlah nilai buku SAP: **{fmt_rm(filtered['Nilai Buku SAP'].sum())}**"
)

tab1, tab2, tab3, tab4 = st.tabs(
    ["📊 Ringkasan", "🔁 Perbandingan", "⚠️ Isu & Pengecualian", "📁 Data Asal"]
)

with tab1:
    c1, c2 = st.columns(2)
    status_summary = (
        filtered.groupby("Status Padanan", as_index=False)["No Aset SAP"]
        .nunique()
        .rename(columns={"No Aset SAP": "Bilangan"})
    )
    fig_status = px.bar(
        status_summary,
        x="Status Padanan",
        y="Bilangan",
        text_auto=",",
        title="Status Padanan Aset",
    )
    fig_status.update_layout(xaxis_title=None, yaxis_title="Bilangan Aset", showlegend=False)
    c1.plotly_chart(fig_status, use_container_width=True)

    source_summary = (
        filtered.groupby("Sumber Rekod", as_index=False)["No Aset SAP"]
        .nunique()
        .rename(columns={"No Aset SAP": "Bilangan"})
    )
    fig_source = px.pie(
        source_summary,
        names="Sumber Rekod",
        values="Bilangan",
        hole=0.55,
        title="Liputan Rekod Mengikut Sistem",
    )
    c2.plotly_chart(fig_source, use_container_width=True)

    c3, c4 = st.columns(2)
    by_group = (
        filtered.groupby(["Eval Group", "Status Padanan"], dropna=False, as_index=False)["No Aset SAP"]
        .nunique()
        .rename(columns={"No Aset SAP": "Bilangan"})
    )
    top_groups = (
        filtered.groupby("Eval Group")["No Aset SAP"].nunique().nlargest(15).index
    )
    by_group = by_group[by_group["Eval Group"].isin(top_groups)]
    fig_group = px.bar(
        by_group,
        x="Eval Group",
        y="Bilangan",
        color="Status Padanan",
        title="15 Eval Group Terbesar",
    )
    c3.plotly_chart(fig_group, use_container_width=True)

    year_summary = (
        filtered.dropna(subset=["Tahun Beli"])
        .groupby("Tahun Beli", as_index=False)["No Aset SAP"]
        .nunique()
        .rename(columns={"No Aset SAP": "Bilangan"})
    )
    fig_year = px.line(
        year_summary,
        x="Tahun Beli",
        y="Bilangan",
        markers=True,
        title="Bilangan Aset Mengikut Tahun Beli (eAsset)",
    )
    c4.plotly_chart(fig_year, use_container_width=True)

with tab2:
    st.subheader("Senarai Perbandingan Aset")
    display_cols = [
        "No Aset SAP", "Status Padanan", "Sumber Rekod", "Eval Group SAP", "Eval Group eAsset",
        "Klasifikasi SAP", "Klasifikasi eAsset", "PTJ SAP", "PTJ eAsset",
        "Berlainan Lokasi", "Salah Klasifikasi", "Deskripsi Aset", "No Siri Pendaftaran",
        "Jenama", "Lokasi", "Pegawai Penempatan",
        "Nilai Perolehan SAP", "Nilai eAsset", "Perbezaan Nilai", "Nilai Buku SAP",
        "Bil Rekod SAP", "Bil Rekod eAsset", "Tarikh Beli"
    ]
    display_cols = [c for c in display_cols if c in filtered.columns]
    st.dataframe(
        filtered[display_cols].sort_values("Perbezaan Mutlak", ascending=False),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Nilai Perolehan SAP": st.column_config.NumberColumn(format="RM %.2f"),
            "Nilai eAsset": st.column_config.NumberColumn(format="RM %.2f"),
            "Perbezaan Nilai": st.column_config.NumberColumn(format="RM %.2f"),
            "Nilai Buku SAP": st.column_config.NumberColumn(format="RM %.2f"),
            "Tarikh Beli": st.column_config.DateColumn(format="DD/MM/YYYY"),
        },
        height=650,
    )
    st.download_button(
        "⬇️ Muat Turun Keputusan Ditapis (CSV)",
        data=csv_bytes(filtered[display_cols]),
        file_name="perbandingan_aset_sap_vs_easset.csv",
        mime="text/csv",
    )

with tab3:
    issue = filtered[filtered["Status Padanan"] != "Padan"].copy()
    i1, i2, i3 = st.columns(3)
    i1.metric("Jumlah Pengecualian", f"{len(issue):,}")
    i2.metric("Jumlah Perbezaan Mutlak", fmt_rm(issue["Perbezaan Mutlak"].sum()))
    dup_ea = (filtered[filtered["Bil Rekod eAsset"].fillna(0) > 1]["No Aset SAP"].nunique())
    i3.metric("No. Aset Duplicate eAsset", f"{dup_ea:,}")

    issue_type = (
        issue.groupby("Status Padanan", as_index=False)
        .agg(Bilangan=("No Aset SAP", "nunique"), Perbezaan_RM=("Perbezaan Mutlak", "sum"))
    )
    st.dataframe(
        issue_type,
        use_container_width=True,
        hide_index=True,
        column_config={"Perbezaan_RM": st.column_config.NumberColumn("Perbezaan (RM)", format="RM %.2f")},
    )

    st.subheader("Aset Berlainan Lokasi")
    lokasi_cols = [c for c in [
        "No Aset SAP", "Eval Group SAP", "PTJ SAP", "Eval Group eAsset", "PTJ eAsset",
        "Deskripsi Aset", "Lokasi", "Pegawai Penempatan"
    ] if c in filtered.columns]
    lokasi_cols = [c for c in [
        "No Aset SAP", "Eval Group SAP", "PTJ SAP", "Eval Group eAsset", "PTJ eAsset",
        "Deskripsi Aset", "Deskripsi Ringkas", "No Siri Pendaftaran",
        "Jenis", "Jenama", "Lokasi", "Pegawai Penempatan"
    ] if c in filtered_location.columns]
    st.dataframe(
        filtered_location[lokasi_cols],
        use_container_width=True,
        hide_index=True,
        height=350,
    )

    st.subheader("Aset Salah Klasifikasi")
    klasifikasi_cols = [c for c in [
        "No Aset SAP", "Eval Group SAP", "Klasifikasi SAP", "Eval Group eAsset",
        "Klasifikasi eAsset", "PTJ SAP", "PTJ eAsset", "Deskripsi Aset", "Jenis"
    ] if c in filtered.columns]
    st.dataframe(
        filtered.loc[filtered["Salah Klasifikasi"], klasifikasi_cols],
        use_container_width=True,
        hide_index=True,
        height=350,
    )

    st.subheader("Top 50 Perbezaan Nilai Tertinggi")
    top_diff = issue.nlargest(50, "Perbezaan Mutlak")
    st.dataframe(
        top_diff[[c for c in ["No Aset SAP", "Status Padanan", "Deskripsi Aset", "Jenama", "Lokasi", "Nilai Perolehan SAP", "Nilai eAsset", "Perbezaan Nilai"] if c in top_diff.columns]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Nilai Perolehan SAP": st.column_config.NumberColumn(format="RM %.2f"),
            "Nilai eAsset": st.column_config.NumberColumn(format="RM %.2f"),
            "Perbezaan Nilai": st.column_config.NumberColumn(format="RM %.2f"),
        },
    )

    st.subheader("Rekod Duplicate dalam eAsset")
    duplicate_numbers = filtered.loc[filtered["Bil Rekod eAsset"].fillna(0) > 1, "No Aset SAP"]
    duplicate_rows = easset[easset["No Aset SAP"].isin(duplicate_numbers)].sort_values("No Aset SAP")
    st.dataframe(duplicate_rows, use_container_width=True, hide_index=True, height=420)

with tab4:
    raw1, raw2, raw3 = st.tabs(["SAP", "eAsset", "Dimensi Eval Group"])
    with raw1:
        st.dataframe(sap, use_container_width=True, hide_index=True, height=600)
    with raw2:
        st.dataframe(easset, use_container_width=True, hide_index=True, height=600)
    with raw3:
        st.dataframe(dim, use_container_width=True, hide_index=True, height=600)

st.divider()
st.caption("Semua KPI menggunakan No. Aset sebagai key utama. Jumlah Aset Berlainan Lokasi mengira bilangan rekod/baris yang mempunyai Eval Group SAP berbeza daripada Eval Group eAsset; jumlah keseluruhan fail semasa ialah 60 rekod.")
