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

    compare["Eval Group"] = compare["Eval Group"].fillna(compare["Eval Group eAsset"])
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

    dim_cols = [c for c in ["Eval Group", "Detail", "Detail 2", "Detail 3"] if c in dim.columns]
    compare = compare.merge(dim[dim_cols].drop_duplicates("Eval Group"), on="Eval Group", how="left")

    # Tahun aset untuk analisis umur rekod eAsset.
    compare["Tahun Beli"] = compare["Tarikh Beli"].dt.year.astype("Int64")
    return sap, easset, dim, compare


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

with st.sidebar:
    st.header("⚙️ Sumber Data")
    upload = st.file_uploader("Muat naik fail Excel terkini", type=["xlsx"])
    if upload is not None:
        file_bytes = upload.getvalue()
        st.success(f"Menggunakan: {upload.name}")
    elif DEFAULT_FILE.exists():
        file_bytes = DEFAULT_FILE.read_bytes()
        st.info("Menggunakan fail Asset.xlsx dalam aplikasi")
    else:
        st.warning("Sila muat naik fail Excel.")
        st.stop()

try:
    sap, easset, dim, compare = load_data(file_bytes)
except Exception as exc:
    st.error(f"Fail tidak dapat dibaca: {exc}")
    st.stop()

with st.sidebar:
    st.divider()
    st.header("🔎 Penapis")
    source_options = sorted(compare["Sumber Rekod"].dropna().unique())
    status_options = sorted(compare["Status Padanan"].dropna().unique())
    group_options = sorted(compare["Eval Group"].dropna().astype(str).unique())
    office_options = sorted(compare.get("Detail 3", pd.Series(dtype=str)).dropna().astype(str).unique())

    selected_source = st.multiselect("Sumber Rekod", source_options, default=source_options)
    selected_status = st.multiselect("Status Padanan", status_options, default=status_options)
    selected_group = st.multiselect("Eval Group", group_options)
    selected_office = st.multiselect("Kategori Pejabat", office_options)
    search = st.text_input("Cari nombor / deskripsi aset")
    min_diff = st.number_input("Perbezaan minimum (RM)", min_value=0.0, value=0.0, step=100.0)

filtered = compare[
    compare["Sumber Rekod"].isin(selected_source)
    & compare["Status Padanan"].isin(selected_status)
    & compare["Perbezaan Mutlak"].ge(min_diff)
].copy()
if selected_group:
    filtered = filtered[filtered["Eval Group"].astype(str).isin(selected_group)]
if selected_office and "Detail 3" in filtered.columns:
    filtered = filtered[filtered["Detail 3"].astype(str).isin(selected_office)]
if search:
    q = search.lower().strip()
    searchable = filtered[[c for c in ["No Aset SAP", "Deskripsi Aset", "Jenama", "Lokasi", "No Siri Pendaftaran"] if c in filtered.columns]].fillna("").astype(str)
    mask = searchable.apply(lambda col: col.str.lower().str.contains(q, regex=False)).any(axis=1)
    filtered = filtered[mask]

matched = (compare["Sumber Rekod"] == "SAP & eAsset").sum()
sap_only = (compare["Sumber Rekod"] == "SAP Sahaja").sum()
ea_only = (compare["Sumber Rekod"] == "eAsset Sahaja").sum()
value_mismatch = (compare["Status Padanan"] == "Nilai Tidak Padan").sum()
match_rate = matched / max(compare["No Aset SAP"].nunique(), 1) * 100

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Aset Unik", f"{compare['No Aset SAP'].nunique():,}")
m2.metric("Padan SAP & eAsset", f"{matched:,}", f"{match_rate:.1f}%")
m3.metric("SAP Sahaja", f"{sap_only:,}")
m4.metric("eAsset Sahaja", f"{ea_only:,}")
m5.metric("Nilai Tidak Padan", f"{value_mismatch:,}")

st.caption(
    f"Jumlah nilai perolehan SAP: **{fmt_rm(compare['Nilai Perolehan SAP'].sum())}**  ·  "
    f"Jumlah nilai eAsset: **{fmt_rm(compare['Nilai eAsset'].sum())}**  ·  "
    f"Jumlah nilai buku SAP: **{fmt_rm(compare['Nilai Buku SAP'].sum())}**"
)

tab1, tab2, tab3, tab4 = st.tabs(
    ["📊 Ringkasan", "🔁 Perbandingan", "⚠️ Isu & Pengecualian", "📁 Data Asal"]
)

with tab1:
    c1, c2 = st.columns(2)
    status_summary = (
        compare.groupby("Status Padanan", as_index=False)["No Aset SAP"]
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
        compare.groupby("Sumber Rekod", as_index=False)["No Aset SAP"]
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
        compare.groupby(["Eval Group", "Status Padanan"], dropna=False, as_index=False)["No Aset SAP"]
        .nunique()
        .rename(columns={"No Aset SAP": "Bilangan"})
    )
    top_groups = (
        compare.groupby("Eval Group")["No Aset SAP"].nunique().nlargest(15).index
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
        compare.dropna(subset=["Tahun Beli"])
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
        "No Aset SAP", "Status Padanan", "Sumber Rekod", "Eval Group", "Detail", "Detail 3",
        "Deskripsi Aset", "No Siri Pendaftaran", "Jenama", "Lokasi", "Pegawai Penempatan",
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
    issue = compare[compare["Status Padanan"] != "Padan"].copy()
    i1, i2, i3 = st.columns(3)
    i1.metric("Jumlah Pengecualian", f"{len(issue):,}")
    i2.metric("Jumlah Perbezaan Mutlak", fmt_rm(issue["Perbezaan Mutlak"].sum()))
    dup_ea = (easset.groupby("No Aset SAP").size() > 1).sum()
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
    duplicate_numbers = easset.groupby("No Aset SAP").size().loc[lambda s: s > 1].index
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
st.caption("Dashboard ini mengagregat rekod berdasarkan No. Aset SAP. Toleransi padanan nilai ditetapkan pada RM1.00.")
