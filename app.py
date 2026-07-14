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
    </style>
    """,
    unsafe_allow_html=True,
)

DEFAULT_FILE = Path(__file__).with_name("Asset.xlsx")


def clean_text(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .str.strip()
        .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})
    )


def clean_asset_no(series: pd.Series) -> pd.Series:
    return clean_text(series).str.replace(r"\.0$", "", regex=True)


def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series.astype("string").str.replace(",", "", regex=False).str.strip(),
        errors="coerce",
    ).fillna(0.0)


def join_unique(series: pd.Series, limit: int = 5) -> str:
    values = series.dropna().astype(str).str.strip()
    values = values[values.ne("")].drop_duplicates().head(limit)
    return " | ".join(values)


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
            "Eval. Grp 1-Block": "Eval Group SAP",
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
            "Harga (RM)": "Nilai eAsset",
            "Eval Group 1": "Eval Group eAsset",
        }
    )
    dim = dim.rename(columns={"Eval. Grp 1-Block": "Eval Group"})

    sap["No Aset SAP"] = clean_asset_no(sap["No Aset SAP"])
    easset["No Aset SAP"] = clean_asset_no(easset["No Aset SAP"])
    sap["Eval Group SAP"] = clean_text(sap["Eval Group SAP"])
    easset["Eval Group eAsset"] = clean_text(easset["Eval Group eAsset"])
    dim["Eval Group"] = clean_text(dim["Eval Group"])

    sap = sap[sap["No Aset SAP"].notna()].copy()
    easset = easset[easset["No Aset SAP"].notna()].copy()

    sap["No Aset Numerik"] = pd.to_numeric(sap["No Aset SAP"], errors="coerce")
    easset["No Aset Numerik"] = pd.to_numeric(easset["No Aset SAP"], errors="coerce")

    sap["Nilai Perolehan SAP"] = to_number(sap["Nilai Perolehan SAP"])
    sap["Nilai Buku SAP"] = to_number(sap["Nilai Buku SAP"])
    easset["Nilai eAsset"] = to_number(easset["Nilai eAsset"])
    easset["Tarikh Beli"] = pd.to_datetime(easset.get("Tarikh Beli"), errors="coerce")
    easset["Tarikh Invois"] = pd.to_datetime(easset.get("Tarikh Invois"), errors="coerce")

    dim_map = (
        dim[["Eval Group", "Detail", "Detail 2", "Detail 3"]]
        .dropna(subset=["Eval Group"])
        .drop_duplicates("Eval Group")
    )

    sap = sap.merge(
        dim_map.rename(
            columns={
                "Eval Group": "Eval Group SAP",
                "Detail": "Klasifikasi SAP",
                "Detail 2": "Zon SAP",
                "Detail 3": "PTJ SAP",
            }
        ),
        on="Eval Group SAP",
        how="left",
    )

    easset = easset.merge(
        dim_map.rename(
            columns={
                "Eval Group": "Eval Group eAsset",
                "Detail": "Klasifikasi eAsset",
                "Detail 2": "Zon eAsset",
                "Detail 3": "PTJ eAsset",
            }
        ),
        on="Eval Group eAsset",
        how="left",
    )

    sap_assets = set(sap["No Aset SAP"].dropna())
    easset_assets = set(easset["No Aset SAP"].dropna())

    # Padanan pada tahap baris. Kaedah ini mengekalkan 60 rekod Eval Group berbeza.
    location_compare = sap.merge(
        easset,
        on="No Aset SAP",
        how="inner",
        suffixes=("", "_eAsset"),
    )
    location_compare["Eval Group SAP"] = clean_text(location_compare["Eval Group SAP"])
    location_compare["Eval Group eAsset"] = clean_text(location_compare["Eval Group eAsset"])
    location_compare = location_compare[
        location_compare["Eval Group SAP"].notna()
        & location_compare["Eval Group eAsset"].notna()
        & location_compare["Eval Group SAP"].ne(location_compare["Eval Group eAsset"])
    ].copy()

    location_compare["Status Lokasi"] = "Berlainan"

    return sap, easset, location_compare, sap_assets, easset_assets


def category_mask(df: pd.DataFrame, selected_category: str) -> pd.Series:
    numeric = pd.to_numeric(df["No Aset SAP"], errors="coerce")
    if selected_category == "Aset Alih":
        return numeric.between(100000000, 699999999, inclusive="both")
    if selected_category == "Aset Tak Ketara":
        return numeric.between(700000000, 999999999, inclusive="both")
    return numeric.between(100000000, 999999999, inclusive="both")


def dataframe_to_excel(df: pd.DataFrame, sheet_name: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name[:31])
        ws = writer.book[sheet_name[:31]]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for column_cells in ws.columns:
            max_length = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in column_cells
            )
            ws.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 45)
    return output.getvalue()


def reports_to_excel(reports: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in reports.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)
            ws = writer.book[safe_name]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for column_cells in ws.columns:
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                ws.column_dimensions[column_cells[0].column_letter].width = min(max_length + 2, 45)
    return output.getvalue()


st.markdown(
    """
    <div class="title-card">
        <h1>🏢 Dashboard Aset SAP vs eAsset</h1>
        <p>Semakan aset berdasarkan nombor aset, PTJ dan Eval Group.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if not DEFAULT_FILE.exists():
    st.error("Fail Asset.xlsx tidak ditemui dalam folder aplikasi.")
    st.stop()

try:
    sap, easset, location_compare, sap_assets, easset_assets = load_data(DEFAULT_FILE.read_bytes())
except Exception as exc:
    st.error(f"Fail tidak dapat dibaca: {exc}")
    st.stop()

all_ptj = sorted(
    pd.concat(
        [
            sap["PTJ SAP"].dropna().astype("string"),
            easset["PTJ eAsset"].dropna().astype("string"),
        ],
        ignore_index=True,
    )
    .str.strip()
    .loc[lambda s: s.ne("")]
    .drop_duplicates()
    .tolist()
)

with st.sidebar:
    st.header("🔎 Penapis")
    selected_category = st.selectbox(
        "Kategori",
        ["Semua", "Aset Alih", "Aset Tak Ketara"],
        index=0,
    )
    selected_ptj = st.selectbox("PTJ", ["Semua"] + all_ptj, index=0)

sap_filtered = sap[category_mask(sap, selected_category)].copy()
easset_filtered = easset[category_mask(easset, selected_category)].copy()
location_filtered = location_compare[category_mask(location_compare, selected_category)].copy()

if selected_ptj != "Semua":
    sap_filtered = sap_filtered[
        sap_filtered["PTJ SAP"].astype("string").str.strip().eq(selected_ptj)
    ]
    easset_filtered = easset_filtered[
        easset_filtered["PTJ eAsset"].astype("string").str.strip().eq(selected_ptj)
    ]
    # Rekod lokasi berlainan dipaparkan jika PTJ dipilih wujud pada salah satu sistem.
    location_filtered = location_filtered[
        location_filtered["PTJ SAP"].astype("string").str.strip().eq(selected_ptj)
        | location_filtered["PTJ eAsset"].astype("string").str.strip().eq(selected_ptj)
    ]

sap_filtered_assets = set(sap_filtered["No Aset SAP"].dropna())
easset_filtered_assets = set(easset_filtered["No Aset SAP"].dropna())

sap_only = len(sap_filtered_assets - easset_assets)
easset_only = len(easset_filtered_assets - sap_assets)
location_mismatch = len(location_filtered)

# Salah klasifikasi: nombor aset sama tetapi klasifikasi DIM berbeza.
classification_compare = sap_filtered[["No Aset SAP", "Klasifikasi SAP"]].merge(
    easset_filtered[["No Aset SAP", "Klasifikasi eAsset"]],
    on="No Aset SAP",
    how="inner",
)
classification_compare = classification_compare[
    classification_compare["Klasifikasi SAP"].notna()
    & classification_compare["Klasifikasi eAsset"].notna()
    & clean_text(classification_compare["Klasifikasi SAP"]).ne(
        clean_text(classification_compare["Klasifikasi eAsset"])
    )
]
classification_mismatch = classification_compare["No Aset SAP"].nunique()

m1, m2, m3, m4 = st.columns(4)
m1.metric("Jumlah Aset di SAP tetapi Tiada di eAsset", f"{sap_only:,}")
m2.metric("Jumlah Aset di eAsset tetapi Tiada di SAP", f"{easset_only:,}")
m3.metric("Jumlah Aset Berlainan Lokasi", f"{location_mismatch:,}")
m4.metric("Aset Salah Klasifikasi", f"{classification_mismatch:,}")

st.markdown("---")

# Data bagi carta: hanya aset yang tidak wujud dalam sistem satu lagi.
sap_only_assets = sap_filtered_assets - easset_assets
easset_only_assets = easset_filtered_assets - sap_assets

sap_only_df = sap_filtered[sap_filtered["No Aset SAP"].isin(sap_only_assets)].copy()
easset_only_df = easset_filtered[easset_filtered["No Aset SAP"].isin(easset_only_assets)].copy()

chart_col1, chart_col2 = st.columns(2, gap="large")

with chart_col1:
    sap_chart = (
        sap_only_df.assign(PTJ=lambda d: d["PTJ SAP"].fillna("Tidak Dikenal Pasti"))
        .groupby("PTJ", as_index=False)["No Aset SAP"]
        .nunique()
        .rename(columns={"No Aset SAP": "Jumlah Aset"})
        .sort_values("Jumlah Aset", ascending=True)
    )
    st.subheader("📊 Aset di SAP tetapi Tiada di eAsset Mengikut PTJ")
    st.metric("Jumlah", f"{len(sap_only_assets):,}")
    if sap_chart.empty:
        st.info("Tiada aset SAP yang tiada dalam eAsset bagi penapis yang dipilih.")
    else:
        fig_sap = px.bar(
            sap_chart,
            x="Jumlah Aset",
            y="PTJ",
            orientation="h",
            text="Jumlah Aset",
        )
        fig_sap.update_traces(textposition="outside", cliponaxis=False)
        fig_sap.update_layout(
            xaxis_title="Jumlah Aset",
            yaxis_title=None,
            height=max(480, len(sap_chart) * 31),
            margin=dict(l=10, r=35, t=10, b=10),
            showlegend=False,
        )
        st.plotly_chart(fig_sap, use_container_width=True)

with chart_col2:
    easset_chart = (
        easset_only_df.assign(PTJ=lambda d: d["PTJ eAsset"].fillna("Tidak Dikenal Pasti"))
        .groupby("PTJ", as_index=False)["No Aset SAP"]
        .nunique()
        .rename(columns={"No Aset SAP": "Jumlah Aset"})
        .sort_values("Jumlah Aset", ascending=True)
    )
    st.subheader("📊 Aset di eAsset tetapi Tiada di SAP Mengikut PTJ")
    st.metric("Jumlah", f"{len(easset_only_assets):,}")
    if easset_chart.empty:
        st.info("Tiada aset eAsset yang tiada dalam SAP bagi penapis yang dipilih.")
    else:
        fig_easset = px.bar(
            easset_chart,
            x="Jumlah Aset",
            y="PTJ",
            orientation="h",
            text="Jumlah Aset",
        )
        fig_easset.update_traces(textposition="outside", cliponaxis=False)
        fig_easset.update_layout(
            xaxis_title="Jumlah Aset",
            yaxis_title=None,
            height=max(480, len(easset_chart) * 31),
            margin=dict(l=10, r=35, t=10, b=10),
            showlegend=False,
        )
        st.plotly_chart(fig_easset, use_container_width=True)

st.markdown("---")
st.subheader("📋 Laporan")

sap_report_cols = [
    "No Aset SAP",
    "Sub No",
    "Eval Group SAP",
    "PTJ SAP",
    "Klasifikasi SAP",
    "Zon SAP",
    "Deskripsi Ringkas",
    "Deskripsi Aset",
    "Nilai Perolehan SAP",
    "Nilai Buku SAP",
]
easset_report_cols = [
    "Bil.",
    "No. Siri Pendaftaran",
    "No Aset SAP",
    "Eval Group eAsset",
    "PTJ eAsset",
    "Klasifikasi eAsset",
    "Zon eAsset",
    "Jenis",
    "Jenama",
    "Lokasi",
    "Pegawai Penempatan",
    "No. Pesanan",
    "Tarikh Beli",
    "Nilai eAsset",
    "Tarikh Invois",
]
location_report_cols = [
    "No Aset SAP",
    "Sub No",
    "Eval Group SAP",
    "PTJ SAP",
    "Klasifikasi SAP",
    "Eval Group eAsset",
    "PTJ eAsset",
    "Klasifikasi eAsset",
    "Deskripsi Aset",
    "No. Siri Pendaftaran",
    "Jenis",
    "Jenama",
    "Lokasi",
    "Pegawai Penempatan",
    "Status Lokasi",
]

sap_report = sap_filtered[[c for c in sap_report_cols if c in sap_filtered.columns]].copy()
easset_report = easset_filtered[[c for c in easset_report_cols if c in easset_filtered.columns]].copy()
location_report = location_filtered[
    [c for c in location_report_cols if c in location_filtered.columns]
].copy()

sap_report = sap_report.sort_values(["PTJ SAP", "No Aset SAP"], na_position="last")
easset_report = easset_report.sort_values(["PTJ eAsset", "No Aset SAP"], na_position="last")
location_report = location_report.sort_values(
    ["PTJ SAP", "PTJ eAsset", "No Aset SAP"], na_position="last"
)

combined_excel = reports_to_excel(
    {
        "SAP": sap_report,
        "eAsset": easset_report,
        "Lokasi Berlainan": location_report,
    }
)
st.download_button(
    "⬇️ Muat Turun Semua Laporan (Excel)",
    data=combined_excel,
    file_name="laporan_aset_sap_easset.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)

tab_sap, tab_easset, tab_location = st.tabs(
    ["SAP", "eAsset", "Lokasi Berlainan"]
)

with tab_sap:
    st.caption(f"Jumlah rekod SAP: {len(sap_report):,} | Aset unik: {sap_report['No Aset SAP'].nunique():,}")
    st.dataframe(sap_report, use_container_width=True, hide_index=True, height=560)
    st.download_button(
        "⬇️ Download Laporan SAP",
        data=dataframe_to_excel(sap_report, "SAP"),
        file_name="laporan_sap.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_sap",
    )

with tab_easset:
    st.caption(
        f"Jumlah rekod eAsset: {len(easset_report):,} | Aset unik: {easset_report['No Aset SAP'].nunique():,}"
    )
    st.dataframe(easset_report, use_container_width=True, hide_index=True, height=560)
    st.download_button(
        "⬇️ Download Laporan eAsset",
        data=dataframe_to_excel(easset_report, "eAsset"),
        file_name="laporan_easset.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_easset",
    )

with tab_location:
    st.caption(
        f"Jumlah rekod lokasi berlainan: {len(location_report):,} | Aset unik: {location_report['No Aset SAP'].nunique():,}"
    )
    st.dataframe(location_report, use_container_width=True, hide_index=True, height=560)
    st.download_button(
        "⬇️ Download Laporan Lokasi Berlainan",
        data=dataframe_to_excel(location_report, "Lokasi Berlainan"),
        file_name="laporan_lokasi_berlainan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="download_location",
    )
