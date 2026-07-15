import io
from pathlib import Path

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
    html {scroll-behavior: smooth;}
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    section.main {
        scroll-snap-type: y mandatory;
        scroll-behavior: smooth;
    }
    .st-key-page_one, .st-key-page_two {
        min-height: calc(100vh - 2rem);
        scroll-snap-align: start;
        scroll-snap-stop: always;
    }
    .st-key-kpi_sap button,
    .st-key-kpi_easset button,
    .st-key-kpi_lokasi button {
        width: 100%;
        min-height: 145px;
        background: white;
        border: 1px solid #dbe4ea;
        border-radius: 18px;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.08);
        font-size: 1.35rem !important;
        font-weight: 800 !important;
        white-space: pre-line;
        line-height: 1.6 !important;
    }
    .st-key-kpi_sap button:hover,
    .st-key-kpi_easset button:hover,
    .st-key-kpi_lokasi button:hover {
        border-color: #0f766e;
        transform: translateY(-2px);
        box-shadow: 0 12px 28px rgba(15, 118, 110, 0.16);
    }
    .st-key-kpi_sap button p,
    .st-key-kpi_easset button p,
    .st-key-kpi_lokasi button p {
        font-size: 1.35rem !important;
        line-height: 1.6 !important;
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


def excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output.getvalue()


@st.cache_data(show_spinner=False)
def load_data(path: str):
    xls = pd.ExcelFile(path)
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
            "Asset": "No Aset",
            "Sub-number": "Sub No",
            "Asset Description": "Deskripsi Ringkas",
            "Asset Description.1": "Nama Aset",
            "Acquis.val.": "Nilai Perolehan SAP",
            "Book val.": "Nilai Buku SAP",
        }
    )
    easset = easset.rename(
        columns={
            "No. Aset SAP": "No Aset",
            "Eval Group 1": "Eval Group eAsset",
            "Harga (RM)": "Nilai eAsset",
        }
    )
    dim = dim.rename(columns={"Eval. Grp 1-Block": "Eval Group"})

    sap["No Aset"] = clean_asset_no(sap["No Aset"])
    easset["No Aset"] = clean_asset_no(easset["No Aset"])
    sap["Eval Group SAP"] = clean_text(sap["Eval Group SAP"])
    easset["Eval Group eAsset"] = clean_text(easset["Eval Group eAsset"])
    dim["Eval Group"] = clean_text(dim["Eval Group"])
    dim["Detail 3"] = clean_text(dim["Detail 3"])

    sap = sap[sap["No Aset"].notna()].copy()
    easset = easset[easset["No Aset"].notna()].copy()

    dim_map = (
        dim[["Eval Group", "Detail 3"]]
        .dropna(subset=["Eval Group"])
        .drop_duplicates("Eval Group")
    )

    sap = sap.merge(
        dim_map.rename(columns={"Eval Group": "Eval Group SAP", "Detail 3": "PTJ SAP"}),
        on="Eval Group SAP",
        how="left",
    )
    easset = easset.merge(
        dim_map.rename(columns={"Eval Group": "Eval Group eAsset", "Detail 3": "PTJ eAsset"}),
        on="Eval Group eAsset",
        how="left",
    )

    sap["PTJ SAP"] = clean_text(sap["PTJ SAP"])
    easset["PTJ eAsset"] = clean_text(easset["PTJ eAsset"])

    # Senarai unik digunakan untuk aset yang hanya wujud dalam satu sistem.
    sap_unique = sap.drop_duplicates("No Aset").copy()
    easset_unique = easset.drop_duplicates("No Aset").copy()

    sap_numbers = set(sap_unique["No Aset"].dropna())
    easset_numbers = set(easset_unique["No Aset"].dropna())

    sap_missing = sap_unique[~sap_unique["No Aset"].isin(easset_numbers)].copy()
    easset_missing = easset_unique[~easset_unique["No Aset"].isin(sap_numbers)].copy()

    # Perbandingan lokasi dibuat mengikut setiap baris yang mempunyai No. Aset sama.
    # Kaedah ini mengekalkan 60 rekod working apabila terdapat duplicate/sub-record.
    lokasi = sap.merge(easset, on="No Aset", how="inner", suffixes=("", "_eAsset"))
    lokasi = lokasi[
        lokasi["Eval Group SAP"].notna()
        & lokasi["Eval Group eAsset"].notna()
        & lokasi["Eval Group SAP"].ne(lokasi["Eval Group eAsset"])
    ].copy()

    # Nama aset diletakkan betul-betul di sebelah No. Aset.
    lokasi["Nama Aset"] = clean_text(lokasi.get("Nama Aset", pd.Series(index=lokasi.index, dtype="string")))
    lokasi["Nama Aset"] = lokasi["Nama Aset"].fillna(
        clean_text(lokasi.get("Deskripsi Ringkas", pd.Series(index=lokasi.index, dtype="string")))
    )

    lokasi_cols = [
        "No Aset",
        "Nama Aset",
        "Eval Group SAP",
        "PTJ SAP",
        "Eval Group eAsset",
        "PTJ eAsset",
        "No. Siri Pendaftaran",
        "Jenis",
        "Jenama",
        "Lokasi",
        "Pegawai Penempatan",
    ]
    lokasi = lokasi[[c for c in lokasi_cols if c in lokasi.columns]].copy()

    return sap, easset, sap_missing, easset_missing, lokasi


st.markdown(
    """
    <div class="title-card">
        <h1>🏢 Dashboard Aset SAP vs eAsset</h1>
        <p>Perbandingan aset berdasarkan No. Aset dan Eval Group.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

if not DEFAULT_FILE.exists():
    st.error("Fail Asset.xlsx tidak ditemui dalam folder aplikasi.")
    st.stop()

try:
    sap, easset, sap_missing_all, easset_missing_all, lokasi_all = load_data(str(DEFAULT_FILE))
except Exception as exc:
    st.error(f"Fail tidak dapat dibaca: {exc}")
    st.stop()


def apply_category(df: pd.DataFrame, asset_col: str = "No Aset") -> pd.DataFrame:
    result = df.copy()
    result["_No Aset Numerik"] = pd.to_numeric(result[asset_col], errors="coerce")
    result = result[result["_No Aset Numerik"] >= 100000000]

    if selected_category == "Aset Tak Alih":
        result = result[result["_No Aset Numerik"].between(100000000, 299999999, inclusive="both")]
    elif selected_category == "Aset Alih":
        result = result[result["_No Aset Numerik"].between(300000000, 799999999, inclusive="both")]
    elif selected_category == "Aset Tak Ketara":
        result = result[result["_No Aset Numerik"] >= 800000000]

    return result.drop(columns="_No Aset Numerik")


ptj_values = pd.concat(
    [
        sap.get("PTJ SAP", pd.Series(dtype="string")),
        easset.get("PTJ eAsset", pd.Series(dtype="string")),
    ],
    ignore_index=True,
)
ptj_options = sorted(clean_text(ptj_values).dropna().unique().tolist())

if "chart_selected_ptj" not in st.session_state:
    st.session_state["chart_selected_ptj"] = None
if "chart_selected_source" not in st.session_state:
    st.session_state["chart_selected_source"] = None
if "kpi_selected_report" not in st.session_state:
    st.session_state["kpi_selected_report"] = None


def clear_chart_selection():
    st.session_state["chart_selected_ptj"] = None
    st.session_state["chart_selected_source"] = None
    st.session_state["kpi_selected_report"] = None


def select_kpi_report(report_name: str):
    st.session_state["kpi_selected_report"] = report_name
    st.session_state["chart_selected_ptj"] = None
    st.session_state["chart_selected_source"] = None


def handle_sap_chart_selection():
    event = st.session_state.get("sap_bar_chart", {})
    points = event.get("selection", {}).get("points", []) if event else []
    if points:
        st.session_state["chart_selected_ptj"] = points[0].get("x")
        st.session_state["chart_selected_source"] = "SAP"


def handle_easset_chart_selection():
    event = st.session_state.get("easset_bar_chart", {})
    points = event.get("selection", {}).get("points", []) if event else []
    if points:
        st.session_state["chart_selected_ptj"] = points[0].get("x")
        st.session_state["chart_selected_source"] = "eAsset"


with st.sidebar:
    st.header("🔎 Penapis")
    selected_category = st.selectbox(
        "Kategori",
        ["Semua", "Aset Tak Alih", "Aset Alih", "Aset Tak Ketara"],
        index=0,
        key="category_filter",
        on_change=clear_chart_selection,
    )
    selected_ptj = st.selectbox(
        "PTJ",
        ["Semua"] + ptj_options,
        index=0,
        key="ptj_filter",
        on_change=clear_chart_selection,
    )

sap_missing = apply_category(sap_missing_all)
easset_missing = apply_category(easset_missing_all)
lokasi = apply_category(lokasi_all)

if selected_ptj != "Semua":
    sap_missing = sap_missing[sap_missing["PTJ SAP"].eq(selected_ptj)]
    easset_missing = easset_missing[easset_missing["PTJ eAsset"].eq(selected_ptj)]
    lokasi = lokasi[
        lokasi["PTJ SAP"].eq(selected_ptj) | lokasi["PTJ eAsset"].eq(selected_ptj)
    ]

# PAGE 1: KPI dan bar chart
with st.container(key="page_one"):
    # KPI boleh diklik untuk memaparkan laporan berkaitan.
    m1, m2, m3 = st.columns(3)
    with m1:
        if st.button(
            f"Aset SAP tiada di eAsset\n\n{sap_missing['No Aset'].nunique():,}",
            key="kpi_sap",
            use_container_width=True,
        ):
            select_kpi_report("SAP")
            st.rerun()
    with m2:
        if st.button(
            f"Aset eAsset tiada di SAP\n\n{easset_missing['No Aset'].nunique():,}",
            key="kpi_easset",
            use_container_width=True,
        ):
            select_kpi_report("eAsset")
            st.rerun()
    with m3:
        if st.button(
            f"Aset Berlainan Lokasi\n\n{len(lokasi):,}",
            key="kpi_lokasi",
            use_container_width=True,
        ):
            select_kpi_report("Lokasi Berlainan")
            st.rerun()

    c1, c2 = st.columns(2)

    sap_chart = (
        sap_missing.assign(PTJ=sap_missing["PTJ SAP"].fillna("Tidak Dikenal Pasti"))
        .groupby("PTJ", as_index=False)["No Aset"]
        .nunique()
        .rename(columns={"No Aset": "Bilangan"})
        .sort_values("Bilangan", ascending=False)
    )

    if sap_chart.empty:
        c1.info("Tiada rekod Aset SAP tiada di eAsset untuk penapis dipilih.")
    else:
        fig_sap = px.bar(
            sap_chart,
            x="PTJ",
            y="Bilangan",
            text="Bilangan",
            title="Aset SAP",
        )
        fig_sap.update_traces(textposition="outside", cliponaxis=False)
        fig_sap.update_layout(
            xaxis_title=None,
            yaxis_title=None,
            showlegend=False,
            margin=dict(t=60, l=20, r=20, b=120),
        )
        with c1:
            st.plotly_chart(
                fig_sap,
                width="stretch",
                key="sap_bar_chart",
                on_select=handle_sap_chart_selection,
                selection_mode="points",
                config={"displayModeBar": False},
            )

    easset_chart = (
        easset_missing.assign(PTJ=easset_missing["PTJ eAsset"].fillna("Tidak Dikenal Pasti"))
        .groupby("PTJ", as_index=False)["No Aset"]
        .nunique()
        .rename(columns={"No Aset": "Bilangan"})
        .sort_values("Bilangan", ascending=False)
    )

    if easset_chart.empty:
        c2.info("Tiada rekod Aset eAsset tiada di SAP untuk penapis dipilih.")
    else:
        fig_easset = px.bar(
            easset_chart,
            x="PTJ",
            y="Bilangan",
            text="Bilangan",
            title="Aset eAsset",
        )
        fig_easset.update_traces(textposition="outside", cliponaxis=False)
        fig_easset.update_layout(
            xaxis_title=None,
            yaxis_title=None,
            showlegend=False,
            margin=dict(t=60, l=20, r=20, b=120),
        )
        with c2:
            st.plotly_chart(
                fig_easset,
                width="stretch",
                key="easset_bar_chart",
                on_select=handle_easset_chart_selection,
                selection_mode="points",
                config={"displayModeBar": False},
            )


# PAGE 2: Laporan
with st.container(key="page_two"):
    # Apabila bar PTJ diklik, semua laporan di bawah akan ikut PTJ tersebut.
    report_ptj = st.session_state.get("chart_selected_ptj")
    report_source = st.session_state.get("chart_selected_source")

    sap_report_data = sap_missing.copy()
    easset_report_data = easset_missing.copy()
    lokasi_report_data = lokasi.copy()

    if report_ptj:
        sap_report_data = sap_report_data[
            sap_report_data["PTJ SAP"].fillna("Tidak Dikenal Pasti").eq(report_ptj)
        ]
        easset_report_data = easset_report_data[
            easset_report_data["PTJ eAsset"].fillna("Tidak Dikenal Pasti").eq(report_ptj)
        ]
        lokasi_report_data = lokasi_report_data[
            lokasi_report_data["PTJ SAP"].fillna("Tidak Dikenal Pasti").eq(report_ptj)
            | lokasi_report_data["PTJ eAsset"].fillna("Tidak Dikenal Pasti").eq(report_ptj)
        ]

    st.markdown("### Laporan")
    kpi_report = st.session_state.get("kpi_selected_report")

    if report_ptj or kpi_report:
        info_col, reset_col = st.columns([5, 1])
        if report_ptj:
            info_col.info(f"Laporan ditapis mengikut bar {report_source}: {report_ptj}")
        elif kpi_report:
            info_col.info(f"Laporan dipilih melalui KPI: {kpi_report}")
        if reset_col.button("Papar Semua", use_container_width=True):
            clear_chart_selection()
            st.rerun()

    sap_report_cols = [
        "No Aset",
        "Nama Aset",
        "Deskripsi Ringkas",
        "Eval Group SAP",
        "PTJ SAP",
        "Sub No",
        "Nilai Perolehan SAP",
        "Nilai Buku SAP",
    ]
    sap_report = sap_report_data[[c for c in sap_report_cols if c in sap_report_data.columns]].copy()

    easset_report_cols = [
        "No Aset",
        "No. Siri Pendaftaran",
        "Jenis",
        "Jenama",
        "Lokasi",
        "Pegawai Penempatan",
        "No. Pesanan",
        "Tarikh Beli",
        "Nilai eAsset",
        "Eval Group eAsset",
        "PTJ eAsset",
    ]
    easset_report = easset_report_data[[c for c in easset_report_cols if c in easset_report_data.columns]].copy()

    if kpi_report == "SAP":
        st.subheader("SAP")
        st.dataframe(sap_report, use_container_width=True, hide_index=True, height=500)
        st.download_button(
            "Muat Turun Laporan SAP",
            data=excel_bytes({"SAP tiada di eAsset": sap_report}),
            file_name="laporan_aset_sap_tiada_di_easset.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    elif kpi_report == "eAsset":
        st.subheader("eAsset")
        st.dataframe(easset_report, use_container_width=True, hide_index=True, height=500)
        st.download_button(
            "Muat Turun Laporan eAsset",
            data=excel_bytes({"eAsset tiada di SAP": easset_report}),
            file_name="laporan_aset_easset_tiada_di_sap.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    elif kpi_report == "Lokasi Berlainan":
        st.subheader("Lokasi Berlainan")
        st.dataframe(lokasi_report_data, use_container_width=True, hide_index=True, height=500)
        st.download_button(
            "Muat Turun Laporan Lokasi Berlainan",
            data=excel_bytes({"Lokasi Berlainan": lokasi_report_data}),
            file_name="laporan_aset_lokasi_berlainan.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        tab_sap, tab_easset, tab_lokasi = st.tabs(["SAP", "eAsset", "Lokasi Berlainan"])

        with tab_sap:
            st.dataframe(sap_report, use_container_width=True, hide_index=True, height=500)
            st.download_button(
                "Muat Turun Laporan SAP",
                data=excel_bytes({"SAP tiada di eAsset": sap_report}),
                file_name="laporan_aset_sap_tiada_di_easset.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with tab_easset:
            st.dataframe(easset_report, use_container_width=True, hide_index=True, height=500)
            st.download_button(
                "Muat Turun Laporan eAsset",
                data=excel_bytes({"eAsset tiada di SAP": easset_report}),
                file_name="laporan_aset_easset_tiada_di_sap.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        with tab_lokasi:
            st.dataframe(lokasi_report_data, use_container_width=True, hide_index=True, height=500)
            st.download_button(
                "Muat Turun Laporan Lokasi Berlainan",
                data=excel_bytes({"Lokasi Berlainan": lokasi_report_data}),
                file_name="laporan_aset_lokasi_berlainan.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    st.download_button(
        "Muat Turun Semua Laporan",
        data=excel_bytes(
            {
                "SAP tiada di eAsset": sap_report,
                "eAsset tiada di SAP": easset_report,
                "Lokasi Berlainan": lokasi_report_data,
            }
        ),
        file_name="laporan_perbandingan_aset.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
