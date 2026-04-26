import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

try:
    import plotly.express as px
except Exception:
    px = None

from nobet_engine import run_schedule


st.set_page_config(
    page_title="AYÇA | Eczane Nöbet Planlayıcı",
    page_icon="💊",
    layout="wide"
)


DEFAULT_GROUPS = {
    "A1": [f"ECZANE{i}" for i in range(1, 10)],
    "A2": [f"ECZANE{i}" for i in range(10, 19)],
    "A3": [f"ECZANE{i}" for i in range(19, 28)],

    "B1": [f"ECZANE{i}" for i in range(28, 36)],
    "B2": [f"ECZANE{i}" for i in range(36, 45)],
    "B3": [f"ECZANE{i}" for i in range(45, 53)],

    "C1": [f"ECZANE{i}" for i in range(53, 61)],
    "C2": [f"ECZANE{i}" for i in range(61, 70)],
    "C3": [f"ECZANE{i}" for i in range(70, 79)],

    "D1": [f"ECZANE{i}" for i in range(79, 87)],
    "D2": [f"ECZANE{i}" for i in range(87, 96)],
    "D3": [f"ECZANE{i}" for i in range(96, 105)],
}


def create_empty_history_df():
    rows = []

    for _, eczaneler in DEFAULT_GROUPS.items():
        for eczane in eczaneler:
            rows.append({
                "Eczane": eczane,
                "Bayram": 0,
                "Pzt": 0,
                "Salı": 0,
                "Çarş": 0,
                "Perş": 0,
                "Cuma": 0,
                "Ctesi": 0,
                "Pazar": 0,
            })

    return pd.DataFrame(rows)


def read_uploaded_excel(uploaded_file):
    uploaded_file.seek(0)
    xls = pd.ExcelFile(uploaded_file)

    ana_sekme = None
    for sheet in xls.sheet_names:
        if sheet.strip().upper() != "GECMIS_BAYRAM":
            ana_sekme = sheet
            break

    if ana_sekme is None:
        raise Exception("Ana geçmiş yük sekmesi bulunamadı.")

    gecmis_yuk_df = pd.read_excel(xls, sheet_name=ana_sekme)

    gecmis_bayram_df = None
    if "GECMIS_BAYRAM" in xls.sheet_names:
        gecmis_bayram_df = pd.read_excel(xls, sheet_name="GECMIS_BAYRAM")

    return gecmis_yuk_df, gecmis_bayram_df, xls.sheet_names, ana_sekme


def get_download_bytes(path):
    with open(path, "rb") as f:
        return f.read()


st.markdown("""
<style>
.stApp {
    background: linear-gradient(180deg, #f6f8fb 0%, #eef3f9 100%);
}

.block-container {
    padding-top: 1.5rem;
    max-width: 1250px;
}

.hero {
    background: linear-gradient(135deg, #0f172a 0%, #1f4b99 55%, #0ea5e9 100%);
    color: white;
    padding: 32px;
    border-radius: 26px;
    box-shadow: 0 18px 42px rgba(15, 23, 42, 0.18);
    margin-bottom: 24px;
}

.hero h1 {
    font-size: 36px;
    margin-bottom: 8px;
    font-weight: 850;
}

.hero p {
    font-size: 16px;
    opacity: 0.9;
    max-width: 850px;
}

.card {
    background: white;
    border-radius: 22px;
    padding: 22px;
    border: 1px solid #e2e8f0;
    box-shadow: 0 10px 26px rgba(15, 23, 42, 0.07);
    margin-bottom: 18px;
}

.metric-card {
    background: white;
    border-radius: 20px;
    padding: 20px;
    border: 1px solid #e2e8f0;
    box-shadow: 0 8px 22px rgba(15, 23, 42, 0.06);
}

.metric-label {
    color: #64748b;
    font-size: 14px;
    font-weight: 700;
}

.metric-value {
    color: #0f172a;
    font-size: 28px;
    font-weight: 850;
    margin-top: 6px;
}

.info-zero {
    background: #fff7ed;
    color: #9a3412;
    border: 1px solid #fed7aa;
    padding: 14px 16px;
    border-radius: 16px;
    font-weight: 650;
}

.info-ok {
    background: #ecfdf5;
    color: #065f46;
    border: 1px solid #a7f3d0;
    padding: 14px 16px;
    border-radius: 16px;
    font-weight: 650;
}

div.stButton > button {
    border-radius: 14px;
    font-weight: 800;
    background: linear-gradient(90deg, #0f172a 0%, #0ea5e9 100%);
    color: white;
    border: none;
}

div.stDownloadButton > button {
    border-radius: 14px;
    font-weight: 800;
    background: linear-gradient(90deg, #0f172a 0%, #10b981 100%);
    color: white;
    border: none;
}
</style>
""", unsafe_allow_html=True)


st.markdown("""
<div class="hero">
    <h1>💊 AYÇA Nöbet Planlama Paneli</h1>
    <p>
        100+ eczane için geçmiş yük, hafta içi / hafta sonu dengesi,
        bayram ve arefe kurallarını dikkate alan modern nöbet planlama sistemi.
    </p>
</div>
""", unsafe_allow_html=True)


with st.sidebar:
    st.header("⚙️ Plan Ayarları")

    yil = st.number_input(
        "Yıl",
        min_value=2025,
        max_value=2035,
        value=2026,
        step=1
    )

    ay = st.selectbox(
        "Başlangıç Ayı",
        list(range(1, 13)),
        index=0,
        format_func=lambda x: f"{x:02d}"
    )

    kac_ay = st.number_input(
        "Kaç Ay Planlansın?",
        min_value=1,
        max_value=24,
        value=1,
        step=1
    )

    st.divider()

    uploaded_file = st.file_uploader(
        "Geçmiş Nöbet Excel",
        type=["xlsx"],
        help="Yüklemezsen sistem tüm eczaneleri sıfır geçmişle başlatır."
    )

    st.caption("Excel yüklenmezse tüm eczaneler 0 geçmişle hesaplanır.")

    st.divider()

    planla = st.button("🚀 Plan Oluştur", use_container_width=True)


tab1, tab2, tab3, tab4 = st.tabs([
    "🚀 Plan",
    "📌 Özet",
    "📊 Grafik",
    "📅 Takvim"
])


with tab1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Plan Oluşturma")

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Plan Yılı</div>
                <div class="metric-value">{yil}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c2:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Başlangıç Ayı</div>
                <div class="metric-value">{ay:02d}</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with c3:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Plan Süresi</div>
                <div class="metric-value">{kac_ay} ay</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.write("")

    if uploaded_file is None:
        gecmis_yuk_df = create_empty_history_df()
        gecmis_bayram_df = None

        st.markdown(
            """
            <div class="info-zero">
                Geçmiş dosya yüklenmedi. Sistem tüm eczaneleri sıfır geçmişle başlatacak.
            </div>
            """,
            unsafe_allow_html=True
        )

        with st.expander("Sıfır geçmiş verisi ön izle"):
            st.dataframe(gecmis_yuk_df, use_container_width=True)

    else:
        try:
            gecmis_yuk_df, gecmis_bayram_df, sheet_names, ana_sekme = read_uploaded_excel(uploaded_file)

            st.markdown(
                f"""
                <div class="info-ok">
                    Excel başarıyla okundu. Ana sekme: {ana_sekme}
                </div>
                """,
                unsafe_allow_html=True
            )

            st.write("Yüklenen sekmeler:", ", ".join(sheet_names))

            with st.expander("Geçmiş nöbet dosyası ön izle"):
                st.dataframe(gecmis_yuk_df, use_container_width=True)

            if gecmis_bayram_df is not None:
                with st.expander("GECMIS_BAYRAM ön izle"):
                    st.dataframe(gecmis_bayram_df, use_container_width=True)

        except Exception as e:
            st.error(f"Excel okunamadı: {e}")
            st.stop()

    if planla:
        with st.spinner("Plan oluşturuluyor..."):
            try:
                plan_file, detail_file = run_schedule(
                    y=int(yil),
                    m=int(ay),
                    nm=int(kac_ay),
                    gecmis_yuk_df=gecmis_yuk_df,
                    gecmis_bayram_df=gecmis_bayram_df,
                    eklenme_input={},
                    cikma_input={}
                )

                st.session_state.plan_file = plan_file
                st.session_state.detail_file = detail_file
                st.success("Plan başarıyla oluşturuldu.")

            except TypeError:
                plan_file, detail_file = run_schedule(
                    int(yil),
                    int(ay),
                    int(kac_ay),
                    gecmis_yuk_df,
                    gecmis_bayram_df
                )

                st.session_state.plan_file = plan_file
                st.session_state.detail_file = detail_file
                st.success("Plan başarıyla oluşturuldu.")

            except Exception as e:
                st.exception(e)

    if "plan_file" in st.session_state:
        st.write("")
        d1, d2 = st.columns(2)

        with d1:
            st.download_button(
                "📥 Nöbet Planını İndir",
                data=get_download_bytes(st.session_state.plan_file),
                file_name="nobet_plani.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        with d2:
            st.download_button(
                "📊 Aylık Detay İndir",
                data=get_download_bytes(st.session_state.detail_file),
                file_name="aylik_detay.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

    st.markdown("</div>", unsafe_allow_html=True)


plan_ready = "plan_file" in st.session_state and "detail_file" in st.session_state

if plan_ready:
    try:
        plan_sheets = pd.read_excel(st.session_state.plan_file, sheet_name=None)
        summary_df = plan_sheets.get("GENEL OZET", pd.DataFrame())
        detail_df = pd.read_excel(st.session_state.detail_file)
    except Exception:
        plan_sheets = {}
        summary_df = pd.DataFrame()
        detail_df = pd.DataFrame()
else:
    plan_sheets = {}
    summary_df = pd.DataFrame()
    detail_df = pd.DataFrame()


with tab2:
    st.subheader("📌 Genel Özet")

    if not plan_ready:
        st.warning("Özet için önce plan oluştur.")
    else:
        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.metric(
                "Toplam Eczane",
                summary_df["Eczane"].nunique() if "Eczane" in summary_df.columns else 0
            )

        with c2:
            st.metric(
                "Toplam Nöbet",
                int(summary_df["Toplam Nöbet"].sum()) if "Toplam Nöbet" in summary_df.columns else 0
            )

        with c3:
            st.metric(
                "Toplam Bayram",
                int(summary_df["Bayram"].sum()) if "Bayram" in summary_df.columns else 0
            )

        with c4:
            st.metric(
                "Ortalama Katsayı",
                round(float(summary_df["Toplam Katsayı"].mean()), 2) if "Toplam Katsayı" in summary_df.columns else 0
            )

        st.dataframe(summary_df, use_container_width=True, height=520)


with tab3:
    st.subheader("📊 Grafik Analizi")

    if not plan_ready:
        st.warning("Grafikler için önce plan oluştur.")
    elif px is None:
        st.warning("Plotly kurulu değil. requirements.txt içine plotly ekleyin.")
    else:
        if "Toplam Katsayı" in summary_df.columns:
            fig = px.bar(
                summary_df.sort_values("Toplam Katsayı", ascending=False),
                x="Eczane",
                y="Toplam Katsayı",
                color="Grup" if "Grup" in summary_df.columns else None,
                title="Eczane Bazlı Toplam Katsayı"
            )
            fig.update_layout(xaxis_tickangle=-60, height=520)
            st.plotly_chart(fig, use_container_width=True)

        if "Bayram" in summary_df.columns:
            fig2 = px.bar(
                summary_df.sort_values("Bayram", ascending=False),
                x="Eczane",
                y="Bayram",
                color="Grup" if "Grup" in summary_df.columns else None,
                title="Bayram Nöbet Dağılımı"
            )
            fig2.update_layout(xaxis_tickangle=-60, height=480)
            st.plotly_chart(fig2, use_container_width=True)


with tab4:
    st.subheader("📅 Aylık Takvim")

    if not plan_ready:
        st.warning("Takvim için önce plan oluştur.")
    else:
        month_sheets = [
            s for s in plan_sheets.keys()
            if s.upper() != "GENEL OZET"
        ]

        if month_sheets:
            selected_sheet = st.selectbox("Ay seç", month_sheets)
            st.dataframe(
                plan_sheets[selected_sheet],
                use_container_width=True,
                height=520
            )

        st.subheader("Aylık Detay")

        if not detail_df.empty:
            st.dataframe(detail_df, use_container_width=True, height=420)
