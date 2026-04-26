import streamlit as st
import pandas as pd
import plotly.express as px

from nobet_engine_104_eczane import run_schedule, create_groups, normalize_name


# =========================================================
# SAYFA AYARI
# =========================================================

st.set_page_config(
    page_title="Eczane Nöbet Planlayıcı",
    page_icon="💊",
    layout="wide",
)


# =========================================================
# CSS TASARIM
# =========================================================

st.markdown(
    """
    <style>
    .main {
        background-color: #f7f9fc;
    }

    .hero-box {
        background: linear-gradient(135deg, #123c69, #1f6feb);
        padding: 32px;
        border-radius: 24px;
        color: white;
        margin-bottom: 24px;
        box-shadow: 0 12px 32px rgba(18, 60, 105, 0.18);
    }

    .hero-title {
        font-size: 34px;
        font-weight: 800;
        margin-bottom: 8px;
    }

    .hero-subtitle {
        font-size: 17px;
        opacity: 0.92;
        max-width: 900px;
    }

    .metric-card {
        background: white;
        border-radius: 20px;
        padding: 22px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.07);
        border: 1px solid #e8eef7;
    }

    .metric-label {
        font-size: 14px;
        color: #64748b;
        margin-bottom: 6px;
    }

    .metric-value {
        font-size: 28px;
        font-weight: 800;
        color: #123c69;
    }

    .section-title {
        font-size: 22px;
        font-weight: 800;
        color: #0f172a;
        margin-top: 18px;
        margin-bottom: 12px;
    }

    .info-card {
        background: white;
        border-radius: 18px;
        padding: 18px;
        border: 1px solid #e8eef7;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
    }

    .soft-warning {
        background: #fff7ed;
        color: #9a3412;
        border: 1px solid #fed7aa;
        padding: 14px 16px;
        border-radius: 16px;
        font-weight: 600;
    }

    .soft-success {
        background: #ecfdf5;
        color: #065f46;
        border: 1px solid #a7f3d0;
        padding: 14px 16px;
        border-radius: 16px;
        font-weight: 600;
    }

    div[data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e8eef7;
    }

    .stButton button {
        border-radius: 14px;
        background-color: #123c69;
        color: white;
        border: none;
        padding: 10px 18px;
        font-weight: 700;
    }

    .stDownloadButton button {
        border-radius: 14px;
        background-color: #123c69;
        color: white;
        border: none;
        padding: 10px 18px;
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# YARDIMCI FONKSİYONLAR
# =========================================================

@st.cache_data(show_spinner=False)
def read_excel_file(uploaded_file):
    return pd.read_excel(uploaded_file)


def create_empty_history_df():
    """
    Geçmiş Excel yüklenmezse motor kodundaki tüm eczaneleri sıfır geçmişle başlatır.
    Bu sayede plan tamamen sıfırdan hesaplanır.
    """
    groups = create_groups()
    rows = []
    seen = set()

    for _, pharmacies in groups.items():
        for pharmacy in pharmacies:
            pharmacy_name = normalize_name(pharmacy)
            if pharmacy_name in seen:
                continue
            seen.add(pharmacy_name)
            rows.append({
                "Eczane": pharmacy_name,
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


def load_generated_files(plan_file="Alternatif.xlsx", detail_file="aylik_nobet_data.xlsx"):
    plan_sheets = pd.read_excel(plan_file, sheet_name=None)
    detail_df = pd.read_excel(detail_file)
    return plan_sheets, detail_df


def prepare_long_schedule(plan_sheets):
    rows = []

    for sheet_name, df in plan_sheets.items():
        if sheet_name.upper() == "GENEL OZET":
            continue
        if "Tarih" not in df.columns or "Gün" not in df.columns:
            continue

        group_cols = [c for c in df.columns if c not in ["Tarih", "Gün"]]

        for _, r in df.iterrows():
            for g in group_cols:
                eczane = r.get(g)
                if pd.notna(eczane) and str(eczane).strip():
                    rows.append({
                        "Sayfa": sheet_name,
                        "Tarih": r["Tarih"],
                        "Gün": r["Gün"],
                        "Grup": g,
                        "Eczane": str(eczane).strip(),
                    })

    return pd.DataFrame(rows)


def get_download_bytes(path):
    with open(path, "rb") as f:
        return f.read()


# =========================================================
# HERO
# =========================================================

st.markdown(
    """
    <div class="hero-box">
        <div class="hero-title">💊 Eczane Nöbet Planlayıcı</div>
        <div class="hero-subtitle">
            100+ eczane için grup bazlı, hafta içi / hafta sonu dengeli, bayram ve arefe kurallarını dikkate alan modern nöbet planlama paneli.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:
    st.header("⚙️ Plan Ayarları")

    yil = st.number_input("Yıl", min_value=2025, max_value=2035, value=2026, step=1)
    ay = st.selectbox(
        "Başlangıç Ayı",
        options=list(range(1, 13)),
        index=0,
        format_func=lambda x: f"{x:02d}",
    )
    ay_sayisi = st.number_input("Kaç Ay Oluşturulsun?", min_value=1, max_value=24, value=1, step=1)

    st.divider()

    st.subheader("📂 Excel Dosyaları")
    gecmis_yuk_file = st.file_uploader(
        "Geçmiş Nöbet Yükü Excel",
        type=["xlsx"],
        help="Boş bırakılırsa tüm eczaneler sıfır geçmişle planlanır.",
    )
    gecmis_bayram_file = st.file_uploader(
        "Geçmiş Bayram / Arefe Excel",
        type=["xlsx"],
        help="Opsiyonel. Boş bırakılırsa bayram geçmişi yok kabul edilir.",
    )

    st.divider()
    st.caption("Excel yüklemezseniz sistem bütün eczaneleri sıfır nöbet geçmişiyle başlatır.")


# =========================================================
# ANA ALAN
# =========================================================

tab_plan, tab_ozet, tab_grafik, tab_detay = st.tabs([
    "🚀 Plan Oluştur",
    "📌 Genel Özet",
    "📊 Grafikler",
    "📅 Aylık Detay",
])


# =========================================================
# TAB 1 - PLAN OLUŞTUR
# =========================================================

with tab_plan:
    st.markdown('<div class="section-title">Plan Oluşturma</div>', unsafe_allow_html=True)

    c1, c2, c3 = st.columns(3)

    with c1:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Seçilen Yıl</div>
                <div class="metric-value">{yil}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c2:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Başlangıç Ayı</div>
                <div class="metric-value">{ay:02d}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with c3:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">Planlanacak Ay</div>
                <div class="metric-value">{ay_sayisi}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.write("")

    if gecmis_yuk_file is None:
        gecmis_yuk_df = create_empty_history_df()
        gecmis_bayram_df = None

        st.markdown(
            """
            <div class="soft-warning">
                Geçmiş nöbet dosyası yüklenmedi. Sistem tüm eczaneleri 0 geçmişle başlatacak.
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.expander("Sıfırdan oluşturulan geçmiş veri ön izle"):
            st.dataframe(gecmis_yuk_df, use_container_width=True)

    else:
        gecmis_yuk_df = read_excel_file(gecmis_yuk_file)
        gecmis_bayram_df = read_excel_file(gecmis_bayram_file) if gecmis_bayram_file is not None else None

        st.markdown(
            """
            <div class="soft-success">
                Geçmiş nöbet dosyası yüklendi. Plan bu geçmiş veriye göre oluşturulacak.
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.expander("Yüklenen geçmiş nöbet dosyasını ön izle"):
            st.dataframe(gecmis_yuk_df, use_container_width=True)

        if gecmis_bayram_df is not None:
            with st.expander("Yüklenen geçmiş bayram / arefe dosyasını ön izle"):
                st.dataframe(gecmis_bayram_df, use_container_width=True)

    if st.button("✅ Nöbet Planını Oluştur", use_container_width=True):
        with st.spinner("Plan oluşturuluyor..."):
            plan_file, detail_file = run_schedule(
                y=int(yil),
                m=int(ay),
                nm=int(ay_sayisi),
                gecmis_yuk_df=gecmis_yuk_df,
                gecmis_bayram_df=gecmis_bayram_df,
            )

        st.session_state["plan_file"] = plan_file
        st.session_state["detail_file"] = detail_file
        st.success("Plan başarıyla oluşturuldu.")

    if "plan_file" in st.session_state:
        st.markdown('<div class="section-title">Dosya İndirme</div>', unsafe_allow_html=True)
        d1, d2 = st.columns(2)

        with d1:
            st.download_button(
                "📥 Plan Excel Dosyasını İndir",
                data=get_download_bytes(st.session_state["plan_file"]),
                file_name=st.session_state["plan_file"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with d2:
            st.download_button(
                "📥 Aylık Detay Dosyasını İndir",
                data=get_download_bytes(st.session_state["detail_file"]),
                file_name=st.session_state["detail_file"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )


# =========================================================
# ORTAK VERİ YÜKLEME
# =========================================================

plan_ready = "plan_file" in st.session_state and "detail_file" in st.session_state

if plan_ready:
    plan_sheets, detail_df = load_generated_files(st.session_state["plan_file"], st.session_state["detail_file"])
    summary_df = plan_sheets.get("GENEL OZET", pd.DataFrame())
    long_schedule_df = prepare_long_schedule(plan_sheets)
else:
    plan_sheets = {}
    detail_df = pd.DataFrame()
    summary_df = pd.DataFrame()
    long_schedule_df = pd.DataFrame()


# =========================================================
# TAB 2 - GENEL ÖZET
# =========================================================

with tab_ozet:
    st.markdown('<div class="section-title">Genel Özet</div>', unsafe_allow_html=True)

    if not plan_ready:
        st.warning("Özet ekranı için önce plan oluşturun.")
    else:
        toplam_eczane = summary_df["Eczane"].nunique() if "Eczane" in summary_df.columns else 0
        toplam_nobet = summary_df["Toplam Nöbet"].sum() if "Toplam Nöbet" in summary_df.columns else 0
        toplam_bayram = summary_df["Bayram"].sum() if "Bayram" in summary_df.columns else 0
        ort_katsayi = summary_df["Toplam Katsayı"].mean() if "Toplam Katsayı" in summary_df.columns else 0

        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.metric("Toplam Eczane", int(toplam_eczane))
        with c2:
            st.metric("Toplam Nöbet", int(toplam_nobet))
        with c3:
            st.metric("Toplam Bayram", int(toplam_bayram))
        with c4:
            st.metric("Ortalama Katsayı", round(float(ort_katsayi), 2))

        st.dataframe(summary_df, use_container_width=True, height=520)


# =========================================================
# TAB 3 - GRAFİKLER
# =========================================================

with tab_grafik:
    st.markdown('<div class="section-title">Grafik Analizi</div>', unsafe_allow_html=True)

    if not plan_ready:
        st.warning("Grafikler için önce plan oluşturun.")
    else:
        g1, g2 = st.columns(2)

        if "Toplam Katsayı" in summary_df.columns:
            fig1 = px.bar(
                summary_df.sort_values("Toplam Katsayı", ascending=False),
                x="Eczane",
                y="Toplam Katsayı",
                color="Grup" if "Grup" in summary_df.columns else None,
                title="Eczane Bazlı Toplam Katsayı",
            )
            fig1.update_layout(xaxis_tickangle=-60, height=520)
            g1.plotly_chart(fig1, use_container_width=True)

        if "Bayram" in summary_df.columns:
            fig2 = px.bar(
                summary_df.sort_values("Bayram", ascending=False),
                x="Eczane",
                y="Bayram",
                color="Grup" if "Grup" in summary_df.columns else None,
                title="Bayram Nöbet Dağılımı",
            )
            fig2.update_layout(xaxis_tickangle=-60, height=520)
            g2.plotly_chart(fig2, use_container_width=True)

        st.markdown('<div class="section-title">Hafta İçi / Hafta Sonu Dengesi</div>', unsafe_allow_html=True)

        needed_cols = ["Pzt", "Salı", "Çarş", "Perş", "Cuma", "Ctesi", "Pazar"]
        if all(c in summary_df.columns for c in needed_cols):
            tmp = summary_df.copy()
            tmp["Hafta İçi"] = tmp[["Pzt", "Salı", "Çarş", "Perş", "Cuma"]].sum(axis=1)
            tmp["Hafta Sonu"] = tmp[["Ctesi", "Pazar"]].sum(axis=1)

            balance_df = tmp[["Eczane", "Grup", "Hafta İçi", "Hafta Sonu"]].melt(
                id_vars=["Eczane", "Grup"],
                value_vars=["Hafta İçi", "Hafta Sonu"],
                var_name="Tür",
                value_name="Adet",
            )

            fig3 = px.bar(
                balance_df,
                x="Eczane",
                y="Adet",
                color="Tür",
                title="Hafta İçi / Hafta Sonu Dağılımı",
                barmode="group",
            )
            fig3.update_layout(xaxis_tickangle=-60, height=560)
            st.plotly_chart(fig3, use_container_width=True)


# =========================================================
# TAB 4 - AYLIK DETAY
# =========================================================

with tab_detay:
    st.markdown('<div class="section-title">Aylık Takvim ve Detay</div>', unsafe_allow_html=True)

    if not plan_ready:
        st.warning("Aylık detay için önce plan oluşturun.")
    else:
        month_sheets = [s for s in plan_sheets.keys() if s.upper() != "GENEL OZET"]
        selected_sheet = st.selectbox("Ay Seç", month_sheets)

        if selected_sheet:
            st.dataframe(plan_sheets[selected_sheet], use_container_width=True, height=520)

        st.markdown('<div class="section-title">Eczane Bazlı Aylık Detay</div>', unsafe_allow_html=True)

        if not detail_df.empty:
            eczane_list = sorted(detail_df["Eczane"].dropna().unique()) if "Eczane" in detail_df.columns else []
            selected_eczane = st.selectbox("Eczane Seç", ["Tümü"] + eczane_list)

            filtered = detail_df.copy()
            if selected_eczane != "Tümü":
                filtered = filtered[filtered["Eczane"] == selected_eczane]

            st.dataframe(filtered, use_container_width=True, height=420)

            day_cols = ["Pazartesi", "Salı", "Çarşamba", "Perşembe", "Cuma", "Cumartesi", "Pazar"]
            if all(c in filtered.columns for c in day_cols):
                chart_df = filtered.melt(
                    id_vars=["Eczane", "Yıl", "Ay"],
                    value_vars=day_cols,
                    var_name="Gün",
                    value_name="Adet",
                )

                fig4 = px.bar(
                    chart_df,
                    x="Gün",
                    y="Adet",
                    color="Eczane" if selected_eczane == "Tümü" else None,
                    title="Gün Bazlı Aylık Dağılım",
                )
                st.plotly_chart(fig4, use_container_width=True)
