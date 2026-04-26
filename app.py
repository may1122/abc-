import streamlit as st
from nobet_engine import run_schedule
import datetime
import pandas as pd
import base64
from pathlib import Path

try:
    import plotly.express as px
except Exception:
    px = None


# ==============================
# SAYFA AYARI
# ==============================

st.set_page_config(
    page_title="AYÇA | Eczane Nöbet Planlayıcı",
    page_icon="💊",
    layout="wide",
)


# ==============================
# YARDIMCI FONKSİYONLAR
# ==============================

def get_base64_image(image_path: str):
    path = Path(image_path)
    if not path.exists():
        return None
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


def read_excel_safely(uploaded_file):
    uploaded_file.seek(0)
    return pd.ExcelFile(uploaded_file)


def find_main_sheet(sheet_names):
    for s in sheet_names:
        if s.strip().upper() != "GECMIS_BAYRAM":
            return s
    return None


def read_generated_excel(file_bytes):
    try:
        return pd.ExcelFile(file_bytes)
    except Exception:
        return None


def create_metric_card(title, value, note=""):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-title">{title}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


logo_base64 = get_base64_image("logo.png")


# ==============================
# CSS / TASARIM
# ==============================

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {
        font-family: 'Inter', sans-serif;
    }

    .stApp {
        background:
            radial-gradient(circle at top left, rgba(14,165,233,0.14), transparent 30%),
            radial-gradient(circle at top right, rgba(16,185,129,0.10), transparent 28%),
            linear-gradient(180deg, #f6f8fb 0%, #eef3f9 100%);
    }

    .block-container {
        padding-top: 1.4rem;
        padding-bottom: 2rem;
        max-width: 1280px;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
        border-right: 1px solid rgba(15,23,42,0.08);
    }

    .sidebar-logo-box {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 20px;
        padding: 16px;
        margin-bottom: 18px;
        text-align: center;
        box-shadow: 0 10px 24px rgba(15, 23, 42, 0.06);
    }

    .sidebar-logo-box img {
        max-width: 180px;
        width: 100%;
        height: auto;
    }

    .hero-card {
        background: linear-gradient(135deg, rgba(15,23,42,0.96), rgba(14,165,233,0.88));
        border: 1px solid rgba(255,255,255,0.16);
        border-radius: 28px;
        padding: 30px 34px;
        box-shadow: 0 18px 45px rgba(15, 23, 42, 0.18);
        margin-bottom: 22px;
        color: white;
        overflow: hidden;
        position: relative;
    }

    .hero-card:after {
        content: "";
        position: absolute;
        right: -80px;
        top: -80px;
        width: 240px;
        height: 240px;
        background: rgba(255,255,255,0.12);
        border-radius: 999px;
    }

    .hero-grid {
        position: relative;
        z-index: 1;
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 28px;
        flex-wrap: wrap;
    }

    .hero-badge {
        display: inline-flex;
        align-items: center;
        gap: 8px;
        background: rgba(255,255,255,0.14);
        color: white;
        font-size: 0.86rem;
        font-weight: 800;
        padding: 9px 14px;
        border-radius: 999px;
        margin-bottom: 14px;
        border: 1px solid rgba(255,255,255,0.18);
    }

    .hero-title {
        font-size: 2.25rem;
        font-weight: 850;
        margin: 0 0 8px 0;
        letter-spacing: -0.035em;
    }

    .hero-sub {
        font-size: 1rem;
        color: rgba(255,255,255,0.88);
        margin: 0;
        line-height: 1.65;
        max-width: 720px;
    }

    .hero-logo {
        background: rgba(255,255,255,0.94);
        border-radius: 24px;
        padding: 16px 20px;
        min-width: 230px;
        text-align: center;
        box-shadow: 0 14px 34px rgba(15, 23, 42, 0.14);
    }

    .hero-logo img {
        max-width: 230px;
        width: 100%;
        height: auto;
    }

    .metric-card {
        background: rgba(255,255,255,0.94);
        border: 1px solid rgba(15, 23, 42, 0.07);
        border-radius: 22px;
        padding: 20px 20px;
        box-shadow: 0 12px 28px rgba(15, 23, 42, 0.07);
        min-height: 120px;
    }

    .metric-title {
        color: #64748b;
        font-size: 0.92rem;
        font-weight: 700;
        margin-bottom: 8px;
    }

    .metric-value {
        color: #0f172a;
        font-size: 1.95rem;
        font-weight: 850;
        letter-spacing: -0.02em;
    }

    .metric-note {
        color: #64748b;
        font-size: 0.84rem;
        margin-top: 4px;
    }

    .section-card {
        background: rgba(255,255,255,0.94);
        border: 1px solid rgba(15, 23, 42, 0.07);
        border-radius: 24px;
        padding: 24px;
        box-shadow: 0 12px 30px rgba(15, 23, 42, 0.065);
        margin-bottom: 20px;
    }

    .section-title {
        font-size: 1.18rem;
        font-weight: 850;
        color: #0f172a;
        margin-bottom: 8px;
    }

    .section-text {
        color: #475569;
        font-size: 0.96rem;
        line-height: 1.6;
        margin-bottom: 14px;
    }

    .mini-info {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 14px 16px;
        color: #334155;
        font-size: 0.93rem;
        line-height: 1.55;
    }

    .success-card {
        background: linear-gradient(135deg, rgba(16,185,129,0.14), rgba(14,165,233,0.10));
        border: 1px solid rgba(16,185,129,0.28);
        border-radius: 22px;
        padding: 20px;
        color: #064e3b;
        font-weight: 750;
        margin-bottom: 16px;
    }

    div[data-testid="stFileUploader"] {
        background: white;
        border: 1px dashed #94a3b8;
        border-radius: 18px;
        padding: 12px;
    }

    div.stButton > button {
        border-radius: 16px;
        font-weight: 850;
        border: none;
        background: linear-gradient(90deg, #0f172a 0%, #0ea5e9 100%);
        color: white;
        padding: 0.78rem 1rem;
        box-shadow: 0 10px 24px rgba(14, 165, 233, 0.25);
    }

    div.stDownloadButton > button {
        border-radius: 16px;
        font-weight: 850;
        border: none;
        background: linear-gradient(90deg, #0f172a 0%, #10b981 100%);
        color: white;
        padding: 0.78rem 1rem;
        box-shadow: 0 10px 24px rgba(16, 185, 129, 0.22);
        width: 100%;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 18px;
        overflow: hidden;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ==============================
# SIDEBAR
# ==============================

with st.sidebar:
    if logo_base64:
        st.markdown(
            f"""
            <div class="sidebar-logo-box">
                <img src="data:image/png;base64,{logo_base64}" alt="AYÇA Logo">
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("## 📅 Plan Parametreleri")

    yil = st.number_input(
        "Yıl",
        min_value=2025,
        max_value=2035,
        value=datetime.datetime.now().year,
        step=1,
    )

    ay = st.number_input(
        "Başlangıç Ayı",
        min_value=1,
        max_value=12,
        value=1,
        step=1,
    )

    kac_ay = st.number_input(
        "Kaç Ay Planlansın",
        min_value=1,
        max_value=24,
        value=3,
        step=1,
    )

    st.markdown(
        """
        <div class="mini-info">
            Planlama motoru; geçmiş yük, hafta sonu dengesi, bayram geçmişi, arefe ve grup rotasyonunu birlikte değerlendirir.
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.divider()
    planla = st.button("🚀 Plan Oluştur", use_container_width=True)


# ==============================
# HERO
# ==============================

logo_html = ""
if logo_base64:
    logo_html = f"""
    <div class="hero-logo">
        <img src="data:image/png;base64,{logo_base64}" alt="AYÇA Logo">
    </div>
    """

st.markdown(
    f"""
    <div class="hero-card">
        <div class="hero-grid">
            <div style="flex:1; min-width:320px;">
                <div class="hero-badge">💊 AYÇA • Akıllı Yazılım Çözüm Asistanı</div>
                <h1 class="hero-title">Eczane Nöbet Planlayıcı</h1>
                <p class="hero-sub">
                    100+ eczane için geçmiş yük, bayram/arefe geçmişi ve dönemsel denge mantığıyla
                    daha adil, daha kontrollü ve daha okunabilir nöbet planları oluşturun.
                </p>
            </div>
            {logo_html}
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ==============================
# ÜST METRİKLER
# ==============================

m1, m2, m3, m4 = st.columns(4)
with m1:
    create_metric_card("Plan Yılı", yil, "Seçilen dönem")
with m2:
    create_metric_card("Başlangıç Ayı", f"{ay:02d}", "İlk oluşturulacak ay")
with m3:
    create_metric_card("Plan Süresi", f"{kac_ay} ay", "Ardışık planlama")
with m4:
    create_metric_card("Durum", "Hazır", "Excel yüklendiğinde planlanır")

st.write("")


# ==============================
# ANA SEKMELER
# ==============================

tab_veri, tab_degisim, tab_sonuc, tab_analiz = st.tabs(
    ["📂 Veri Yükle", "🔧 Değişiklikler", "📥 Sonuçlar", "📊 Analiz"]
)


# ==============================
# TAB 1 - VERİ YÜKLE
# ==============================

with tab_veri:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📂 Geçmiş Nöbet Dosyası</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-text">Ana geçmiş yük sekmesini ve varsa <b>GECMIS_BAYRAM</b> sekmesini içeren Excel dosyasını yükleyin.</div>',
        unsafe_allow_html=True,
    )

    uploaded_file = st.file_uploader(
        "Geçmiş nöbet Excel dosyasını yükleyin",
        type=["xlsx"],
        label_visibility="collapsed",
    )

    if uploaded_file:
        try:
            xls = read_excel_safely(uploaded_file)
            sheet_names = xls.sheet_names
            ana_sekme = find_main_sheet(sheet_names)

            if ana_sekme is None:
                st.error("Ana geçmiş yük sekmesi bulunamadı.")
                st.stop()

            st.success(f"Excel başarıyla okundu. Sekmeler: {', '.join(sheet_names)}")

            c1, c2, c3 = st.columns(3)
            with c1:
                st.metric("Ana Sekme", ana_sekme)
            with c2:
                st.metric("Toplam Sekme", len(sheet_names))
            with c3:
                st.metric("Bayram Sekmesi", "Var" if "GECMIS_BAYRAM" in sheet_names else "Yok")

            preview_df = pd.read_excel(xls, sheet_name=ana_sekme)
            with st.expander("👀 Ana geçmiş yük dosyası ön izleme", expanded=True):
                st.dataframe(preview_df.head(20), use_container_width=True)

            if "GECMIS_BAYRAM" in sheet_names:
                bayram_preview = pd.read_excel(xls, sheet_name="GECMIS_BAYRAM")
                with st.expander("🎉 GECMIS_BAYRAM ön izleme", expanded=False):
                    st.dataframe(bayram_preview.head(20), use_container_width=True)

        except Exception as e:
            st.error(f"Excel okunamadı: {e}")
            st.stop()
    else:
        st.info("Plan oluşturmak için önce Excel dosyasını yükleyin.")

    st.markdown('</div>', unsafe_allow_html=True)


# ==============================
# TAB 2 - ECZANE DEĞİŞİKLİKLERİ
# ==============================

with tab_degisim:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">🔧 Eczane Değişiklikleri</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-text">Yeni eczane ekleyebilir veya belirli bir tarihten sonra nöbete dahil edilmeyecek eczane tanımlayabilirsiniz.</div>',
        unsafe_allow_html=True,
    )

    degisim = st.toggle("Eczane ekleme / çıkarma yapılacak mı?")

    eklenme = {}
    cikma = {}

    if degisim:
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("### ➕ Eczane Ekle")
            eczane_ekle = st.text_input("Eklenecek Eczane İsmi")
            eczane_grup = st.selectbox(
                "Grup",
                [
                    "A1", "A2", "A3",
                    "B1", "B2", "B3",
                    "C1", "C2", "C3",
                    "D1", "D2", "D3",
                    "E1", "E2", "E3",
                    "F1", "F2", "F3",
                    "G1", "G2", "G3",
                ],
            )
            ekleme_tarihi = st.date_input("Eklenme Tarihi", value=datetime.date.today())

            if eczane_ekle:
                eklenme[eczane_ekle.upper()] = {
                    "tarih": ekleme_tarihi,
                    "grup": eczane_grup,
                }
                st.success(f"{eczane_ekle.upper()} {eczane_grup} grubuna eklenecek.")

        with col2:
            st.markdown("### ➖ Eczane Çıkar")
            eczane_cikar = st.text_input("Çıkarılacak Eczane İsmi", key="cikar")
            cikis_tarihi = st.date_input("Çıkış Tarihi", value=datetime.date.today())

            if eczane_cikar:
                cikma[eczane_cikar.upper()] = cikis_tarihi
                st.warning(f"{eczane_cikar.upper()} {cikis_tarihi} tarihinden itibaren çıkarılacak.")
    else:
        st.info("Değişiklik yapılmayacaksa bu alanı kapalı bırakabilirsiniz.")

    st.markdown('</div>', unsafe_allow_html=True)


# ==============================
# PLAN OLUŞTUR
# ==============================

if planla:
    if "uploaded_file" not in locals() or uploaded_file is None:
        st.error("Lütfen önce 'Veri Yükle' sekmesinden geçmiş nöbet Excel dosyasını yükleyin.")
        st.stop()

    with st.spinner("Plan oluşturuluyor..."):
        try:
            xls = read_excel_safely(uploaded_file)
            ana_sekme = find_main_sheet(xls.sheet_names)

            if ana_sekme is None:
                st.error("Ana geçmiş yük sekmesi bulunamadı.")
                st.stop()

            gecmis_yuk_df = pd.read_excel(xls, sheet_name=ana_sekme)

            gecmis_bayram_df = None
            if "GECMIS_BAYRAM" in xls.sheet_names:
                gecmis_bayram_df = pd.read_excel(xls, sheet_name="GECMIS_BAYRAM")

            file1, file2 = run_schedule(
                y=int(yil),
                m=int(ay),
                nm=int(kac_ay),
                gecmis_yuk_df=gecmis_yuk_df,
                gecmis_bayram_df=gecmis_bayram_df,
                eklenme_input=eklenme if "eklenme" in locals() else {},
                cikma_input=cikma if "cikma" in locals() else {},
            )

            with open(file1, "rb") as f:
                st.session_state.plan_data = f.read()

            with open(file2, "rb") as f:
                st.session_state.aylik_data = f.read()

            st.session_state.plan_file_name = file1
            st.session_state.aylik_file_name = file2

            st.markdown(
                '<div class="success-card">✅ Plan başarıyla oluşturuldu. Sonuçlar sekmesinden indirebilir ve analizleri inceleyebilirsiniz.</div>',
                unsafe_allow_html=True,
            )

        except Exception as e:
            st.exception(e)


# ==============================
# TAB 3 - SONUÇLAR
# ==============================

with tab_sonuc:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📥 Çıktı Dosyaları</div>', unsafe_allow_html=True)

    if "plan_data" not in st.session_state:
        st.info("Henüz plan oluşturulmadı. Sol menüden parametreleri seçip 'Plan Oluştur' butonuna basın.")
    else:
        st.markdown(
            '<div class="section-text">Oluşturulan nöbet planı ve aylık istatistik dosyalarını indirebilirsiniz.</div>',
            unsafe_allow_html=True,
        )

        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                "📄 Nöbet Planını İndir",
                st.session_state.plan_data,
                "nobet_plani.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with col2:
            st.download_button(
                "📊 Aylık İstatistik İndir",
                st.session_state.aylik_data,
                "aylik_detay.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.markdown('</div>', unsafe_allow_html=True)


# ==============================
# TAB 4 - ANALİZ
# ==============================

with tab_analiz:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📊 Plan Analizi</div>', unsafe_allow_html=True)

    if "plan_data" not in st.session_state:
        st.info("Analiz için önce plan oluşturun.")
    else:
        try:
            from io import BytesIO

            plan_xls = pd.ExcelFile(BytesIO(st.session_state.plan_data))

            if "GENEL OZET" in plan_xls.sheet_names:
                summary_df = pd.read_excel(plan_xls, sheet_name="GENEL OZET")

                c1, c2, c3, c4 = st.columns(4)
                with c1:
                    st.metric("Toplam Eczane", summary_df["Eczane"].nunique() if "Eczane" in summary_df.columns else 0)
                with c2:
                    st.metric("Toplam Nöbet", int(summary_df["Toplam Nöbet"].sum()) if "Toplam Nöbet" in summary_df.columns else 0)
                with c3:
                    st.metric("Toplam Bayram", int(summary_df["Bayram"].sum()) if "Bayram" in summary_df.columns else 0)
                with c4:
                    st.metric("Ort. Katsayı", round(summary_df["Toplam Katsayı"].mean(), 2) if "Toplam Katsayı" in summary_df.columns else 0)

                st.markdown("### Genel Özet Tablosu")
                st.dataframe(summary_df, use_container_width=True, height=420)

                if px is not None and "Toplam Katsayı" in summary_df.columns:
                    st.markdown("### Toplam Katsayı Dağılımı")
                    fig = px.bar(
                        summary_df.sort_values("Toplam Katsayı", ascending=False),
                        x="Eczane",
                        y="Toplam Katsayı",
                        color="Grup" if "Grup" in summary_df.columns else None,
                        title="Eczane Bazlı Toplam Katsayı",
                    )
                    fig.update_layout(xaxis_tickangle=-60, height=520)
                    st.plotly_chart(fig, use_container_width=True)

                if px is not None and "Bayram" in summary_df.columns:
                    st.markdown("### Bayram Dağılımı")
                    fig2 = px.bar(
                        summary_df.sort_values("Bayram", ascending=False),
                        x="Eczane",
                        y="Bayram",
                        color="Grup" if "Grup" in summary_df.columns else None,
                        title="Eczane Bazlı Bayram Nöbeti",
                    )
                    fig2.update_layout(xaxis_tickangle=-60, height=480)
                    st.plotly_chart(fig2, use_container_width=True)

                month_sheets = [s for s in plan_xls.sheet_names if s != "GENEL OZET"]
                if month_sheets:
                    st.markdown("### Aylık Takvim Ön İzleme")
                    selected_sheet = st.selectbox("Ay seç", month_sheets)
                    month_df = pd.read_excel(plan_xls, sheet_name=selected_sheet)
                    st.dataframe(month_df, use_container_width=True, height=420)
            else:
                st.warning("GENEL OZET sayfası bulunamadı.")

        except Exception as e:
            st.exception(e)

    st.markdown('</div>', unsafe_allow_html=True)
