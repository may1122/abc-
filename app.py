import streamlit as st
from nobet_engine import run_schedule
import datetime
import pandas as pd
import base64
from pathlib import Path


# ==============================
# SAYFA AYARI
# ==============================

st.set_page_config(
    page_title="AYÇA | Eczane Nöbet Planlayıcı",
    page_icon="💊",
    layout="wide"
)


# ==============================
# YARDIMCI
# ==============================

def get_base64_image(image_path: str):
    path = Path(image_path)
    if not path.exists():
        return None
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()


logo_base64 = get_base64_image("logo.png")


# ==============================
# CSS / TASARIM
# ==============================

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

    html, body, [class*="css"] {{
        font-family: 'Inter', sans-serif;
    }}

    .stApp {{
        background: linear-gradient(180deg, #f6f8fb 0%, #eef3f9 100%);
    }}

    .block-container {{
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 1200px;
    }}

    .hero-card {{
        background: rgba(255,255,255,0.88);
        border: 1px solid rgba(15, 23, 42, 0.06);
        border-radius: 24px;
        padding: 28px 32px;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.08);
        margin-bottom: 24px;
    }}

    .hero-grid {{
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 24px;
        flex-wrap: wrap;
    }}

    .hero-title {{
        font-size: 2.1rem;
        font-weight: 800;
        color: #0f172a;
        margin: 0 0 8px 0;
        letter-spacing: -0.02em;
    }}

    .hero-sub {{
        font-size: 1rem;
        color: #475569;
        margin: 0;
        line-height: 1.6;
    }}

    .hero-badge {{
        display: inline-block;
        background: linear-gradient(90deg, #0ea5e9 0%, #10b981 100%);
        color: white;
        font-size: 0.85rem;
        font-weight: 700;
        padding: 8px 14px;
        border-radius: 999px;
        margin-bottom: 14px;
    }}

    .logo-wrap {{
        display: flex;
        justify-content: center;
        align-items: center;
        min-width: 260px;
    }}

    .logo-wrap img {{
        max-width: 320px;
        width: 100%;
        height: auto;
    }}

    .section-card {{
        background: rgba(255,255,255,0.92);
        border: 1px solid rgba(15, 23, 42, 0.06);
        border-radius: 22px;
        padding: 22px;
        box-shadow: 0 10px 25px rgba(15, 23, 42, 0.06);
        margin-bottom: 20px;
    }}

    .section-title {{
        font-size: 1.2rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 10px;
    }}

    .section-text {{
        color: #475569;
        font-size: 0.96rem;
        line-height: 1.6;
    }}

    div[data-testid="stFileUploader"] {{
        background: white;
        border: 1px dashed #94a3b8;
        border-radius: 18px;
        padding: 12px;
    }}

    div[data-testid="stNumberInput"] > div,
    div[data-testid="stDateInput"] > div,
    div[data-testid="stTextInput"] > div,
    div[data-testid="stSelectbox"] > div {{
        border-radius: 14px;
    }}

    div.stButton > button {{
        border-radius: 14px;
        font-weight: 700;
        border: none;
        background: linear-gradient(90deg, #0f172a 0%, #0ea5e9 100%);
        color: white;
        padding: 0.7rem 1rem;
        box-shadow: 0 8px 20px rgba(14, 165, 233, 0.25);
    }}

    div.stDownloadButton > button {{
        border-radius: 14px;
        font-weight: 700;
        border: none;
        background: linear-gradient(90deg, #0f172a 0%, #10b981 100%);
        color: white;
        padding: 0.7rem 1rem;
        box-shadow: 0 8px 20px rgba(16, 185, 129, 0.22);
        width: 100%;
    }}

    section[data-testid="stSidebar"] {{
        background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
        border-right: 1px solid rgba(15,23,42,0.06);
    }}

    .mini-info {{
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 14px 16px;
        color: #334155;
        font-size: 0.93rem;
    }}
    </style>
    """,
    unsafe_allow_html=True
)


# ==============================
# HERO
# ==============================

logo_html = ""
if logo_base64:
    logo_html = f"""
    <div class="logo-wrap">
        <img src="data:image/png;base64,{logo_base64}" alt="AYÇA Logo">
    </div>
    """

st.markdown(
    f"""
    <div class="hero-card">
        <div class="hero-grid">
            <div style="flex:1; min-width:320px;">
                <div class="hero-badge">AYÇA • Akıllı Yazılım Çözüm Asistanı</div>
                <h1 class="hero-title">Eczane Nöbet Planlayıcı</h1>
                <p class="hero-sub">
                    Geçmiş yük verileri, bayram geçmişi ve dönemsel denge mantığı ile
                    daha adil, daha kontrollü ve daha okunabilir nöbet planları oluşturun.
                </p>
            </div>
            {logo_html}
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


# ==============================
# GEÇMİŞ YÜK DOSYASI
# ==============================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">📂 Geçmiş Nöbet Dosyası</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-text">Ana geçmiş yük sekmesini ve varsa <b>GECMIS_BAYRAM</b> sekmesini içeren Excel dosyasını yükleyin.</div>',
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    "Geçmiş nöbet Excel dosyasını yükleyin",
    type=["xlsx"],
    label_visibility="collapsed"
)

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names

        st.success(f"Yüklenen sekmeler: {', '.join(sheet_names)}")

        ana_sekme = None
        for s in sheet_names:
            if s.strip().upper() != "GECMIS_BAYRAM":
                ana_sekme = s
                break

        if ana_sekme is None:
            st.error("Ana geçmiş yük sekmesi bulunamadı.")
            st.stop()

        st.info(f"Ana geçmiş yük sekmesi: {ana_sekme}")

    except Exception as e:
        st.error(f"Excel okunamadı: {e}")
        st.stop()

st.markdown('</div>', unsafe_allow_html=True)


# ==============================
# SIDEBAR - PLAN AYARLARI
# ==============================

with st.sidebar:
    st.markdown("## 📅 Plan Parametreleri")

    yil = st.number_input(
        "Yıl",
        value=datetime.datetime.now().year
    )

    ay = st.number_input(
        "Başlangıç Ayı",
        min_value=1,
        max_value=12,
        value=1
    )

    kac_ay = st.number_input(
        "Kaç Ay Planlansın",
        min_value=1,
        max_value=12,
        value=3
    )

    st.markdown(
        """
        <div class="mini-info">
            Planlama, yüklenen geçmiş veri ve bayram geçmişi dikkate alınarak oluşturulur.
        </div>
        """,
        unsafe_allow_html=True
    )

    st.divider()
    planla = st.button("🚀 Plan Oluştur", use_container_width=True)


# ==============================
# ECZANE DEĞİŞİKLİK PANELİ
# ==============================

st.markdown('<div class="section-card">', unsafe_allow_html=True)
st.markdown('<div class="section-title">🔧 Eczane Değişiklikleri</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-text">İsterseniz yeni eczane ekleyebilir veya geçici olarak çıkış tarihi tanımlayabilirsiniz.</div>',
    unsafe_allow_html=True
)

degisim = st.toggle("Eczane ekleme / çıkarma yapılacak mı?")

eklenme = {}
cikma = {}

if degisim:
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("### ➕ Eczane Ekle")

        eczane_ekle = st.text_input("Eczane İsmi")
        eczane_grup = st.selectbox(
            "Grup",
            [
                "A1", "A2", "A3",
                "B1", "B2", "B3",
                "C1", "C2", "C3",
                "D1", "D2", "D3",
                "E1", "E2", "E3",
                "F1", "F2", "F3",
                "G1", "G2", "G3"
            ]
        )
        ekleme_tarihi = st.date_input(
            "Eklenme Tarihi",
            value=datetime.date.today()
        )

        if eczane_ekle:
            eklenme[eczane_ekle.upper()] = {
                "tarih": ekleme_tarihi,
                "grup": eczane_grup
            }

    with col2:
        st.markdown("### ➖ Eczane Çıkar")

        eczane_cikar = st.text_input("Eczane İsmi", key="cikar")
        cikis_tarihi = st.date_input(
            "Çıkış Tarihi",
            value=datetime.date.today()
        )

        if eczane_cikar:
            cikma[eczane_cikar.upper()] = cikis_tarihi

st.markdown('</div>', unsafe_allow_html=True)


# ==============================
# PLAN OLUŞTUR
# ==============================

if planla:
    if uploaded_file is None:
        st.error("Lütfen önce geçmiş nöbet Excel dosyasını yükleyin.")
        st.stop()

    with st.spinner("Plan oluşturuluyor..."):
        try:
            uploaded_file.seek(0)
            xls = pd.ExcelFile(uploaded_file)

            ana_sekme = None
            for s in xls.sheet_names:
                if s.strip().upper() != "GECMIS_BAYRAM":
                    ana_sekme = s
                    break

            if ana_sekme is None:
                st.error("Ana geçmiş yük sekmesi bulunamadı.")
                st.stop()

            gecmis_yuk_df = pd.read_excel(xls, sheet_name=ana_sekme)

            gecmis_bayram_df = None
            if "GECMIS_BAYRAM" in xls.sheet_names:
                gecmis_bayram_df = pd.read_excel(xls, sheet_name="GECMIS_BAYRAM")

            file1, file2 = run_schedule(
                y=yil,
                m=ay,
                nm=kac_ay,
                gecmis_yuk_df=gecmis_yuk_df,
                gecmis_bayram_df=gecmis_bayram_df,
                eklenme_input=eklenme,
                cikma_input=cikma
            )

            with open(file1, "rb") as f:
                st.session_state.plan_data = f.read()

            with open(file2, "rb") as f:
                st.session_state.aylik_data = f.read()

            st.success("Plan başarıyla oluşturuldu!")

        except Exception as e:
            st.exception(e)


# ==============================
# DOWNLOAD PANELİ
# ==============================

if "plan_data" in st.session_state:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.markdown('<div class="section-title">📥 Çıktı Dosyaları</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="section-text">Oluşturulan nöbet planı ve aylık detay dosyalarını indirebilirsiniz.</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "📄 Nöbet Planını İndir",
            st.session_state.plan_data,
            "nobet_plani.xlsx",
            use_container_width=True
        )

    with col2:
        st.download_button(
            "📊 Aylık İstatistik İndir",
            st.session_state.aylik_data,
            "aylik_detay.xlsx",
            use_container_width=True
        )

    st.markdown('</div>', unsafe_allow_html=True)
