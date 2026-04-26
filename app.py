import streamlit as st
import pandas as pd
import plotly.express as px

from nobet_engine_104_eczane import run_schedule

st.set_page_config(
page_title="AYÇA | Nöbet Planlama",
page_icon="💊",
layout="wide"
)

# ==============================

# CSS

# ==============================

st.markdown("""

<style>
.main {background-color: #f7f9fc;}

.hero {
    background: linear-gradient(135deg,#1f4b99,#2f6ee5);
    padding:30px;
    border-radius:20px;
    color:white;
    margin-bottom:20px;
}

.card {
    background:white;
    padding:20px;
    border-radius:16px;
    box-shadow:0 6px 18px rgba(0,0,0,0.06);
}

.stButton button {
    background:#1f4b99;
    color:white;
    border-radius:12px;
    font-weight:bold;
}
</style>

""", unsafe_allow_html=True)

# ==============================

# HERO

# ==============================

st.markdown("""

<div class="hero">
<h2>💊 AYÇA Nöbet Planlama Paneli</h2>
<p>100+ eczane için akıllı, dengeli ve adil nöbet planlama sistemi</p>
</div>
""", unsafe_allow_html=True)

# ==============================

# SIDEBAR

# ==============================

with st.sidebar:
st.header("⚙️ Plan Ayarları")

```
yil = st.number_input("Yıl", 2025, 2035, 2026)
ay = st.selectbox("Başlangıç Ayı", list(range(1, 13)))
ay_sayisi = st.number_input("Kaç Ay", 1, 12, 1)

st.divider()

gecmis_file = st.file_uploader("Geçmiş Nöbet Excel", type=["xlsx"])
bayram_file = st.file_uploader("Bayram Excel (opsiyonel)", type=["xlsx"])
```

# ==============================

# TABLAR

# ==============================

tab1, tab2, tab3 = st.tabs(["🚀 Plan Oluştur", "📊 Özet", "📈 Grafik"])

# ==============================

# PLAN TAB

# ==============================

with tab1:

```
if gecmis_file is None:
    st.info("Devam etmek için geçmiş nöbet Excel yükleyin.")
else:
    df = pd.read_excel(gecmis_file)
    df_b = pd.read_excel(bayram_file) if bayram_file else None

    st.dataframe(df, use_container_width=True)

    if st.button("Planı Oluştur", use_container_width=True):

        with st.spinner("Plan oluşturuluyor..."):

            plan_file, detail_file = run_schedule(
                y=int(yil),
                m=int(ay),
                nm=int(ay_sayisi),
                gecmis_yuk_df=df,
                gecmis_bayram_df=df_b
            )

        st.success("Plan oluşturuldu!")

        # İNDİRME
        with open(plan_file, "rb") as f:
            st.download_button(
                "📥 Plan Excel indir",
                f,
                file_name=plan_file
            )

        with open(detail_file, "rb") as f:
            st.download_button(
                "📥 Detay Excel indir",
                f,
                file_name=detail_file
            )
```

# ==============================

# ÖZET TAB

# ==============================

with tab2:
try:
df_summary = pd.read_excel("Alternatif.xlsx", sheet_name="GENEL OZET")

```
    c1, c2, c3 = st.columns(3)

    c1.metric("Toplam Eczane", df_summary["Eczane"].nunique())
    c2.metric("Toplam Nöbet", df_summary["Toplam Nöbet"].sum())
    c3.metric("Toplam Bayram", df_summary["Bayram"].sum())

    st.dataframe(df_summary, use_container_width=True)

except:
    st.warning("Özet için önce plan oluştur.")
```

# ==============================

# GRAFİK TAB

# ==============================

with tab3:
try:
df_summary = pd.read_excel("Alternatif.xlsx", sheet_name="GENEL OZET")

```
    fig = px.bar(
        df_summary,
        x="Eczane",
        y="Toplam Katsayı",
        color="Grup"
    )

    fig.update_layout(xaxis_tickangle=-60)

    st.plotly_chart(fig, use_container_width=True)

except:
    st.warning("Grafik için önce plan oluştur.")
```
