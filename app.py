import streamlit as st
import pandas as pd
import plotly.express as px
from nobet_engine_104_eczane import run_schedule

st.set_page_config(page_title="AYÇA", layout="wide")

# ================= HERO =================

st.markdown("""

<div style='background:linear-gradient(135deg,#1f4b99,#2f6ee5);
padding:25px;border-radius:15px;color:white'>
<h2>💊 AYÇA Nöbet Planlama</h2>
<p>Akıllı nöbet planlama sistemi</p>
</div>
""", unsafe_allow_html=True)

# ================= SIDEBAR =================

with st.sidebar:
st.header("⚙️ Plan Ayarları")

```
yil = st.number_input("Yıl", 2025, 2035, 2026)
ay = st.selectbox("Ay", list(range(1,13)))
ay_sayisi = st.number_input("Kaç Ay",1,12,1)

gecmis_file = st.file_uploader("Geçmiş Excel", type=["xlsx"])
bayram_file = st.file_uploader("Bayram Excel", type=["xlsx"])
```

# ================= TAB =================

tab1, tab2, tab3 = st.tabs(["Plan","Özet","Grafik"])

# ================= PLAN =================

with tab1:

```
if gecmis_file is None:
    st.info("Excel yükle")
else:
    df = pd.read_excel(gecmis_file)
    df_b = pd.read_excel(bayram_file) if bayram_file else None

    if st.button("Plan oluştur"):

        plan_file, detail_file = run_schedule(
            yil, ay, ay_sayisi, df, df_b
        )

        st.success("Hazır")

        with open(plan_file, "rb") as f:
            st.download_button("Plan indir", f)
```

# ================= ÖZET =================

with tab2:
try:
df = pd.read_excel("Alternatif.xlsx", sheet_name="GENEL OZET")
st.dataframe(df)
except:
st.warning("Plan yok")

# ================= GRAFİK =================

with tab3:
try:
df = pd.read_excel("Alternatif.xlsx", sheet_name="GENEL OZET")

```
    fig = px.bar(df, x="Eczane", y="Toplam Katsayı")
    st.plotly_chart(fig)

except:
    st.warning("Plan yok")
```
