import streamlit as st
import pandas as pd
import plotly.express as px
from nobet_engine_104_eczane import run_schedule

st.set_page_config(page_title="AYÇA", layout="wide")

# ================= SIDEBAR =================

with st.sidebar:
st.header("⚙️ Plan Ayarları")

```
yil = st.number_input("Yıl", 2025, 2035, 2026)
ay = st.selectbox("Ay", list(range(1, 13)))
ay_sayisi = st.number_input("Kaç Ay", 1, 12, 1)

gecmis_file = st.file_uploader("Geçmiş Excel", type=["xlsx"])
bayram_file = st.file_uploader("Bayram Excel", type=["xlsx"])
```

# ================= ANA =================

st.title("💊 AYÇA Nöbet Planlama")

tab1, tab2 = st.tabs(["Plan", "Özet"])

# ================= PLAN =================

with tab1:

```
if gecmis_file is None:
    st.info("Excel yükleyin")
else:
    df = pd.read_excel(gecmis_file)
    df_b = pd.read_excel(bayram_file) if bayram_file else None

    if st.button("Plan oluştur"):

        plan_file, detail_file = run_schedule(
            yil, ay, ay_sayisi, df, df_b
        )

        st.success("Plan hazır")

        with open(plan_file, "rb") as f:
            st.download_button("Plan indir", f)
```

# ================= ÖZET =================

with tab2:
try:
df = pd.read_excel("Alternatif.xlsx", sheet_name="GENEL OZET")
st.dataframe(df)
except:
st.warning("Önce plan oluştur")
