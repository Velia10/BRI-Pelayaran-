import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import timedelta
from io import BytesIO

st.set_page_config(layout="wide")

# --- UPLOAD FILE ---
uploaded_file = st.file_uploader("Upload file main_data.xlsx", type="xlsx")

if uploaded_file is not None:
    file_bytes = uploaded_file.read()
    xls = BytesIO(file_bytes)

    df = pd.read_excel(xls, sheet_name='Transaksi Detail')
    summary = pd.read_excel(xls, sheet_name='Ringkasan Akhir')
else:
    st.warning("Silakan upload file Excel.")
    st.stop()

# --- DISPLAY SUMMARY ---
st.header("Ringkasan Cash Flow")
st.dataframe(summary, use_container_width=True)

# --- GRAFIK CASH IN MINGGUAN ---
df['date'] = pd.to_datetime(df['date'])
df_ci = df[df['credit'] > 0].copy()
df_ci['start_of_week'] = df_ci['date'] - pd.to_timedelta(df_ci['date'].dt.weekday, unit='D')
df_ci['end_of_week'] = df_ci['start_of_week'] + timedelta(days=6)
df_ci['minggu_ke'] = df_ci.groupby('start_of_week', sort=False).ngroup() + 1
df_ci['label'] = df_ci.apply(lambda r: f"Minggu {r['minggu_ke']}\n{r['start_of_week'].day}-{r['end_of_week'].day} {r['start_of_week'].strftime('%B')} {r['start_of_week'].year}", axis=1)

chart_data = df_ci.groupby('label')['credit'].sum() / 100000

st.subheader("Grafik Cash In Mingguan (dalam ratusan ribu)")
fig, ax = plt.subplots(figsize=(10, 4))
chart_data.plot(kind='bar', ax=ax)
ax.set_ylabel("Nominal (x100.000)")
ax.set_xlabel("Minggu")
ax.set_title("Cash In per Minggu")
plt.xticks(rotation=0, ha='center')
for i, val in enumerate(chart_data):
    ax.text(i, val + 1, f"Rp {int(val * 100000):,}".replace(",", "."), ha='center', fontsize=8)
st.pyplot(fig)

# --- DOWNLOAD DATA ---
st.download_button("Download Excel Rekap", data=file_bytes, file_name="rekap_cash_flow.xlsx")
