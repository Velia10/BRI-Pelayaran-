# app.py
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

st.set_page_config(layout="wide")
st.markdown("""
# Rekap Cash In dan Cash Out BRI Pelayaran  
### PT ASDP Indonesia Ferry (Persero)
""")

uploaded_file = st.file_uploader("Silahkan upload file Excel (.xlsx) sesuai format:", type="xlsx")

if uploaded_file:
    df_raw = pd.read_excel(uploaded_file, header=None)
    df = pd.read_excel(uploaded_file, header=1)
    df.columns = df.columns.str.lower().str.strip()
    df.rename(columns={'debet': 'debit', 'ledger': 'balance'}, inplace=True)

    for col in ['credit', 'debit', 'balance']:
        df[col] = pd.to_numeric(
            df[col].astype(str)
                .str.replace(r'[^\d,.-]', '', regex=True)
                .str.replace('.', '', regex=False)
                .str.replace(',', '.', regex=False),
            errors='coerce'
        ).fillna(0)

    df['remark'] = df['remark'].astype(str).str.lower()
    df['time'] = pd.to_datetime(df['time'], format='%H:%M:%S', errors='coerce').dt.time
    df['date'] = pd.to_datetime(df['date'], format='%d/%m/%y', errors='coerce')

    def cat_cash_in(row):
        if row['credit'] == 0:
            return None
        rmk = row['remark']
        t = row['time']
        if any(x in rmk for x in ["penyeberangan", "pelabuhan", "cmspool", "kantin", "pdptn", "ticketing", "non terpadu", "non", "trpdu", "mrk", "fee asuransi", "mrk pinbuk"]):
            return "Pendapatan Cash Pooling"
        if "dari" in rmk and t == pd.to_datetime("23:59:59", format="%H:%M:%S").time():
            return "Pendapatan Cash Pooling"
        if "dari" in rmk and t == pd.to_datetime("00:00:01", format="%H:%M:%S").time():
            return "Pendapatan Bunga Deposito"
        if "interest on account" in rmk:
            return "Pendapatan Bunga Rekening"
        if "pinbuk ke" in rmk:
            return "Penerimaan Pinbuk"
        return "Pendapatan Lainnya"

    def cat_cash_out(row):
        if row['debit'] == 0:
            return None
        rmk = row['remark']
        t = row['time']
        if any(x in rmk for x in ["mandiri", "pinbuk ke mandiri", "mand"]):
            return "Pinbuk SAP"
        if "prop" in rmk or any(x in rmk for x in ["gaji", "gaji direksi", "gaji karyawan"]):
            return "Pinbuk Fidias"
        if "paypro" in rmk:
            return "Pinbuk SAP"
        if any(x in rmk for x in ["pinbuk ke", "pinbuk cicilan"]):
            return "Pinbuk Bank Lainnya"
        if "tax" in rmk:
            return "Pajak"
        if ("dari" in rmk and t == pd.to_datetime("23:59:59", format="%H:%M:%S").time()) or "fee" in rmk:
            return "Biaya Admin"
        return "Cash Out lainnya"

    df['kategori_cash_in'] = df.apply(cat_cash_in, axis=1)
    df['kategori_cash_out'] = df.apply(cat_cash_out, axis=1)

    min_date = df['date'].min().strftime('%-d %B %Y')
    max_date = df['date'].max().strftime('%-d %B %Y')
    st.markdown(f"### Periode: {min_date} - {max_date}")

    def sum_group(df, col, group):
        s = df[df[col].notna()].groupby(col)[group].sum().reset_index()
        s[group] = s[group].map(lambda x: f"{x:,.0f}".replace(",", "."))
        return s

    # Ringkasan
    open_row = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains("opening balance", case=False).any(), axis=1)]
    open_val = pd.to_numeric(open_row.iloc[0], errors='coerce') if not open_row.empty else pd.Series([])
    open_nom = open_val[open_val.notna()].iloc[-1] if not open_val.empty else 0

    close_row = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains("closing balance", case=False).any(), axis=1)]
    close_val = pd.to_numeric(close_row.iloc[0], errors='coerce') if not close_row.empty else pd.Series([])
    close_nom = close_val[close_val.notna()].iloc[-1] if not close_val.empty else 0

    cash_in = df['credit'].sum()
    cash_out = df['debit'].sum()

    summary = pd.DataFrame({
        'Kategori': ['Opening Balance', 'Total Cash In', 'Total Cash Out', 'Ending Balance'],
        'Nominal': [open_nom, cash_in, cash_out, close_nom]
    })
    summary['Nominal'] = summary['Nominal'].map(lambda x: f"{x:,.0f}".replace(",", "."))

    st.subheader("Ringkasan")
    st.dataframe(summary)

    st.subheader("Rekap Cash In")
    st.dataframe(sum_group(df, 'kategori_cash_in', 'credit'))

    st.subheader("Rekap Cash Out")
    st.dataframe(sum_group(df, 'kategori_cash_out', 'debit'))

    # Grafik
    df_chart = df[df['credit'] > 0].copy()
    df_chart = df_chart[df_chart['date'].notna()]
    df_chart['start'] = df_chart['date'] - pd.to_timedelta(df_chart['date'].dt.weekday, unit='D')
    df_chart['end'] = df_chart['start'] + pd.Timedelta(days=6)
    df_chart['label'] = df_chart.groupby('start').ngroup() + 1
    df_chart['label'] = df_chart.apply(lambda row: f"Minggu {row['label']}\n{row['start'].strftime('%-d')}–{row['end'].strftime('%-d %B %Y')}", axis=1)
    chart_data = df_chart.groupby('label')['credit'].sum() / 100000

    st.subheader("Grafik Cash In per Minggu")
    fig, ax = plt.subplots(figsize=(10, 4))
    chart_data.plot(kind='bar', ax=ax)
    for i, val in enumerate(chart_data):
        ax.text(i, val, f"Rp {int(val*100000):,}".replace(",", "."), ha='center', va='bottom')
    ax.set_ylabel("Nominal (x100,000)")
    ax.set_xlabel("Minggu")
    ax.set_title("Cash In per Minggu (x100rb)")
    plt.tight_layout()
    st.pyplot(fig)

    # Excel
    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Transaksi Detail")
        sum_group(df, 'kategori_cash_in', 'credit').to_excel(writer, index=False, sheet_name="Rekap Cash In")
        sum_group(df, 'kategori_cash_out', 'debit').to_excel(writer, index=False, sheet_name="Rekap Cash Out")
        summary.to_excel(writer, index=False, sheet_name="Ringkasan Akhir")
    towrite.seek(0)
    st.download_button("📥 Download Rekap Excel", towrite, file_name="rekap_cash_flow.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
