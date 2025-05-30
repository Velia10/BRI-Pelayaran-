import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from datetime import timedelta
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

    df['remark'] = df['remark'].astype(str).str.lower()
    df['time'] = pd.to_datetime(df['time'], format='%H:%M:%S', errors='coerce').dt.time
    df['date'] = pd.to_datetime(df['date'], format='%d/%m/%y', errors='coerce')

    for col in ['credit', 'debit', 'balance']:
        df[col] = (df[col].astype(str).str.replace(r'[^\d,.-]', '', regex=True)
                    .str.replace('.', '', regex=False).str.replace(',', '.', regex=False)
                    .astype(float).fillna(0))

    def classify_cash_in(row):
        if row['credit'] == 0:
            return None
        rmk = row['remark']
        t = row['time']
        if any(k in rmk for k in ["penyeberangan", "pelabuhan", "cmspool", "kantin", "pdptn", "fee asuransi", "mrk pinbuk", "ticketing", "non terpadu", "non", "trpdu", "mrk"]):
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

    df['Kategori Cash In'] = df.apply(classify_cash_in, axis=1)

    def classify_cash_out(row):
        if row['debit'] == 0:
            return None
        rmk = row['remark']
        t = row['time']
        if any(k in rmk for k in ["mandiri", "pinbuk ke mandiri", "mand"]):
            return "Pinbuk SAP"
        if "prop" in rmk or any(k in rmk for k in ["gaji karyawan", "gaji", "gaji direksi"]):
            return "Pinbuk Fidias"
        if any(k in rmk for k in ["paypro",]):
            return "Pinbuk SAP"
        if any(k in rmk for k in ["pinbuk ke", "pinbuk cicilan"]):
            return "Pinbuk Bank Lainnya"
        if "tax" in rmk:
            return "Pajak"
        if ("dari" in rmk and t == pd.to_datetime("23:59:59", format="%H:%M:%S").time()) or "fee" in rmk:
            return "Biaya Admin"
        return "Cash Out lainnya"

    df['Kategori Cash Out'] = df.apply(classify_cash_out, axis=1)

    min_date = df['date'].min().strftime('%-d %B %Y')
    max_date = df['date'].max().strftime('%-d %B %Y')
    st.markdown(f"### Periode: {min_date} - {max_date}")

    dfi = df[df['credit'] > 0].copy()
    dfi['start_of_week'] = dfi['date'] - dfi['date'].dt.weekday * pd.Timedelta(days=1)
    dfi['end_of_week'] = dfi['start_of_week'] + pd.Timedelta(days=6)
    dfi['minggu_ke'] = dfi.groupby(['start_of_week']).ngroup() + 1
    dfi['label'] = dfi.apply(lambda row: f"Minggu {row['minggu_ke']}\n{row['start_of_week'].day}–{row['end_of_week'].day} {row['start_of_week'].strftime('%B')} {row['start_of_week'].year}", axis=1)
    weekly_chart = dfi.groupby('label')['credit'].sum().sort_index() / 100000

    st.subheader("Grafik Cash In per Minggu")
    fig, ax = plt.subplots(figsize=(10, 4))
    weekly_chart.plot(kind='bar', ax=ax)
    ax.set_title('Cash In per Minggu (x100rb)')
    ax.set_ylabel('Nominal (x100,000)')
    ax.set_xlabel('Minggu')
    for idx, value in enumerate(weekly_chart):
        ax.text(idx, value + 1, f"Rp {int(value * 100000):,}".replace(",", "."), ha='center', va='bottom', fontsize=8)
    plt.xticks(rotation=0)
    st.pyplot(fig)

    opening_row = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains("opening balance", case=False, na=False)).any(axis=1)]
    opening_balance = opening_row[2].values[0] if not opening_row.empty else 0
    closing_row = df_raw[df_raw.apply(lambda row: row.astype(str).str.contains("closing balance", case=False, na=False)).any(axis=1)]
    closing_balance = closing_row[13].values[0] if not closing_row.empty else 0

    total_in = df['credit'].sum()
    total_out = df['debit'].sum()

    summary = pd.DataFrame({
        'Kategori': ['Opening Balance', 'Total Cash In', 'Total Cash Out', 'Ending Balance'],
        'Nominal': [opening_balance, total_in, total_out, closing_balance]
    })

    st.subheader("Ringkasan")
    st.dataframe(summary)

    st.subheader("Rekap Cash In")
    st.dataframe(df.groupby('Kategori Cash In')['credit'].sum().reset_index())

    st.subheader("Rekap Cash Out")
    st.dataframe(df.groupby('Kategori Cash Out')['debit'].sum().reset_index())

    towrite = BytesIO()
    with pd.ExcelWriter(towrite, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name="Transaksi Detail")
        df.groupby('Kategori Cash In')['credit'].sum().reset_index().to_excel(writer, index=False, sheet_name="Rekap Cash In")
        df.groupby('Kategori Cash Out')['debit'].sum().reset_index().to_excel(writer, index=False, sheet_name="Rekap Cash Out")
        summary.to_excel(writer, index=False, sheet_name="Ringkasan Akhir")
    towrite.seek(0)
    st.download_button("📥 Download Rekap Excel", towrite, file_name="rekap_cash_flow.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
