# cash_flow_main_data.py

import pandas as pd
import matplotlib.pyplot as plt
from datetime import timedelta

# Load main data Excel
MAIN_DATA_PATH = "rekap_cash_flow_summary_final_with_year_labels.xlsx"

def load_data(filepath=MAIN_DATA_PATH):
    df = pd.read_excel(filepath, sheet_name="Transaksi Detail")
    df['date'] = pd.to_datetime(df['date'])
    df['remark'] = df['remark'].astype(str).str.lower()
    return df

def get_cash_in_summary(df):
    return df[df['credit'] > 0].groupby('Kategori Cash In Pelayaran', as_index=False)['credit'].sum()

def get_cash_out_summary(df):
    return df[df['debit'] > 0].groupby('Kategori Cash Out Pelayaran', as_index=False)['debit'].sum()

def get_weekly_cash_in_chart(df):
    df_ci = df[df['credit'] > 0].copy()
    df_ci['start_of_week'] = df_ci['date'] - pd.to_timedelta(df_ci['date'].dt.weekday, unit='d')
    df_ci['end_of_week'] = df_ci['start_of_week'] + timedelta(days=6)
    df_ci['minggu_ke'] = df_ci.groupby('start_of_week').ngroup() + 1
    df_ci['label'] = df_ci.apply(lambda row: f"Minggu {row['minggu_ke']}\n{row['start_of_week'].day}-{row['end_of_week'].day} {row['start_of_week'].strftime('%B')} {row['start_of_week'].year}", axis=1)

    chart_data = df_ci.groupby('label')['credit'].sum() / 100000
    fig, ax = plt.subplots(figsize=(12, 4))
    chart_data.plot(kind='bar', ax=ax)
    ax.set_title('Cash In per Minggu (x100,000)')
    ax.set_ylabel('Nominal')
    ax.set_xlabel('Minggu')
    for idx, value in enumerate(chart_data):
        ax.text(idx, value + 1, f"Rp {int(value * 100000):,}".replace(",", "."), ha='center', fontsize=8)
    plt.tight_layout()
    plt.show()

if __name__ == '__main__':
    df_main = load_data()
    print("Summary Cash In:")
    print(get_cash_in_summary(df_main))
    print("\nSummary Cash Out:")
    print(get_cash_out_summary(df_main))
    get_weekly_cash_in_chart(df_main)
