import pandas as pd
import os, random
import matplotlib.pyplot as plt
import numpy as np
from sqlalchemy import create_engine

total_responden = 375

skor_maksimal_per_orang = 5
skor_ideal = total_responden * skor_maksimal_per_orang # 1875

variabel_config = {
    'Beban_Kerja': {'cols': [f'BK{i}' for i in range(1, 7)], 'code': 'X1'},
    'Work_Life_Balance': {'cols': [f'WB{i}' for i in range(1, 7)], 'code': 'X2'},
    'Pelatihan': {'cols': [f'P{i}' for i in range(1, 11)], 'code': 'X3'},
    'Digitalisasi': {'cols': [f'D{i}' for i in range(1, 12)], 'code': 'Z'},
    'Produktivitas': {'cols': [f'PK{i}' for i in range(1, 9)], 'code': 'Y'}
}

profile_cols = ['usia', 'jenis kelamin', 'pendidikan terakhir', 'pengalaman kerja']

DB_CONFIG = {
    "host": "localhost",
    "database": "testing",
    "user": "postgres",
    "password": "Popoyan123!",
    "port": "5432"
}

os.makedirs('source', exist_ok=True)
os.makedirs('target', exist_ok=True)

def get_data_from_db():
    conn_str = (
        f"postgresql+psycopg2://{DB_CONFIG['user']}:{DB_CONFIG['password']}"
        f"@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
    )

    engine = create_engine(conn_str)

    query = """WITH QTE AS (
        SELECT
            ms.nik,
            (CASE WHEN ms.kodejeniskelamin = 'LK' THEN 'Pria' ELSE 'Wanita' END) AS jenis_kelamin,
            LOWER(COALESCE(dt.email_pribadi, dt2.email_pribadi)) AS email_pribadi
        FROM ms_karyawan ms
        LEFT JOIN tbl_detail_karyawan dt
            ON dt.nik = ms.nik
            AND COALESCE(dt.email_pribadi, '') != ''
        LEFT JOIN tbl_detail_karyawan dt2
            ON dt2.nik_lama = ms.nik
            AND COALESCE(dt2.email_pribadi, '') != ''
    ) SELECT *
    FROM QTE
    WHERE email_pribadi IS NOT NULL
    AND email_pribadi LIKE '%%gmail%%'
    """

    with engine.connect() as conn:
        df_db = pd.read_sql(query, conn)

    return df_db

def safe_sample(df, n_target):
    if len(df) >= n_target:
        return df.sample(n=n_target)
    else:
        print(f"Peringatan: Data {df.iloc[0]['jenis_kelamin'] if not df.empty else ''} kurang! Dibutuhkan {n_target}, tersedia {len(df)}.")
        return pd.concat([df] * (n_target // len(df) + 1)).iloc[:n_target]

def get_kategori(persentase):
    if 20.00 <= persentase <= 36.00: return 'Sangat Tidak Setuju'
    elif 36.01 <= persentase <= 52.00: return 'Tidak Setuju'
    elif 52.01 <= persentase <= 68.00: return 'Netral'
    elif 68.01 <= persentase <= 84.00: return 'Setuju'
    elif 84.01 <= persentase <= 100.00: return 'Sangat Setuju'
    return '-'

df_pool = get_data_from_db()

pool_pria = df_pool[df_pool['jenis_kelamin'] == 'Pria'].copy()
pool_wanita = df_pool[df_pool['jenis_kelamin'] == 'Wanita'].copy()

sample_pria = safe_sample(pool_pria, 199)
sample_wanita = safe_sample(pool_wanita, 176)

df_final_pool = pd.concat([sample_pria, sample_wanita]).sample(frac=1).reset_index(drop=True)

gender = (['Pria'] * 199) + (['Wanita'] * 176)

usia = (['18-25 tahun'] * 54) + (['26-32 tahun'] * 131) + \
       (['33-40 tahun'] * 140) + (['41-50 tahun'] * 39) + \
       (['di atas 50 tahun'] * 11)

pendidikan = (['Kurang dari SMA'] * 5) + (['SMA/Sederajat'] * 41) + \
             (['D3'] * 85) + (['S1'] * 209) + (['Lebih dari S1'] * 35)

pengalaman = (['Kurang dari 1 tahun'] * 13) + (['1-2 tahun'] * 65) + \
             (['3-4 tahun'] * 130) + (['5-10 tahun'] * 124) + \
             (['Lebih dari 10 tahun'] * 43)

random.shuffle(gender)
random.shuffle(usia)
random.shuffle(pendidikan)
random.shuffle(pengalaman)

df_csv = pd.read_csv(os.path.join('source', 'pernyataan.csv'))
df_csv_shuffled = df_csv.sample(frac=1).reset_index(drop=True).head(375)

df_profil = pd.DataFrame({
    'email': df_final_pool['email_pribadi'],
    'usia': usia,
    'jenis kelamin': gender,
    'pendidikan terakhir': pendidikan,
    'pengalaman kerja': pengalaman
})

df_csv_final_part = df_csv_shuffled.head(375).reset_index(drop=True)

df_final = pd.concat([df_profil.reset_index(drop=True), df_csv_shuffled.reset_index(drop=True)], axis=1)

df_final.to_excel(os.path.join('target', 'Hasil_Profil_Responden.xlsx'), index=False)
df_csv_shuffled.to_csv(os.path.join('target', 'pernyataan.csv'), index=False)

for col in profile_cols:
    plt.figure(figsize=(8, 6))
    data_counts = df_final[col].value_counts()

    plt.pie(data_counts, labels=data_counts.index, autopct='%1.1f%%', startangle=140, colors=plt.cm.Paired.colors)
    plt.title(f'Distribusi Responden Berdasarkan {col.title()}')
    plt.axis('equal')

    plt.savefig(os.path.join('target', f'pie_chart_{col.replace(" ", "_")}.png'))
    plt.close()

for var_name, config in variabel_config.items():
    summary_data = []
    columns = config['cols']
    var_code = config['code']

    for idx, col in enumerate(columns, 1):
        counts = df_final[col].value_counts().reindex([1, 2, 3, 4, 5], fill_value=0)

        skor_aktual = sum(counts[i] * i for i in range(1, 6))
        persentase = (skor_aktual / skor_ideal) * 100

        row = {
            'No': idx,
            'Indikator': col,
            'Pertanyaan': f"{var_code}.{idx}",
            'Skor 1': counts[1],
            'Skor 2': counts[2],
            'Skor 3': counts[3],
            'Skor 4': counts[4],
            'Skor 5': counts[5],
            'Skor Aktual': skor_aktual,
            'Skor Ideal': skor_ideal,
            'Persentase (%)': round(persentase, 2),
            'Kategori': get_kategori(persentase)
        }
        summary_data.append(row)

    df_summary = pd.DataFrame(summary_data)
    output_path = os.path.join('target', f'Analisis_{var_name}.xlsx')
    df_summary.to_excel(output_path, index=False)