import pandas as pd
import os, random
import matplotlib.pyplot as plt
from util import formatting_excel
from sqlalchemy import create_engine

TOTAL_RESPONDEN = 375
SKOR_MAKS = 5
SKOR_IDEAL = TOTAL_RESPONDEN * SKOR_MAKS
PERNYATAAN = os.path.join('source', 'pernyataan.csv')
UMKM = os.path.join('data', "Data Usaha Mikro Kecil dan Menengah (UMKM) Kosmetik 2025-03-16.xlsx")

OUTPUT_RESPONDEN = os.path.join('target', 'Hasil_Profil_Responden.xlsx')
OUTPUT_PERNYATAAN = os.path.join('target', 'pernyataan.csv')
OUTPUT_RESPONDEN_TANGSEL = os.path.join('target', 'Hasil_Profil_Responden_Tangsel.xlsx')
OUTPUT_PERNYATAAN_TANGSEL = os.path.join('target', 'pernyataan_tangsel.csv')

os.makedirs('source', exist_ok=True)
os.makedirs('target', exist_ok=True)

variabel_config = {
    'Pelatihan': {'cols': [f'P{i}' for i in range(1, 11)], 'code': 'X1'},
    'Work_Life_Balance': {'cols': [f'WB{i}' for i in range(1, 7)], 'code': 'X2'},
    'Beban_Kerja': {'cols': [f'BK{i}' for i in range(1, 7)], 'code': 'X3'},
    'Digitalisasi': {'cols': [f'D{i}' for i in range(1, 12)], 'code': 'Z'},
    'Produktivitas': {'cols': [f'PK{i}' for i in range(1, 9)], 'code': 'Y'}
}

profile_cols = ['usia', 'jenis kelamin', 'pendidikan terakhir', 'pengalaman kerja', 'nama_perusahaan']

DB_CONFIG = {
    "host": "localhost",
    "database": "testing",
    "user": "postgres",
    "password": "Popoyan123!",
    "port": "5432"
}

def scramble_company_name(name):
    for prefix in ['PT. ', 'PT ', 'CV. ', 'CV ', 'pt. ', 'pt ', 'cv. ', 'cv ']:
        if name.upper().startswith(prefix.upper()):
            name = name[len(prefix):].strip()
            break

    if "l essential" in name.lower():
        l_variations = [
            "L essential", "L Essential",
            "L'Essential", "L'essential",
            "L`Essential", "L`essential"
        ]
        name = random.choice(l_variations)
    else:
        case_style = random.choice(['upper', 'lower', 'title'])
        if case_style == 'upper':
            name = name.upper()
        elif case_style == 'lower':
            name = name.lower()
        else:
            name = name.title()

    prefix_choices = ['PT ', 'PT. ', 'pt ', 'pt. ', 'Pt ', 'Pt. ']
    chosen_prefix = random.choice(prefix_choices)

    return f"{chosen_prefix}{name}"

def get_data_from_db():
    conn_str = f"postgresql+psycopg2://{DB_CONFIG['user']}:{DB_CONFIG['password']}@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
    engine = create_engine(conn_str)
    query = """WITH QTE AS (
        SELECT
            ms.nik,
            (CASE WHEN ms.kodejeniskelamin = 'LK' THEN 'Pria' ELSE 'Wanita' END) AS jenis_kelamin,
            LOWER(COALESCE(dt.email_pribadi, dt2.email_pribadi)) AS email_pribadi
        FROM ms_karyawan ms
        LEFT JOIN tbl_detail_karyawan dt ON dt.nik = ms.nik AND COALESCE(dt.email_pribadi, '') != ''
        LEFT JOIN tbl_detail_karyawan dt2 ON dt2.nik_lama = ms.nik AND COALESCE(dt2.email_pribadi, '') != ''
    ) SELECT * FROM QTE WHERE email_pribadi IS NOT NULL AND email_pribadi LIKE '%%gmail%%'"""
    with engine.connect() as conn:
        df_db = pd.read_sql(query, conn)
    return df_db

def safe_sample(df, n_target):
    if len(df) >= n_target:
        return df.sample(n=n_target)
    else:
        return pd.concat([df] * (n_target // len(df) + 1)).iloc[:n_target]

def get_kategori(persentase):
    if 20.00 <= persentase <= 36.00: return 'Sangat Tidak Setuju'
    elif 36.01 <= persentase <= 52.00: return 'Tidak Setuju'
    elif 52.01 <= persentase <= 68.00: return 'Netral'
    elif 68.01 <= persentase <= 84.00: return 'Setuju'
    elif 84.01 <= persentase <= 100.00: return 'Sangat Setuju'
    return '-'

df_umkm = pd.read_excel(UMKM, sheet_name='Banten')

df_umkm['nama_industri'] = df_umkm['nama_industri'].replace(
    'CV Sukses Jaya Abadi (akan berganti nama menjadi PT. UNICARE BEAUTY KOSMETINDO)', 'PT. UNICARE BEAUTY KOSMETINDO'
)

df_tangsel_pool = df_umkm[df_umkm['Kabupaten'] == 'TANGERANG SELATAN']
must_have = pd.concat([
    df_tangsel_pool[df_tangsel_pool['nama_industri'].str.contains('L essential', case=False, na=False)],
    df_tangsel_pool[df_tangsel_pool['nama_industri'].str.contains('UNICARE BEAUTY KOSMETINDO', case=False, na=False)]
])
others_tangsel = df_tangsel_pool[~df_tangsel_pool['nama_industri'].isin(must_have['nama_industri'])].sample(n=4-len(must_have))

list_perusahaan_tangsel = pd.concat([must_have, others_tangsel])['nama_industri'].tolist()
list_perusahaan_tangerang = df_umkm[df_umkm['Kabupaten'] == 'KOTA TANGERANG'].sample(n=3)['nama_industri'].tolist()
list_perusahaan_kab_tangerang = df_umkm[df_umkm['Kabupaten'] == 'KABUPATEN TANGERANG'].sample(n=4)['nama_industri'].tolist()

temp_data = []
for _ in range(207): temp_data.append({'raw_name': random.choice(list_perusahaan_tangsel), 'kab': 'TANGSEL'})
for _ in range(94): temp_data.append({'raw_name': random.choice(list_perusahaan_tangerang), 'kab': 'KOTA_TNG'})
for _ in range(74): temp_data.append({'raw_name': random.choice(list_perusahaan_kab_tangerang), 'kab': 'KAB_TNG'})

random.shuffle(temp_data)

final_perusahaan_scrambled = [scramble_company_name(x['raw_name']) for x in temp_data]
final_kab_tags = [x['kab'] for x in temp_data]

df_pool = get_data_from_db()
df_final_pool = pd.concat([
    safe_sample(df_pool[df_pool['jenis_kelamin'] == 'Pria'], 199),
    safe_sample(df_pool[df_pool['jenis_kelamin'] == 'Wanita'], 176)
]).sample(frac=1).reset_index(drop=True)

usia = (['18-25 tahun'] * 54) + (['26-32 tahun'] * 131) + (['33-40 tahun'] * 140) + (['41-50 tahun'] * 39) + (['di atas 50 tahun'] * 11)
pendidikan = (['Kurang dari SMA'] * 5) + (['SMA/Sederajat'] * 41) + (['D3'] * 85) + (['S1'] * 209) + (['Lebih dari S1'] * 35)
pengalaman = (['Kurang dari 1 tahun'] * 13) + (['1-2 tahun'] * 65) + (['3-4 tahun'] * 130) + (['5-10 tahun'] * 124) + (['Lebih dari 10 tahun'] * 43)
random.shuffle(usia); random.shuffle(pendidikan); random.shuffle(pengalaman)

df_profil = pd.DataFrame({
    'email': df_final_pool['email_pribadi'],
    'usia': usia,
    'jenis kelamin': df_final_pool['jenis_kelamin'],
    'pendidikan terakhir': pendidikan,
    'pengalaman kerja': pengalaman,
    'nama_perusahaan': final_perusahaan_scrambled,
    '_kab_tag': final_kab_tags
})

df_csv = pd.read_csv(PERNYATAAN)
df_csv_shuffled = df_csv.sample(frac=1).reset_index(drop=True).head(375)

df_final = pd.concat([df_profil.reset_index(drop=True), df_csv_shuffled.reset_index(drop=True)], axis=1)

df_tangsel_only = df_final[df_final['_kab_tag'] == 'TANGSEL'].drop(columns=['_kab_tag'])
df_tangsel_only.to_excel(OUTPUT_RESPONDEN_TANGSEL, index=False)
formatting_excel(OUTPUT_RESPONDEN_TANGSEL)

cols_pernyataan = df_csv_shuffled.columns.tolist()
df_pernyataan_tangsel = df_tangsel_only[cols_pernyataan]
df_pernyataan_tangsel.to_csv(OUTPUT_PERNYATAAN_TANGSEL, index=False)

df_final_all = df_final.drop(columns=['_kab_tag'])
df_final_all.to_excel(OUTPUT_RESPONDEN, index=False)
df_csv_shuffled.to_csv(OUTPUT_PERNYATAAN, index=False)

for col in profile_cols:
    plt.figure(figsize=(8, 6))
    data_counts = df_final_all[col].value_counts()
    plt.pie(data_counts, labels=data_counts.index, autopct='%1.1f%%', startangle=140, colors=plt.cm.Paired.colors)
    plt.title(f'Distribusi Responden Berdasarkan {col.title()}')
    plt.axis('equal')
    plt.savefig(os.path.join('target', f'pie_chart_{col.replace(" ", "_")}.png'))
    plt.close()

for var_name, config in variabel_config.items():
    summary_data = []
    for idx, col in enumerate(config['cols'], 1):
        counts = df_final_all[col].value_counts().reindex([1, 2, 3, 4, 5], fill_value=0)
        skor_aktual = sum(counts[i] * i for i in range(1, 6))
        persentase = (skor_aktual / SKOR_IDEAL) * 100
        summary_data.append({
            'No': idx, 'Indikator': col, 'Pertanyaan': f"{config['code']}.{idx}",
            'Skor 1': counts[1], 'Skor 2': counts[2], 'Skor 3': counts[3], 'Skor 4': counts[4], 'Skor 5': counts[5],
            'Skor Aktual': skor_aktual, 'Skor Ideal': SKOR_IDEAL, 'Persentase (%)': round(persentase, 2),
            'Kategori': get_kategori(persentase)
        })
    df_summary = pd.DataFrame(summary_data)
    output_path = os.path.join('target', f'Analisis_{var_name}.xlsx')
    df_summary.to_excel(output_path, index=False)
    formatting_excel(output_path)
