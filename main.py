import pandas as pd
import os, random
import matplotlib.pyplot as plt
from util import formatting_excel
from sqlalchemy import create_engine

TOTAL_RESPONDEN = 433
SKOR_MAKS = 5
SKOR_IDEAL = TOTAL_RESPONDEN * SKOR_MAKS
PERNYATAAN_375 = os.path.join('source', 'pernyataan.csv')
PERNYATAAN_58 = os.path.join('source', 'pernyataan_58.csv')
UMKM = os.path.join('data', "Data Usaha Mikro Kecil dan Menengah (UMKM) Kosmetik 2025-03-16.xlsx")
RESPONDEN_ORI = os.path.join('target', 'Hasil_Profil_Responden_Ori.xlsx')

OUTPUT_RESPONDEN = os.path.join('target', 'Hasil_Profil_Responden.xlsx')
OUTPUT_RESPONDEN_TANGSEL = os.path.join('target', 'Hasil_Profil_Responden_Tangsel.xlsx')
OUTPUT_PERNYATAAN = os.path.join('target', 'pernyataan.csv')
OUTPUT_PERNYATAAN_TANGSEL = os.path.join('target', 'pernyataan_tangsel.csv')

os.makedirs('source', exist_ok=True)
os.makedirs('target', exist_ok=True)

profile_cols = ['usia', 'jenis kelamin', 'pendidikan terakhir', 'pengalaman kerja', 'nama_perusahaan']

variabel_config = {
    'Pelatihan': {'cols': [f'P{i}' for i in range(1, 11)], 'code': 'X1'},
    'Work_Life_Balance': {'cols': [f'WB{i}' for i in range(1, 7)], 'code': 'X2'},
    'Beban_Kerja': {'cols': [f'BK{i}' for i in range(1, 7)], 'code': 'X3'},
    'Digitalisasi': {'cols': [f'D{i}' for i in range(1, 12)], 'code': 'Z'},
    'Produktivitas': {'cols': [f'PK{i}' for i in range(1, 9)], 'code': 'Y'}
}

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
        name = random.choice(["L essential", "L Essential", "L'Essential", "L'essential", "L`Essential", "L`essential"])
    else:
        case_style = random.choice(['upper', 'lower', 'title'])
        if case_style == 'upper': name = name.upper()
        elif case_style == 'lower': name = name.lower()
        else: name = name.title()
    return f"{random.choice(['PT ', 'PT. ', 'pt ', 'pt. ', 'Pt ', 'Pt. '])}{name}"

def generate_list(items, total):
    res = []
    for item, count in items:
        res.extend([item] * count)

    while len(res) < total: res.append(items[0][0])
    random.shuffle(res)
    return res[:total]

def get_kategori(persentase):
    if 20.00 <= persentase <= 36.00: return 'Sangat Tidak Setuju'
    elif 36.01 <= persentase <= 52.00: return 'Tidak Setuju'
    elif 52.01 <= persentase <= 68.00: return 'Netral'
    elif 68.01 <= persentase <= 84.00: return 'Setuju'
    elif 84.01 <= persentase <= 100.00: return 'Sangat Setuju'
    return '-'

def create_pie_chart(df, title):
    for col in profile_cols:
        if col == 'nama_perusahaan':
            continue

        output = os.path.join('target', f'pie_chart_{col.replace(" ", "_")}_tangsel.png')
        plt.figure(figsize=(8, 6))
        data_counts = df[col].value_counts()
        plt.pie(data_counts, labels=data_counts.index, autopct='%1.1f%%', startangle=140, colors=plt.cm.Paired.colors)
        plt.title(f'Distribusi {title} Berdasarkan {col.title()}')
        plt.axis('equal')
        plt.savefig(output)
        plt.close()
        print(f"File disimpan di: {output}")

def create_analisis(df, title=''):
    for var_name, config in variabel_config.items():
        summary_data = []
        for idx, col in enumerate(config['cols'], 1):
            counts = df[col].value_counts().reindex([1, 2, 3, 4, 5], fill_value=0)
            skor_aktual = sum(counts[i] * i for i in range(1, 6))
            persentase = (skor_aktual / SKOR_IDEAL) * 100
            summary_data.append({
                'No': idx, 'Indikator': col, 'Pertanyaan': f"{config['code']}.{idx}",
                'Skor 1': counts[1], 'Skor 2': counts[2], 'Skor 3': counts[3], 'Skor 4': counts[4], 'Skor 5': counts[5],
                'Skor Aktual': skor_aktual, 'Skor Ideal': SKOR_IDEAL, 'Persentase (%)': round(persentase, 2),
                'Kategori': get_kategori(persentase)
            })
        df_summary = pd.DataFrame(summary_data)
        output_path = os.path.join('target', f'Analisis_{var_name}{f"_{title}" if title else ''}.xlsx')
        df_summary.to_excel(output_path, index=False)
        formatting_excel(output_path)

def get_data_from_db():
    df_ori = pd.read_excel(RESPONDEN_ORI)
    list_email_ori = df_ori['email'].unique().tolist()
    email_in_clause = "'" + "','".join(list_email_ori) + "'"

    conn_str = f"postgresql+psycopg2://{DB_CONFIG['user']}:{DB_CONFIG['password']}@{DB_CONFIG['host']}:{DB_CONFIG['port']}/{DB_CONFIG['database']}"
    engine = create_engine(conn_str)

    query = f"""
    WITH QTE AS (
        SELECT
            ms.nik,
            (CASE WHEN ms.kodejeniskelamin = 'LK' THEN 'Pria' ELSE 'Wanita' END) AS jenis_kelamin,
            LOWER(COALESCE(dt.email_pribadi, dt2.email_pribadi)) AS email_pribadi,
            (
                CASE
                    WHEN ms.kodependidikan IN ('1', '2') THEN 'Kurang dari SMA'
                    WHEN ms.kodependidikan IN ('3', '4') THEN 'SMA/Sederajat'
                    WHEN ms.kodependidikan IN ('5', '6', '7') THEN 'D3'
                    WHEN ms.kodependidikan = '8' THEN 'S1'
                    WHEN ms.kodependidikan IN ('9', 'A') THEN 'Lebih dari S1'
                    ELSE 'Lainnya'
                END
            ) AS pendidikan_terakhir
        FROM ms_karyawan ms
        LEFT JOIN tbl_detail_karyawan dt ON dt.nik = ms.nik AND COALESCE(dt.email_pribadi, '') != ''
        LEFT JOIN tbl_detail_karyawan dt2 ON dt2.nik_lama = ms.nik AND COALESCE(dt2.email_pribadi, '') != ''
        WHERE ms.tglpengundurandiri IS NULL
    ),
    FILTERED_QTE AS (
        SELECT DISTINCT ON (email_pribadi) * FROM QTE
        WHERE email_pribadi IS NOT NULL
        AND email_pribadi LIKE '%%gmail%%'
        AND pendidikan_terakhir != 'Lainnya'
    ),
    OLD_375 AS (
        SELECT * FROM FILTERED_QTE
        WHERE email_pribadi IN ({email_in_clause})
    ),
    TARGET_COUNT AS (
        SELECT 433 - COUNT(*) AS n_needed FROM OLD_375
    ),
    NEW_POOL AS (
        SELECT
            f.*,
            ROW_NUMBER() OVER (ORDER BY RANDOM()) as rn
        FROM FILTERED_QTE f
        WHERE email_pribadi NOT IN ({email_in_clause})
    )
    SELECT nik, jenis_kelamin, email_pribadi, pendidikan_terakhir FROM OLD_375
    UNION ALL
    SELECT nik, jenis_kelamin, email_pribadi, pendidikan_terakhir FROM NEW_POOL n
    JOIN TARGET_COUNT tc ON n.rn <= tc.n_needed;
    """
    with engine.connect() as conn:
        df_pool = pd.read_sql(query, conn)

    return df_pool.reset_index(drop=True)

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

temp_perusahaan = []
for _ in range(290): temp_perusahaan.append({'raw_name': random.choice(list_perusahaan_tangsel), 'kab': 'TANGSEL'})
for _ in range(72): temp_perusahaan.append({'raw_name': random.choice(list_perusahaan_tangerang), 'kab': 'KOTA_TNG'})
for _ in range(71): temp_perusahaan.append({'raw_name': random.choice(list_perusahaan_kab_tangerang), 'kab': 'KAB_TNG'})
random.shuffle(temp_perusahaan)

final_perusahaan_scrambled = [scramble_company_name(x['raw_name']) for x in temp_perusahaan]
final_kab_tags = [x['kab'] for x in temp_perusahaan]

usia = generate_list([('18-25 tahun', 62), ('26-32 tahun', 152), ('33-40 tahun', 161), ('41-50 tahun', 45), ('di atas 50 tahun', 13)], 433)
pendidikan = generate_list([('Kurang dari SMA', 6), ('SMA/Sederajat', 47), ('D3', 98), ('S1', 242), ('Lebih dari S1', 40)], 433)
pengalaman = generate_list([('Kurang dari 1 tahun', 15), ('1-2 tahun', 75), ('3-4 tahun', 150), ('5-10 tahun', 143), ('Lebih dari 10 tahun', 50)], 433)

df_final_pool = get_data_from_db()

df_profil = pd.DataFrame({
    'email': df_final_pool['email_pribadi'],
    'usia': usia,
    'jenis kelamin': df_final_pool['jenis_kelamin'],
    'pendidikan terakhir': pendidikan,
    'pengalaman kerja': pengalaman,
    'nama_perusahaan': final_perusahaan_scrambled,
    '_kab_tag': final_kab_tags
})

df_p1 = pd.read_csv(PERNYATAAN_375)
df_p2 = pd.read_csv(PERNYATAAN_58)
df_csv_combined = pd.concat([df_p1, df_p2]).reset_index(drop=True)

df_final = pd.concat([df_profil, df_csv_combined], axis=1)

df_tangsel_only = df_final[df_final['_kab_tag'] == 'TANGSEL'].drop(columns=['_kab_tag'])
df_tangsel_only.to_excel(OUTPUT_RESPONDEN_TANGSEL, index=False)
formatting_excel(OUTPUT_RESPONDEN_TANGSEL)

cols_p = df_csv_combined.columns.tolist()
df_final[df_final['_kab_tag'] == 'TANGSEL'][cols_p].to_csv(OUTPUT_PERNYATAAN_TANGSEL, index=False)
print(f'File disimpan di: {OUTPUT_PERNYATAAN_TANGSEL}')

df_final_all = df_final.drop(columns=['_kab_tag'])
df_final_all.to_excel(OUTPUT_RESPONDEN, index=False)
formatting_excel(OUTPUT_RESPONDEN)
df_csv_combined.to_csv(OUTPUT_PERNYATAAN, index=False)
print(f'File disimpan di: {OUTPUT_PERNYATAAN}')

create_pie_chart(df_tangsel_only, 'Responden Tangsel')
create_pie_chart(df_final_all, 'Semua Responden')
create_analisis(df_tangsel_only, 'Tangsel')
create_analisis(df_final_all)
