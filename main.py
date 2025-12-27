import pandas as pd
import os, re, random

def extract_emails_from_sql(sql_file_path):
    emails = []
    pattern = re.compile(r"VALUES\s*\((.*?)\);?", re.IGNORECASE)

    with open(sql_file_path, 'r', encoding='utf-8') as f:
        for line in f:
            match = pattern.search(line)
            if match:
                values = re.findall(r"N?'(.*?)'|NULL|(\d+)", match.group(1))
                cleaned_values = [v[0] if v[0] else v[1] for v in values]

                if len(cleaned_values) >= 10:
                    email = cleaned_values[9].strip()
                    if email and "@" in email:
                        emails.append(email)

    return emails

all_emails = extract_emails_from_sql(os.path.join('source', 'detail_karyawan.sql'))
if len(all_emails) < 375:
    print(f"Peringatan: Email ditemukan hanya {len(all_emails)}, target 375.")
    sample_emails = (all_emails * (375 // len(all_emails) + 1))[:375]
else:
    sample_emails = random.sample(all_emails, 375)

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
if len(df_csv) > 375:
    df_csv = df_csv.sample(n=375).reset_index(drop=True)

df_profil = pd.DataFrame({
    'email': sample_emails,
    'usia': usia,
    'jenis kelamin': gender,
    'pendidikan terakhir': pendidikan,
    'pengalaman kerja': pengalaman
})

df_csv_shuffled = df_csv.sample(frac=1, random_state=42).reset_index(drop=True)
df_csv_shuffled.to_csv(os.path.join('source', 'pernyataan_random.csv'), index=False)

df_csv_final_part = df_csv_shuffled.head(375).reset_index(drop=True)

df_final = pd.concat([df_profil, df_csv_final_part], axis=1)

output_excel = 'Hasil_Profil_Responden.xlsx'
df_final.to_excel(output_excel, index=False)