import re, os

def convert_mssql_to_postgresql(input_file, output_file):
    input_file = os.path.join('source', input_file)
    if not os.path.exists(input_file):
        print(f"File {input_file} tidak ditemukan. Lewati...")
        return

    with open(input_file, 'r', encoding='utf-8') as f:
        content = f.read()

    output_file = os.path.join('source', output_file)
    content = content.replace('[dbo].', 'public.')
    content = content.replace('[', '').replace(']', '')

    content = re.sub(r"N'([^']*)'", r"'\1'", content)

    content = re.sub(r"COLLATE\s+[a-zA-Z0-9_]+\s*", "", content, flags=re.IGNORECASE)

    content = re.sub(r"int\s+IDENTITY\(\d+,\d+\)", "SERIAL", content, flags=re.IGNORECASE)
    content = re.sub(r"\bdatetime\b", "timestamp", content, flags=re.IGNORECASE)
    content = re.sub(r"varchar\(max\)", "text", content, flags=re.IGNORECASE)
    content = re.sub(r"\btinyint\b", "smallint", content, flags=re.IGNORECASE)
    content = re.sub(r"\bnumeric\b", "decimal", content, flags=re.IGNORECASE)

    def to_lowercase_identifiers(match):
        full_statement = match.group(0)
        parts = re.split(r"('[^']*')", full_statement)
        for i in range(len(parts)):
            if not parts[i].startswith("'"):
                parts[i] = parts[i].lower()
        return "".join(parts)

    content = to_lowercase_identifiers(re.search(r".*", content, re.DOTALL))

    if not content.strip().endswith(';'):
        content = content.strip() + ';'

    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(content)

    print(f"Berhasil: {input_file} -> {output_file}")

files_to_process = {
    'ms_karyawan.sql': 'pg_ms_karyawan.sql',
    'detail_karyawan.sql': 'pg_detail_karyawan.sql'
}

for inp, outp in files_to_process.items():
    convert_mssql_to_postgresql(inp, outp)

def toUcwords():
    arr = [
        'PENGARUH KEPERCAYAAN DAN KEPUASAN KONSUMEN TERHADAP SISTEM TRANSAKSI ONLINE SHOP SHOPEE',
        'EFEKTIVITAS PEMANFAATAN INFORMASI GOOGLE ASSISTANT PADA SMARTPHONE ANDROID TERHADAP PEMENUHAN KEBUTUHAN INFORMASI BAGI MAHASISWA UNIVERSITAS BHAYANGKARA JAKARTA RAYA',
        'PENGARUH LINGKUNGAN KERJA DAN BEBAN KERJA TERHADAP PRODUKTIVITAS KERJA KARYAWAN',
        'PENGARUH KEPERCAYAAN DAN KEPUASAN KONSUMEN TERHADAP SISTEM TRANSAKSI ONLINE SHOP SHOPEE',
        'PENGARUH PERSEPSI KEGUNAAN, PERSEPSI KEMUDAHAN PENGGUNAAN DAN KEPERCAYAAN TERHADAP MINAT MENGGUNAKAN ARTIFICIAL INTELLIGENCE LILY BANK J TRUST',
        'PENGARUH PROMOSI, KUALITAS PELAYANAN DAN KEPUTUSAN PEMBELIAN TERHADAP PEMBELIAN ULANG (LITERATURE REVIEW MANAJEMEN PEMASARAN)',
        'PENGARUH WORK-LIFE BALANCE DAN PENGEMBANGAN KARIR TERHADAP PRODUKTIVITAS KERJA KARYAWAN PADA PT. BANK MANDIRI (PERSERO), Tbk. KCP GOWA SKRIPS Itle'
    ]

    for a in arr:
        print(a.title())

toUcwords()