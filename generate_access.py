import os
import random
import string
import pandas as pd

# --- KONFIGURASI PATH ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PATH_MASTER = os.path.join(BASE_DIR, "1_Data_Master")
OUTPUT_EXCEL = os.path.join(BASE_DIR, "Daftar_Akses_Portal.xlsx")

def generate_random_pin(length=6):
    """Membuat PIN acak alfanumerik (huruf kapital dan angka)"""
    chars = string.ascii_uppercase + string.digits
    return ''.join(random.choice(chars) for _ in range(length))

def clean_nrp(val):
    """Membersihkan format angka NRP agar tidak ada titik nol (.0)"""
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s

def build_access_list():
    print("=" * 60)
    print(" MENYUSUN DATA MASTER AKSES PORTAL KARYAWAN ")
    print("=" * 60)
    
    # Load data akses lama supaya PIN karyawan yang sudah terdaftar tidak berubah
    existing_pins = {}
    if os.path.exists(OUTPUT_EXCEL):
        try:
            df_old = pd.read_excel(OUTPUT_EXCEL, dtype=str)
            for _, row in df_old.iterrows():
                nrp_key = clean_nrp(row.get('NRP', ''))
                pin_val = row.get('PIN Akses', '')
                if nrp_key and pin_val:
                    existing_pins[nrp_key] = pin_val
            print(f"📂 Berhasil memuat data PIN akses lama ({len(existing_pins)} data terbaca).")
        except Exception as e:
            print(f"⚠️ Catatan: Gagal membaca file lama ({e}), membuat baru...")

    all_data = []
    
    sources = [
        {"file": "Data_NonStaff.xlsx", "company": "NON STAFF"},
        {"file": "Data Absensi_BMP.xlsx", "company": "BMP"},
        {"file": "Data Absensi_CMU.xlsx", "company": "CMU"}
    ]
    
    for src in sources:
        file_path = os.path.join(PATH_MASTER, src["file"])
        if not os.path.exists(file_path):
            print(f"❌ File tidak ditemukan: {src['file']}")
            continue
            
        print(f"📖 Membaca {src['file']}...")
        try:
            # Baca file excel tanpa baris header otomatis, ambil berdasarkan posisi kolom (B=Kolom indeks 1, C=Kolom indeks 2)
            df = pd.read_excel(file_path, header=None, dtype=str)
            
            # Mulai dari baris ke-1 (baris ke-0 adalah judul/header tabel)
            for index, row in df.iterrows():
                if index == 0:
                    continue
                
                nrp_raw = row[1] if len(row) > 1 else None  # Kolom B (NRP)
                name_raw = row[2] if len(row) > 2 else None  # Kolom C (Nama)
                
                nrp = clean_nrp(nrp_raw)
                name = str(name_raw).strip() if pd.notna(name_raw) else ""
                
                # Validasi baris kosong atau tidak valid
                if not nrp or not name or name.lower() == 'nan' or nrp.lower() == 'nan':
                    continue
                    
                # Gunakan PIN lama jika sudah ada, atau buat baru jika karyawan baru
                if nrp in existing_pins:
                    pin = existing_pins[nrp]
                else:
                    pin = generate_random_pin(6)
                    existing_pins[nrp] = pin
                    
                all_data.append({
                    "Perusahaan": src["company"],
                    "NRP": nrp,
                    "Nama": name,
                    "PIN Akses": pin
                })
                
        except Exception as e:
            print(f"❌ Gagal memproses {src['file']}: {e}")
            
    if not all_data:
        print("❌ Tidak ada data yang berhasil dikumpulkan!")
        return
        
    # Simpan ke file Excel terpusat
    df_result = pd.DataFrame(all_data)
    # Hapus duplikat berdasarkan NRP jika ada
    df_result = df_result.drop_duplicates(subset=['NRP']).sort_values(by=["Perusahaan", "Nama"]).reset_index(drop=True)
    
    df_result.to_excel(OUTPUT_EXCEL, index=False)
    print("\n" + "=" * 60)
    print(f"✅ SUKSES! File master akses diperbarui: 'Daftar_Akses_Portal.xlsx'")
    print(f"📊 Total Karyawan terdaftar: {len(df_result)} orang")
    print("=" * 60)

if __name__ == "__main__":
    build_access_list()