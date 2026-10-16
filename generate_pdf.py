import pandas as pd
import random
import string
import sqlite3
import os
import win32com.client as win32
import pythoncom

# 1. Fungsi PIN Acak
def generate_pin(length=6):
    characters = string.ascii_uppercase + string.digits
    return ''.join(random.choice(characters) for i in range(length))

# 2. Persiapan Database
conn = sqlite3.connect('database.db')
cursor = conn.cursor()
cursor.execute('''
    CREATE TABLE IF NOT EXISTS users (
        nrp TEXT PRIMARY KEY,
        nama TEXT,
        pin TEXT,
        perusahaan TEXT,
        pdf_path TEXT
    )
''')
conn.commit()

# 3. Path Folder Absolut
base_dir = os.path.dirname(os.path.abspath(__file__))
template_path = os.path.abspath(os.path.join(base_dir, '2_Template_Timesheet', 'Template_NonStaff.xlsx'))
pdf_output_dir = os.path.abspath(os.path.join(base_dir, '3_Output_PDF', 'NonStaff'))
os.makedirs(pdf_output_dir, exist_ok=True)

excel = None

try:
    print("--- Membaca Data Master NON STAFF ---")
    df_master = pd.read_excel(os.path.join(base_dir, '1_Data_Master', 'Data_NonStaff.xlsx'))
    
    # Standarisasi kolom Date di Data Master menjadi format string 'DD.MM.YYYY'
    df_master['Clean_Date'] = pd.to_datetime(df_master['Date'], errors='coerce').dt.strftime('%d.%m.%Y')

    unique_nrps = df_master['NRP'].unique()
    daftar_akses = []

    # Jalankan MS Excel Engine
    pythoncom.CoInitialize()
    excel = win32.Dispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    for nrp in unique_nrps:
        emp_data = df_master[df_master['NRP'] == nrp]
        
        nrp_str = str(nrp)
        nama = str(emp_data['Nama'].iloc[0]) if 'Nama' in emp_data.columns else "-"
        section = str(emp_data['Section'].iloc[0]) if 'Section' in emp_data.columns else "-"
        job_title = str(emp_data['Job Title'].iloc[0]) if 'Job Title' in emp_data.columns else "-"

        # Otomatis deteksi Bulan & Tahun dari data tanggal pertama karyawan
        raw_date_sample = emp_data['Date'].iloc[0]
        parsed_date = pd.to_datetime(raw_date_sample, errors='coerce')
        if pd.notnull(parsed_date):
            # Mengubah nama bulan ke dalam bahasa Indonesia (opsional, atau biarkan standar Inggris)
            bulan_dict = {
                1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April', 
                5: 'Mei', 6: 'Juni', 7: 'Juli', 8: 'Agustus', 
                9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
            }
            nama_bulan = bulan_dict.get(parsed_date.month, parsed_date.strftime('%B'))
            tahun = parsed_date.year
            periode_str = f"{nama_bulan}_{tahun}"
        else:
            periode_str = "Timesheet"

        print(f"Memproses: {nrp_str} - {nama} (Periode: {periode_str})...")

        # Buka Template Asli Langsung di Excel
        wb = excel.Workbooks.Open(template_path)
        ws = wb.Sheets(1)

        # Isi Header
        ws.Range('D5').Value = nrp_str
        ws.Range('D6').Value = nama
        ws.Range('D7').Value = job_title
        ws.Range('D8').Value = section

        # Buat dictionary pencocokan harian
        daily_dict = {}
        for index, row in emp_data.iterrows():
            tgl_Key = str(row['Clean_Date']).strip()
            
            in_val = row['Actual Start Time']
            out_val = row['Actual End Time']
            
            in_str = str(in_val).split(' ')[-1] if pd.notnull(in_val) else ""
            if len(in_str) > 5 and in_str.count(':') == 2:
                in_str = in_str[:5]

            out_str = str(out_val).split(' ')[-1] if pd.notnull(out_val) else ""
            if len(out_str) > 5 and out_str.count(':') == 2:
                out_str = out_str[:5]

            daily_dict[tgl_Key] = {
                'dws': str(row['DWS']) if pd.notnull(row['DWS']) else "",
                'in': in_str if in_str != 'NaT' else "",
                'out': out_str if out_str != 'NaT' else ""
            }

        # Isi Grid Absensi Harian (Baris 14 s/d 44)
        for r in range(14, 45):
            cell_date = ws.Cells(r, 3).Value # Kolom C (Date)
            if cell_date is not None:
                ws.Cells(r, 2).Value = nrp_str 
                
                if isinstance(cell_date, pd.Timestamp) or hasattr(cell_date, 'strftime'):
                    tgl_template_str = cell_date.strftime('%d.%m.%Y')
                else:
                    parsed_dt = pd.to_datetime(cell_date, errors='coerce')
                    tgl_template_str = parsed_dt.strftime('%d.%m.%Y') if pd.notnull(parsed_dt) else str(cell_date).strip()

                if tgl_template_str in daily_dict:
                    ws.Cells(r, 4).Value = daily_dict[tgl_template_str]['dws']
                    ws.Cells(r, 5).Value = daily_dict[tgl_template_str]['in']
                    ws.Cells(r, 6).Value = daily_dict[tgl_template_str]['out']

        # Format Nama File PDF otomatis (Contoh: TS_80001861_AL_FAIZAL_WIJAYA_Agustus_2026.pdf)
        nama_bersih = nama.replace(" ", "_")
        pdf_filename = f"TS_{nrp_str}_{nama_bersih}_{periode_str}.pdf"
        pdf_path = os.path.join(pdf_output_dir, pdf_filename)
        
        wb.ExportAsFixedFormat(0, pdf_path)
        wb.Close(False)

        # Update Database & PIN
        cursor.execute("SELECT pin FROM users WHERE nrp=?", (nrp_str,))
        result = cursor.fetchone()
        pin_karyawan = result[0] if result else generate_pin()
        
        cursor.execute('''
            INSERT OR REPLACE INTO users (nrp, nama, pin, perusahaan, pdf_path) 
            VALUES (?, ?, ?, ?, ?)
        ''', (nrp_str, nama, pin_karyawan, 'NON STAFF', pdf_path))
        conn.commit()
        
        daftar_akses.append({
            'Perusahaan': 'NON STAFF',
            'NRP': nrp_str,
            'Nama': nama,
            'PIN Akses': pin_karyawan
        })

    print("\n=== SEMUA PROSES SELESAI ===")
    print("Cek folder 3_Output_PDF/NonStaff untuk melihat hasil PDF-nya.")

except Exception as e:
    print(f"Terjadi kesalahan: {e}")

finally:
    if excel:
        try:
            excel.Quit()
        except:
            pass
    conn.close()

# Buat File Rekap Akses
if 'daftar_akses' in locals() and daftar_akses:
    df_akses = pd.DataFrame(daftar_akses)
    df_akses.to_excel('Daftar_Akses_Portal.xlsx', index=False)