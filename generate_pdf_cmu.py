import os
import calendar
from datetime import datetime
import pandas as pd
import win32com.client as win32

# --- KONFIGURASI PATH ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PATH_MASTER = os.path.join(BASE_DIR, "1_Data_Master")
PATH_TEMPLATE = os.path.join(BASE_DIR, "2_Template_Timesheet")
PATH_OUTPUT = os.path.join(BASE_DIR, "3_Output_PDF", "CMU")

def generate_cmu_timesheet():
    print("Memulai proses generate Timesheet CMU...")
    
    os.makedirs(PATH_OUTPUT, exist_ok=True)
    file_master = os.path.join(PATH_MASTER, "Data Absensi_CMU.xlsx")
    
    try:
        df_absensi = pd.read_excel(file_master, sheet_name="ABSENSI")
        df_spl = pd.read_excel(file_master, sheet_name="SPL")
        df_finger = pd.read_excel(file_master, sheet_name="DATA FINGER")
        
        # Bersihkan spasi pada nama kolom agar aman
        df_absensi.columns = df_absensi.columns.str.strip()
        df_spl.columns = df_spl.columns.str.strip()
        df_finger.columns = df_finger.columns.str.strip()
        
        # Konversi kolom Tanggal ke objek datetime agar aman dari perbedaan format
        df_finger['Tanggal_Dt'] = pd.to_datetime(df_finger['Tanggal'], errors='coerce')
        if 'TANGGAL' in df_spl.columns:
            df_spl['Tanggal_Dt'] = pd.to_datetime(df_spl['TANGGAL'], errors='coerce')
        elif 'Tanggal' in df_spl.columns:
            df_spl['Tanggal_Dt'] = pd.to_datetime(df_spl['Tanggal'], errors='coerce')
        
        print(f"Berhasil memuat data:")
        print(f"- ABSENSI: {len(df_absensi)} baris")
        print(f"- SPL: {len(df_spl)} baris")
        print(f"- DATA FINGER: {len(df_finger)} baris")
        
    except Exception as e:
        print(f"Gagal membaca file master: {e}")
        return

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        template_path = os.path.join(PATH_TEMPLATE, "Template_Timesheet_CMU.xlsx")
        
        for idx, row in df_absensi.iterrows():
            nrp_raw = row.get('NRP', row.get('nrp', ''))
            nama_raw = row.get('NAMA', row.get('Nama', ''))
            
            nrp = str(nrp_raw).split('.')[0].strip()
            nama = str(nama_raw).strip()
            
            if not nrp or nrp.lower() in ['nan', 'nat', 'none', '']:
                continue
            if not nama or nama.lower() in ['nan', 'nat', 'none', '']:
                continue
                
            jabatan = str(row.get('JABATAN', row.get('Jabatan', ''))).strip()
            departemen = str(row.get('DEPARTEMEN', row.get('Departemen', ''))).strip()
            
            # Ambil nilai Keterangan umum dari sheet ABSENSI
            keterangan_val = ""
            for col_k in ['KETERANGAN', 'Keterangan', 'keterangan']:
                if col_k in row:
                    val_ket = row.get(col_k, "")
                    if pd.notna(val_ket):
                        keterangan_val = str(val_ket).strip()
                    break
            
            # Ambil nilai awal dari kolom "HARI KE-" di sheet ABSENSI
            raw_jsm_awal = row.get('HARI KE-', row.get('HARI KE', row.get('JSM', 1)))
            try:
                jsm_awal = int(float(raw_jsm_awal))
            except:
                jsm_awal = 1
            
            print(f"Memproses CMU: {nrp} - {nama} (Hari Ke- Awal: {jsm_awal})...")
            
            # Filter data finger dan SPL khusus karyawan ini
            df_finger['NRP_Clean'] = df_finger['NRP'].astype(str).str.split('.').str[0].str.strip()
            df_finger_emp = df_finger[df_finger['NRP_Clean'] == nrp]
            
            if 'NRP' in df_spl.columns:
                df_spl['NRP_Clean'] = df_spl['NRP'].astype(str).str.split('.').str[0].str.strip()
                df_spl_emp = df_spl[df_spl['NRP_Clean'] == nrp]
            else:
                df_spl_emp = pd.DataFrame()
            
            # Tentukan Tahun dan Bulan (Default Agustus 2026)
            year, month = 2026, 8
            if not df_finger_emp.empty and df_finger_emp['Tanggal_Dt'].notna().any():
                try:
                    sample_dt = df_finger_emp['Tanggal_Dt'].dropna().iloc[0]
                    year = sample_dt.year
                    month = sample_dt.month
                except:
                    pass
            
            num_days = calendar.monthrange(year, month)[1]
            
            wb = excel.Workbooks.Open(template_path)
            ws = wb.Sheets(1)
            
            # 1. Isi Header Identitas di Kolom D
            ws.Cells(5, 4).Value = nama
            ws.Cells(6, 4).Value = nrp
            ws.Cells(7, 4).Value = jabatan
            ws.Cells(8, 4).Value = departemen
            ws.Cells(9, 4).Value = f"AGUSTUS {year}"
            
            ws.Columns("D").AutoFit()
            
            # --- PEMETAAN KOLOM TANGGAL DI SHEET ABSENSI YANG PRESISI ---
            date_col_mapping = {}
            for col in df_absensi.columns:
                parsed_col_date = pd.to_datetime(col, errors='coerce')
                if pd.notna(parsed_col_date):
                    date_col_mapping[parsed_col_date.day] = col
                else:
                    col_str = str(col).strip()
                    for d in range(1, 32):
                        if col_str.startswith(f"{d:02d}-") or col_str.startswith(f"{d}-") or f" {d} " in col_str:
                            if d not in date_col_mapping:
                                date_col_mapping[d] = col
            
            # 2. Looping baris tanggal otomatis di Kolom B (mulai baris 14 ke bawah)
            start_row = 14
            for i in range(num_days):
                current_row = start_row + i
                day_num = i + 1
                
                target_date = datetime(year, month, day_num)
                tgl_display = target_date.strftime("%d/%m/%Y")
                
                # Masukkan Tanggal ke Kolom B
                ws.Cells(current_row, 2).Value = "'" + tgl_display
                
                # --- SIKLUS JSM / HARI KE- ---
                current_jsm = ((jsm_awal - 1 + i) % 14) + 1
                ws.Cells(current_row, 3).Value = current_jsm # Kolom C (JSM)
                
                # --- AMBIL DATA ROSTER (D / N / OFF / C / CTN / I / A / MCU) ---
                roster_val = ""
                if day_num in date_col_mapping:
                    val_r = row.get(date_col_mapping[day_num], "")
                    if pd.notna(val_r):
                        roster_val = str(val_r).strip().upper()
                
                if not roster_val or roster_val == 'NAN':
                    try:
                        col_idx = 6 + day_num
                        if col_idx < len(df_absensi.columns):
                            val_r = row.iloc[col_idx]
                            if pd.notna(val_r):
                                roster_val = str(val_r).strip().upper()
                    except:
                        pass
                
                ws.Cells(current_row, 4).Value = roster_val
                
                # --- AMBIL JAM FINGER (SHIFT I & SHIFT II) ---
                match_finger = df_finger_emp[df_finger_emp['Tanggal_Dt'].dt.date == target_date.date()]
                
                clock_in = ""
                clock_out = ""
                
                if not match_finger.empty:
                    p10_rows = match_finger[match_finger['Kode'].astype(str).str.strip().str.upper() == 'P10']
                    if not p10_rows.empty:
                        clock_in = str(p10_rows.iloc[0]['Jam']).strip()
                        
                    p20_rows = match_finger[match_finger['Kode'].astype(str).str.strip().str.upper() == 'P20']
                    if not p20_rows.empty:
                        clock_out = str(p20_rows.iloc[-1]['Jam']).strip()
                
                if clock_in and clock_in.lower() != 'nan' and clock_in != '':
                    if roster_val == 'D':
                        ws.Cells(current_row, 5).Value = clock_in  # Shift I - IN (Kolom E)
                    elif roster_val == 'N':
                        ws.Cells(current_row, 7).Value = clock_in  # Shift II - IN (Kolom G)
                        
                if clock_out and clock_out.lower() != 'nan' and clock_out != '':
                    if roster_val == 'D':
                        ws.Cells(current_row, 6).Value = clock_out # Shift I - OUT (Kolom F)
                    elif roster_val == 'N':
                        ws.Cells(current_row, 8).Value = clock_out # Shift II - OUT (Kolom H)
                
                # --- AMBIL DATA LEMBUR (IN & OUT DARI SHEET SPL) ---
                if not df_spl_emp.empty and 'Tanggal_Dt' in df_spl_emp.columns:
                    match_spl = df_spl_emp[df_spl_emp['Tanggal_Dt'].dt.date == target_date.date()]
                    if not match_spl.empty:
                        spl_in = ""
                        spl_out = ""
                        
                        try:
                            val_sin = match_spl.iloc[0, 13]  # Start Akhir Shift
                            val_sout = match_spl.iloc[0, 14] # End Akhir Shift
                            
                            if pd.notna(val_sin):
                                spl_in = str(val_sin).strip()
                            if pd.notna(val_sout):
                                spl_out = str(val_sout).strip()
                        except:
                            pass
                        
                        if spl_in and spl_in.lower() != 'nan':
                            ws.Cells(current_row, 9).Value = spl_in
                        if spl_out and spl_out.lower() != 'nan':
                            ws.Cells(current_row, 10).Value = spl_out
                            
                # --- TENTUKAN KETERANGAN HARIAN BERDASARKAN STATUS ROSTER ---
                final_keterangan = ""
                if roster_val in ['C', 'CUTI']:
                    final_keterangan = "CUTI"
                elif roster_val in ['CTN']:
                    final_keterangan = "CUTI TAHUNAN"
                elif roster_val in ['I', 'IZIN']:
                    final_keterangan = "IZIN"
                elif roster_val in ['A', 'ALPHA']:
                    final_keterangan = "ALPHA"
                elif roster_val in ['MCU']:
                    final_keterangan = "MCU"
                elif roster_val in ['OFF']:
                    final_keterangan = "OFF"
                else:
                    final_keterangan = keterangan_val
                
                if final_keterangan and final_keterangan.lower() != 'nan':
                    ws.Cells(current_row, 11).Value = final_keterangan
                
           # --- PEMBUATAN STR PERIODE ---
            # Nama bulan dalam Bahasa Indonesia / Bahasa Inggris
            nama_bulan = [
                "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
                "Juli", "Agustus", "September", "Oktober", "November", "Desember"
            ]
            periode_str = f"{nama_bulan[month]}_{year}"  # Contoh hasil: "Agustus_2026"
            
            # 3. Export ke PDF dengan Periode & Sub-Folder Otomatis
            nama_clean = nama.replace(' ', '_').replace('/', '_')
    
            # 1. Tambahkan informasi periode di nama file
            pdf_filename = f"Timesheet_CMU_{nrp}_{nama_clean}_{periode_str}.pdf"
    
            # 2. Buat sub-folder berdasarkan periode (misal: .../CMU/Agustus_2026/)
            period_folder = os.path.join(PATH_OUTPUT, periode_str)
            os.makedirs(period_folder, exist_ok=True)
    
            # 3. Gabungkan path output akhir
            pdf_path = os.path.abspath(os.path.join(period_folder, pdf_filename))

            ws.ExportAsFixedFormat(0, pdf_path)
            print(f"✅ Berhasil export: {pdf_filename} ke folder {periode_str}")

            wb.Close(False)
            
        print("\nSelesai! Semua PDF Timesheet CMU berhasil digenerate.")
        
    except Exception as e:
        print(f"Terjadi kesalahan pada Excel: {e}")
        
    finally:
        excel.Quit()

if __name__ == "__main__":
    generate_cmu_timesheet()