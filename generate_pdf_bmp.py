import os
import pandas as pd
import win32com.client as win32
import subprocess
from datetime import datetime

# Matikan paksa proses Excel yang nyangkut di background agar tidak 'Call was rejected by callee'
os.system("taskkill /f /im excel.exe >nul 2>&1")

# Fungsi helper untuk membulatkan angka ke 1 desimal terdekat (misal: 0.5, 1.0, 1.5)
def round_to_half(val):
    try:
        return round(float(val) * 2) / 2
    except:
        return 0.0

def generate_bmp_timesheets():
    print("Sedang memproses Timesheet BMP & mengambil data dari sheet SPL...")
    
    # 1. Definisikan base_dir
    base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 2. Tentukan path file master dan template
    master_path = os.path.join(base_dir, "1_Data_Master", "Data Absensi_BMP.xlsx")
    template_path = os.path.join(base_dir, "2_Template_Timesheet", "Template_Timesheet_BMP.xlsx")
    
    # 3. Tentukan output_dir secara langsung
    output_dir = os.path.join(base_dir, "3_Output_PDF", "BMP")
    os.makedirs(output_dir, exist_ok=True)
    
    if not os.path.exists(master_path) or not os.path.exists(template_path):
        print("Error: File master atau template tidak ditemukan!")
        return

    print("--- Membaca Data Master & Sheet SPL ---")
    try:
        df_absensi = pd.read_excel(master_path, sheet_name="MASTER ABSENSI")
        df_absensi.columns = [str(c).strip() for c in df_absensi.columns]

        try:
            df_spl = pd.read_excel(master_path, sheet_name="SPL")
            df_spl.columns = [str(c).strip() for c in df_spl.columns]
            print(f"Berhasil memuat sheet SPL dengan {len(df_spl)} baris data.")
        except Exception as e:
            print(f"Peringatan: Sheet SPL tidak ditemukan atau gagal dibaca: {e}")
            df_spl = pd.DataFrame()
            
    except Exception as e:
        print(f"Gagal membaca file master: {e}")
        return

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    
    try:
        wb_template = excel.Workbooks.Open(os.path.abspath(template_path))
        ws_template = wb_template.Sheets(1)
    except Exception as e:
        print(f"Gagal membuka template Excel: {e}")
        excel.Quit()
        return

    # Looping setiap baris karyawan di master absensi
    for index, row_emp in df_absensi.iterrows():
        try:
            nrp = row_emp.get('NRP', '')
            nama = row_emp.get('NAMA', '')
            
            if pd.isna(nrp) or pd.isna(nama) or str(nrp).strip() == "":
                continue
                
            nrp_str = str(int(float(nrp))) if str(nrp).replace('.','',1).isdigit() else str(nrp).strip()
            nama_str = str(nama).strip()
            
            jabatan = str(row_emp.get('JABATAN', row_emp.get('Jabatan', 'ADMIN'))).strip()
            lokasi_kerja = str(row_emp.get('LOKASI', row_emp.get('Lokasi', 'CHPP'))).strip()
            
            # Ambil nilai awal "Hari ke-" dari data master
            raw_hari_ke = row_emp.get('Hari ke-', row_emp.get('HARI KE-', row_emp.get('Hari Ke', 1)))
            try:
                start_hari_ke = int(float(raw_hari_ke))
            except:
                start_hari_ke = 1
            
            # Ambil dan bersihkan Kategori untuk sub-folder
            kategori_raw = str(row_emp.get('KATEGORI', row_emp.get('Kategori', 'Lainnya'))).strip()
            kategori_clean = "".join(c for c in kategori_raw if c not in r'/\:*?"<>|').strip()
    
            # Tentukan path sub-folder kategori di dalam folder output
            folder_kategori = os.path.join(output_dir, kategori_clean)
            if not os.path.exists(folder_kategori):
                os.makedirs(folder_kategori, exist_ok=True)
                print(f"📁 Membuat sub-folder kategori: {kategori_clean}")
                            
            # Ambil Aktifitas Kerja langsung dari kolom master Excel
            aktifitas_master = str(row_emp.get('Aktifitas Kerja', row_emp.get('AKTIFITAS KERJA', 'Pekerjaan operasional rutin'))).strip()
            if pd.isna(row_emp.get('Aktifitas Kerja')) or aktifitas_master == "":
                aktifitas_master = "Pekerjaan operasional rutin"

            print(f"Memproses BMP: {nrp_str} - {nama_str}...")

            # Filter data SPL khusus untuk karyawan ini berdasarkan NRP
            emp_spl = pd.DataFrame()
            if not df_spl.empty:
                spl_nrp_col = next((c for c in df_spl.columns if 'nrp' in c.lower()), None)
                if spl_nrp_col:
                    emp_spl = df_spl[df_spl[spl_nrp_col].astype(str).str.contains(nrp_str, na=False)]

            # Salin template bersih
            ws_template.Copy(Before=ws_template)
            ws = wb_template.ActiveSheet
            ws.Name = f"TS_{nrp_str[:10]}"

            # 1. Isi Header Karyawan
            ws.Cells(6, 7).Value = ":"
            ws.Cells(6, 8).Value = nama_str
            ws.Cells(7, 7).Value = ":"
            ws.Cells(7, 8).Value = jabatan
            ws.Cells(8, 7).Value = ":"
            ws.Cells(8, 8).Value = nrp_str
            ws.Cells(9, 7).Value = ":"
            
            # --- AMBIL PERIODE DARI HEADER EXCEL MASTER ---
            try:
                # Cari kolom tanggal pertama (biasanya di kolom index ke-8 atau setelah 'Hari ke-')
                col_list_temp = list(df_absensi.columns)
                idx_hk = next((i for i, c in enumerate(col_list_temp) if "Hari ke" in str(c)), 7)
                kolom_tgl = str(col_list_temp[idx_hk + 1]).strip()
    
                dt_parsed = pd.to_datetime(kolom_tgl, errors='coerce')
    
                if pd.notna(dt_parsed):
                    bulan_indo = ["", "JANUARI", "FEBRUARI", "MARET", "APRIL", "MEI", "JUNI", 
                                  "JULI", "AGUSTUS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DESEMBER"]
                    periode_header = f"{bulan_indo[dt_parsed.month]} {dt_parsed.year}"
                else:
                    periode_header = "SEPTEMBER 2026"
            except Exception as e:
                periode_header = "SEPTEMBER 2026"

            ws.Cells(9, 8).Value = periode_header
            periode_parts = periode_header.strip().title().split()
            periode_str = "_".join(periode_parts)

            # --- 2. LOOPING TANGGAL 1 SAMPAI 31 (Baris 13 s/d 43) -----
            col_list = list(df_absensi.columns)
            idx_hari_ke = -1
            for idx, c in enumerate(col_list):
                if "Hari ke" in str(c):
                    idx_hari_ke = idx
                    break

            for r in range(13, 44):
                day_num = r - 12  # Tanggal 1, 2, 3... 31
                
                # Hitung Hari Ke- dengan siklus 14 hari (1 s/d 14)
                current_hari_ke = ((start_hari_ke + (day_num - 1) - 1) % 14) + 1
                
                target_col = None
                if idx_hari_ke != -1 and (idx_hari_ke + day_num) < len(col_list):
                    target_col = col_list[idx_hari_ke + day_num]

                # VALIDASI: Jika tanggal 31 tapi kolom di master tidak ada (misal September cuma 30 hari)
                # Maka kosongkan total seluruh sel di baris ini agar tidak menerobos garis tabel
                if target_col is None or target_col not in df_absensi.columns or pd.isna(row_emp.get(target_col)) and day_num > 30:
                    ws.Cells(r, 3).Value = ""  # Tanggal
                    ws.Cells(r, 4).Value = ""  # Lokasi
                    ws.Cells(r, 5).Value = ""  # Hari Ke
                    ws.Cells(r, 6).Value = ""  # Shift
                    ws.Cells(r, 7).Value = ""  # In
                    ws.Cells(r, 8).Value = ""  # Out
                    ws.Cells(r, 9).Value = ""  # SPL In
                    ws.Cells(r, 10).Value = "" # SPL Out
                    ws.Cells(r, 11).Value = "" # Total SPL
                    ws.Cells(r, 12).Value = "" # Total Jam Kerja
                    ws.Cells(r, 13).Value = "" # Aktifitas
                    continue

                in_time = "-"
                out_time = "-"
                shift_str = ""
                
                # Default info baris
                ws.Cells(r, 3).Value = f"{day_num}-Sept-2026"  # Kolom C: Tanggal
                ws.Cells(r, 4).Value = lokasi_kerja          # Kolom D: Lokasi
                ws.Cells(r, 5).Value = current_hari_ke       # Kolom E: Hari Ke-

                if target_col and target_col in df_absensi.columns:
                    shift_val = row_emp[target_col]
                    
                    if pd.notnull(shift_val) and str(shift_val).strip() != "":
                        shift_str = str(shift_val).strip().upper()
                        ws.Cells(r, 6).Value = shift_str         # Kolom F: Shift
                        
                        jabatan_upper = str(jabatan).upper()
                        
                        if 'D' in shift_str:
                            ws.Cells(r, 13).Value = aktifitas_master
                            if 'ADMIN FOG' in jabatan_upper:
                                in_time, out_time = "06:00", "18:00"
                            elif 'ADMIN' in jabatan_upper or 'PROCUREMENT' in jabatan_upper:
                                in_time, out_time = "06:30", "16:30"
                            elif 'FUELMAN' in jabatan_upper or 'OILMAN' in jabatan_upper:
                                in_time, out_time = "06:00", "18:00"
                            elif 'FIELD SUPPORT' in jabatan_upper:
                                in_time, out_time = "06:30", "18:00"
                            else:
                                in_time, out_time = "06:00", "18:00"
                                
                        elif 'N' in shift_str:
                            ws.Cells(r, 13).Value = aktifitas_master
                            if 'ADMIN FOG' in jabatan_upper:
                                in_time, out_time = "18:00", "06:00"
                            elif 'FIELD SUPPORT' in jabatan_upper:
                                in_time, out_time = "19:00", "06:30"
                            else:
                                in_time, out_time = "18:00", "06:00"
                                
                        elif 'OFF' in shift_str or 'L' in shift_str:
                            ws.Cells(r, 13).Value = "LIBUR"
                        elif 'MCU' in shift_str:
                            ws.Cells(r, 13).Value = "MCU"
                        elif 'C' in shift_str or 'CTN' in shift_str:
                            ws.Cells(r, 13).Value = "CUTI"
                        elif shift_str == 'S':
                            ws.Cells(r, 13).Value = "SAKIT"
                        elif shift_str == 'I':
                            ws.Cells(r, 13).Value = "IZIN"
                        elif shift_str == 'A':
                            ws.Cells(r, 13).Value = "ALPHA"
                        else:
                            ws.Cells(r, 13).Value = "HADIR KHUSUS"
                    else:
                        ws.Cells(r, 6).Value = "OFF"
                        ws.Cells(r, 13).Value = "LIBUR"
                else:
                    ws.Cells(r, 6).Value = "OFF"
                    ws.Cells(r, 13).Value = "LIBUR"

                ws.Cells(r, 7).Value = in_time   # Kolom G: Jam Masuk Normal
                ws.Cells(r, 8).Value = out_time  # Kolom H: Jam Keluar Normal

                # --- 3. Ambil Data SPL berdasarkan NRP dan Tanggal ---
                has_spl = False
                spl_start_val = "-"
                spl_end_val = "-"
                total_spl_val = 0.0
                
                if not emp_spl.empty:
                    for _, spl_row in emp_spl.iterrows():
                        try:
                            raw_nrp = str(spl_row.iloc[2]).strip()
                            if raw_nrp.endswith('.0'):
                                raw_nrp = raw_nrp[:-2]
                            
                            current_nrp_clean = str(nrp).strip()
                            if current_nrp_clean.endswith('.0'):
                                current_nrp_clean = current_nrp_clean[:-2]

                            if raw_nrp == current_nrp_clean:
                                raw_tgl_val = spl_row.iloc[4]
                                spl_day = None
                                try:
                                    parsed_date = pd.to_datetime(raw_tgl_val, errors='coerce')
                                    if pd.notnull(parsed_date):
                                        spl_day = parsed_date.day
                                    else:
                                        import re
                                        nums = re.findall(r'\b\d+\b', str(raw_tgl_val))
                                        for n in nums:
                                            if int(n) == day_num:
                                                spl_day = day_num
                                                break
                                except:
                                    pass
                                
                                if spl_day == day_num:
                                    spl_start = spl_row.iloc[12] if len(spl_row) > 12 else None 
                                    spl_end = spl_row.iloc[13] if len(spl_row) > 13 else None   
                                    
                                    if pd.notnull(spl_start) and pd.notnull(spl_end) and str(spl_start).strip() != "":
                                        s_start = str(spl_start).strip()[:5]
                                        s_end = str(spl_end).strip()[:5]
                                        spl_start_val = s_start
                                        spl_end_val = s_end
                                        
                                        try:
                                            t1 = pd.to_datetime(s_start, format='%H:%M', errors='coerce')
                                            t2 = pd.to_datetime(s_end, format='%H:%M', errors='coerce')
                                            diff_hours = (t2 - t1).total_seconds() / 3600.0
                                            total_spl_val = round(diff_hours, 1)
                                        except:
                                            total_spl_val = 0.0

                                        has_spl = True
                                        break
                        except Exception as ex:
                            pass

                if has_spl:
                    ws.Cells(r, 9).Value = spl_start_val   # Kolom I (IN)
                    ws.Cells(r, 10).Value = spl_end_val    # Kolom J (OUT)
                    ws.Cells(r, 11).Value = round(total_spl_val, 1) # Kolom K (Total SPL)
                else:
                    ws.Cells(r, 9).Value = "-"
                    ws.Cells(r, 10).Value = "-"
                    ws.Cells(r, 11).Value = "-"

                # --- 4. Total Jam Kerja (Kolom L) ---
                if out_time != "-":
                    if 'N' in shift_str:
                        ws.Cells(r, 12).Value = f"=IF(H{r}=\"-\", 0, ((1 - G{r} + H{r})) * 24) + IF(OR(K{r}=\"-\", ISBLANK(K{r})), 0, K{r})"
                    else:
                        ws.Cells(r, 12).Value = f"=IF(H{r}=\"-\", 0, ((H{r} - G{r})) * 24) + IF(OR(K{r}=\"-\", ISBLANK(K{r})), 0, K{r})"
                else:
                    ws.Cells(r, 12).Value = "-"

            # --- 5. Export ke PDF dengan Sub-Folder Periode ---
            nama_clean = nama_str.replace(' ', '_').replace('/', '_')
            pdf_filename = f"Timesheet_BMP_{nrp_str}_{nama_clean}_{periode_str}.pdf"
            
            period_folder = os.path.join(folder_kategori, periode_str)
            os.makedirs(period_folder, exist_ok=True)
            
            pdf_path = os.path.abspath(os.path.join(period_folder, pdf_filename))

            ws.ExportAsFixedFormat(0, pdf_path)
            print(f"✅ Berhasil export: {pdf_filename} ke folder {periode_str}")
            ws.Delete()

        except Exception as err:
            print(f"Gagal memproses karyawan {nama_str}: {err}")

    wb_template.Close(False)
    excel.Quit()
    print("\nSelesai! Semua PDF Timesheet BMP berhasil digenerate berdasarkan folder periode.")

if __name__ == "__main__":
    generate_bmp_timesheets()