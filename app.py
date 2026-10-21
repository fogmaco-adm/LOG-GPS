import os
import pandas as pd
import io
import zipfile
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, flash, send_file
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

app = Flask(__name__)
app.secret_key = 'kunci_rahasia_portal_timesheet_sangat_aman'  # Ganti dengan key acak

import sqlite3

def get_db_connection():
    conn = sqlite3.connect('database.db')
    conn.row_factory = sqlite3.Row
    return conn

# Buat tabel status jika belum ada saat aplikasi berjalan
def init_db():
    conn = get_db_connection()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS timesheet_status (
            file_name TEXT PRIMARY KEY,
            status TEXT
        )
    ''')
    conn.commit()
    conn.close()

# Panggil fungsi inisialisasi database
init_db()

def update_file_status(file_name, status):
    conn = get_db_connection()
    conn.execute('''
        INSERT INTO timesheet_status (file_name, status)
        VALUES (?, ?)
        ON CONFLICT(file_name) DO UPDATE SET status = ?
    ''', (file_name, status, status))
    conn.commit()
    conn.close()

def get_file_status(file_name):
    conn = get_db_connection()
    row = conn.execute('SELECT status FROM timesheet_status WHERE file_name = ?', (file_name,)).fetchone()
    conn.close()
    return row['status'] if row else 'PENDING'

# --- KONFIGURASI PATH ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ACCESS_EXCEL = os.path.join(BASE_DIR, "Daftar_Akses_Portal.xlsx")
PATH_OUTPUT_ROOT = os.path.join(BASE_DIR, "3_Output_PDF")

def get_user_data(nrp_input, pin_input):
    """Mengecek kecocokan NRP dan PIN dari file Excel master"""
    if not os.path.exists(ACCESS_EXCEL):
        return None
    
    df = pd.read_excel(ACCESS_EXCEL, dtype=str)
    for _, row in df.iterrows():
        if str(row.get('NRP', '')).strip() == str(nrp_input).strip() and str(row.get('PIN Akses', '')).strip() == str(pin_input).strip():
            return {
                "perusahaan": row.get('Perusahaan'),
                "nrp": row.get('NRP'),
                "nama": row.get('Nama')
            }
    return None

@app.route('/')
def index():
    if 'user' in session:
        if session['user'].get('is_admin'):
            return redirect(url_for('admin_dashboard'))
        return redirect(url_for('employee_dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        nrp = request.form.get('nrp').strip()
        pin = request.form.get('pin').strip()
        
        # Login khusus Admin
        if nrp.lower() == 'admin' and pin == 'admin123':
            session['user'] = {
                "nama": "Administrator",
                "nrp": "ADMIN",
                "is_admin": True
            }
            flash('Berhasil login sebagai Administrator!', 'success')
            return redirect(url_for('admin_dashboard'))
            
        # Login Karyawan via Database Excel
        user_data = get_user_data(nrp, pin)
        if user_data:
            user_data['is_admin'] = False
            session['user'] = user_data
            flash(f"Selamat datang, {user_data['nama']}!", 'success')
            return redirect(url_for('employee_dashboard'))
        else:
            flash('NRP atau PIN Akses salah! Periksa kembali.', 'danger')
            
    return render_template('login.html')

@app.route('/dashboard')
def employee_dashboard():
    if 'user' not in session or session['user'].get('is_admin'):
        return redirect(url_for('login'))
    
    user = session['user']
    user_nrp = user['nrp']
    user_company = user['perusahaan']
    
    target_subfolder = ""
    if user_company == "NON STAFF":
        target_subfolder = "NonStaff"
    elif user_company == "BMP":
        target_subfolder = "BMP"
    elif user_company == "CMU":
        target_subfolder = "CMU"
        
    folder_path = os.path.join(PATH_OUTPUT_ROOT, target_subfolder)
    my_files = []
    
    # OS.WALK menelusuri folder utama hingga ke dalam sub-folder periode
    for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf') and user_nrp in file:
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, folder_path)
                    sub_path_url = f"{target_subfolder}/{rel_path.replace(os.sep, '/')}"
                    
                    # Ambil nama folder induk terdekat sebagai informasi Periode/Kategori
                    folder_parts = rel_path.split(os.sep)
                    periode_or_category = folder_parts[-2] if len(folder_parts) > 1 else target_subfolder
                    
                    # Ambil status dari database SQLite
                    file_status = get_file_status(file)
                    
                    my_files.append({
                        "filename": file,
                        "subfolder": sub_path_url,
                        "category": periode_or_category,
                        "status": file_status  # <--- Pastikan status ini ada agar terbaca di HTML
                    })
                
    return render_template('dashboard_employee.html', user=user, files=my_files)

@app.route('/admin')
def admin_dashboard():
    if 'user' not in session or not session['user'].get('is_admin'):
        return redirect(url_for('login'))
        
    status_dict = {}
    try:
        conn = get_db_connection()
        status_rows = conn.execute('SELECT file_name, status FROM timesheet_status').fetchall()
        conn.close()
        status_dict = {row['file_name']: row['status'] for row in status_rows}
    except Exception as e:
        print(f"Error read db: {e}")
    
    # Struktur Data: all_files[Perusahaan][Kategori][Bulan] = [file1, file2, ...]
    all_files = {'BMP': {}, 'CMU': {}, 'NONSTAFF': {}}
    
    if os.path.exists(PATH_OUTPUT_ROOT):
        for root, dirs, files in os.walk(PATH_OUTPUT_ROOT):
            for file in files:
                if file.lower().endswith('.pdf'):
                    file_upper = file.upper()
                    root_upper = root.upper()
                    
                    # 1. Tentukan Perusahaan
                    comp = 'NONSTAFF'
                    if 'BMP' in file_upper or 'BMP' in root_upper:
                        comp = 'BMP'
                    elif 'CMU' in file_upper or 'CMU' in root_upper:
                        comp = 'CMU'
                        
                    # 2. Ambil folder relatif (mengabaikan nama file)
                    rel_dir = os.path.relpath(root, PATH_OUTPUT_ROOT)
                    parts = [p for p in rel_dir.split(os.sep) if p and p != '.']
                    
                    category = 'General'
                    periode = 'Umum'
                    
                    # Daftar kata kunci bulan/tahun untuk validasi folder Periode
                    month_keywords = ['2025', '2026', 'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 
                                      'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER']
                    
                    for part in parts:
                        part_upper = part.upper()
                        # Jika bagian folder ini adalah Perusahaan, abaikan
                        if part_upper in ['BMP', 'CMU', 'NONSTAFF', 'NON_STAFF']:
                            continue
                            
                        # Cek apakah folder ini folder Bulan/Periode
                        if any(m in part_upper for m in month_keywords):
                            periode = part
                        else:
                            category = part

                    file_status = status_dict.get(file, 'PENDING')
                    
                    # Masukkan ke kamus bertingkat: Perusahaan -> Kategori -> Bulan
                    if category not in all_files[comp]:
                        all_files[comp][category] = {}
                    if periode not in all_files[comp][category]:
                        all_files[comp][category][periode] = []
                        
                    all_files[comp][category][periode].append({
                        'file_name': file,
                        'status': file_status,
                        'path': os.path.join(root, file)
                    })
                
    return render_template('dashboard_admin.html', files_by_company=all_files)

# --- ROUTE ADMIN: DOWNLOAD ALL APPROVED PDF PER PERUSAHAAN (ZIP) ---
@app.route('/admin/download_zip/<company>/<path:category>')
def admin_download_zip(company, category):
    if 'user' not in session or not session['user'].get('is_admin'):
        return redirect(url_for('login'))
        
    company = company.upper()
    target_periode = request.args.get('periode')
    
    try:
        conn = get_db_connection()
        approved_rows = conn.execute("SELECT file_name FROM timesheet_status WHERE status = 'APPROVED'").fetchall()
        conn.close()
        
        approved_files = {row['file_name'] for row in approved_rows}
        
        memory_file = io.BytesIO()
        found_count = 0
        
        month_keywords = ['2025', '2026', '2027', 'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI', 
                          'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER']
        
        if os.path.exists(PATH_OUTPUT_ROOT):
            with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root, dirs, files in os.walk(PATH_OUTPUT_ROOT):
                    for file in files:
                        if file in approved_files and file.lower().endswith('.pdf'):
                            file_upper = file.upper()
                            root_upper = root.upper()
                            
                            is_comp_match = False
                            if company == 'BMP' and ('BMP' in file_upper or 'BMP' in root_upper):
                                is_comp_match = True
                            elif company == 'CMU' and ('CMU' in file_upper or 'CMU' in root_upper):
                                is_comp_match = True
                            elif company == 'NONSTAFF' and ('BMP' not in file_upper and 'CMU' not in file_upper and 'BMP' not in root_upper and 'CMU' not in root_upper):
                                is_comp_match = True
                                
                            rel_dir = os.path.relpath(root, PATH_OUTPUT_ROOT)
                            parts = [p for p in rel_dir.split(os.sep) if p and p != '.']
                            
                            file_category = 'General'
                            file_periode = 'Umum'
                            
                            for part in parts:
                                part_upper = part.upper()
                                if part_upper in ['BMP', 'CMU', 'NONSTAFF', 'NON_STAFF']:
                                    continue
                                if any(m in part_upper for m in month_keywords):
                                    file_periode = part
                                else:
                                    file_category = part
                            
                            # Cocokkan kategori dan periode (jika ada filter periode)
                            if is_comp_match and file_category.strip().lower() == category.strip().lower():
                                if not target_periode or target_periode.strip().lower() == file_periode.strip().lower():
                                    file_path = os.path.join(root, file)
                                    if os.path.exists(file_path):
                                        zf.write(file_path, arcname=file)
                                        found_count += 1
                                
        if found_count == 0:
            flash(f"Tidak ada file Approved yang ditemukan untuk {company} - Kategori {category}.", "warning")
            return redirect(url_for('admin_dashboard'))
            
        memory_file.seek(0)
        zip_name = f'Timesheet_Approved_{company}_{category}.zip' if not target_periode else f'Timesheet_Approved_{company}_{category}_{target_periode}.zip'
        
        return send_file(
            memory_file,
            mimetype='application/zip',
            as_attachment=True,
            download_name=zip_name
        )
        
    except Exception as e:
        print(f"ERROR DOWNLOAD ZIP: {e}")
        flash(f"Terjadi kesalahan saat memproses file ZIP: {str(e)}", "danger")
        return redirect(url_for('admin_dashboard'))

# --- ROUTE ADMIN: RESET STATUS TIMESHEET ---
@app.route('/admin/reset_status/<path:file_name>', methods=['POST'])
def admin_reset_status(file_name):
    if 'user' not in session or not session['user'].get('is_admin'):
        return redirect(url_for('login'))
        
    try:
        conn = get_db_connection()
        # Ubah status kembali menjadi PENDING atau hapus rekamannya dari database
        conn.execute("UPDATE timesheet_status SET status = 'PENDING' WHERE file_name = ?", (file_name,))
        conn.commit()
        conn.close()
        
        flash(f"Status untuk file '{file_name}' berhasil di-reset menjadi Pending.", "success")
    except Exception as e:
        flash(f"Gagal mereset status file: {e}", "danger")
        
    return redirect(url_for('admin_dashboard'))

@app.route('/download/<path:subpath>')
def download_file(subpath):
    if 'user' not in session:
        return redirect(url_for('login'))
        
    user = session['user']
    # Validasi keamanan: pastikan non-admin hanya bisa mendownload file miliknya sendiri
    if not user.get('is_admin') and user['nrp'] not in subpath:
        flash('Akses ditolak! Anda tidak memiliki izin mengunduh file ini.', 'danger')
        return redirect(url_for('employee_dashboard'))
        
    directory = PATH_OUTPUT_ROOT
    return send_from_directory(directory, subpath, as_attachment=True)

# --- ROUTE 1: APPROVE TIMESHEET & AUTOMATIC DIGITAL STAMP ---
@app.route('/approve_timesheet', methods=['POST'])
def approve_timesheet():
    if 'user' not in session or session['user'].get('is_admin'):
        return redirect(url_for('login'))
    
    file_id = request.form.get('file_id') # Nama file PDF
    user_nrp = session['user']['nrp']
    user_company = str(session['user'].get('perusahaan', '')).strip().upper()
    
    # Format Tanggal Saja (Tanpa Jam & Tanpa WIB)
    now_str = datetime.now().strftime("%d/%m/%Y")
    stamp_text = f"Digitally Signed by {user_nrp} {now_str}"
    
    try:
        # Cari file PDF secara rekursif menggunakan PATH_OUTPUT_ROOT
        pdf_path = None
        for root, dirs, files in os.walk(PATH_OUTPUT_ROOT):
            if file_id in files:
                pdf_path = os.path.join(root, file_id)
                break
                
        if not pdf_path or not os.path.exists(pdf_path):
            flash(f"File '{file_id}' tidak ditemukan di server!", "danger")
            return redirect(url_for('employee_dashboard'))
            
        # 1. Buat overlay TTD/Stamp menggunakan ReportLab di memori (BytesIO)
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=letter)
        
        # Atur font & ukuran stempel
        can.setFont("Helvetica-Bold", 7.5)
        can.setFillColorRGB(0.05, 0.2, 0.5) # Warna Biru Tua Professional
        
        # 2. Pengaturan Koordinat Presisi Berdasarkan Perusahaan (Dibuat 2 Baris)
        if user_company == "CMU":
            # Baris 1: "Digitally Signed by 10151021"
            # Baris 2: "18/10/2026"
            x_position = 310
            y_position_line1 = 55
            y_position_line2 = 45
            
            can.drawString(x_position, y_position_line1, f"Digitally Signed by {user_nrp}")
            can.drawString(x_position, y_position_line2, now_str)
            
        elif user_company == "BMP":
            x_position = 500
            y_position_line1 = 320
            y_position_line2 = 310
            
            can.drawString(x_position, y_position_line1, f"Digitally Signed by {user_nrp}")
            can.drawString(x_position, y_position_line2, now_str)
            
        else:
            # NonStaff / Default
            x_position = 380
            y_position_line1 = 60
            y_position_line2 = 50
            
            can.drawString(x_position, y_position_line1, f"Digitally Signed by {user_nrp}")
            can.drawString(x_position, y_position_line2, now_str)
            
        can.save()
        
        # 3. Gabungkan Overlay Stamp ke PDF Asli menggunakan PyPDF
        packet.seek(0)
        stamp_pdf = PdfReader(packet)
        existing_pdf = PdfReader(pdf_path)
        output = PdfWriter()
        
        page = existing_pdf.pages[0]
        page.merge_page(stamp_pdf.pages[0])
        output.add_page(page)
        
        # Tambahkan sisa halaman jika ada
        for i in range(1, len(existing_pdf.pages)):
            output.add_page(existing_pdf.pages[i])
            
        # Simpan kembali PDF yang sudah di-stamp
        with open(pdf_path, "wb") as output_stream:
            output.write(output_stream)

        update_file_status(file_id, 'APPROVED')
            
        flash(f"Berhasil! Timesheet '{file_id}' telah disetujui & ditandatangani secara digital.", "success")
        
    except Exception as e:
        flash(f"Gagal memproses tanda tangan digital: {e}", "danger")
        
    return redirect(url_for('employee_dashboard'))

# --- ROUTE 2: SUBMIT CATATAN REVISI ---
@app.route('/submit_revision', methods=['POST'])
def submit_revision():
    if 'user' not in session:
        return redirect(url_for('login'))
        
    file_id = request.form.get('file_id')
    catatan_revisi = request.form.get('catatan_revisi')
    user_nrp = session['user']['nrp']
    
    try:
        # Catat status REVISI beserta isi alasannya agar muncul di Dashboard Admin nanti
        # (Contoh: simpan_catatan_revisi(file_id, user_nrp, catatan_revisi))
        
        flash(f"Catatan revisi untuk '{file_id}' berhasil dikirim ke Admin. Terima kasih!", "warning")
    except Exception as e:
        flash(f"Gagal mengirim catatan revisi: {e}", "danger")
        
    return redirect(url_for('dashboard'))

@app.route('/logout')
def logout():
    session.clear()
    flash('Anda telah keluar dari portal.', 'info')
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True, port=5000)