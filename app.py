import os
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, session, send_from_directory, flash

app = Flask(__name__)
app.secret_key = 'kunci_rahasia_portal_timesheet_sangat_aman'  # Ganti dengan key acak

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
    if os.path.exists(folder_path):
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                if file.lower().endswith('.pdf') and user_nrp in file:
                    full_path = os.path.join(root, file)
                    rel_path = os.path.relpath(full_path, folder_path)
                    sub_path_url = f"{target_subfolder}/{rel_path.replace(os.sep, '/')}"
                    
                    # Ambil nama folder induk terdekat sebagai informasi Periode/Kategori
                    folder_parts = rel_path.split(os.sep)
                    periode_or_category = folder_parts[-2] if len(folder_parts) > 1 else target_subfolder
                    
                    my_files.append({
                        "filename": file,
                        "subfolder": sub_path_url,
                        "category": periode_or_category  # Menampilkan nama sub-folder periode (misal: September_2026)
                    })
                
    return render_template('dashboard_employee.html', user=user, files=my_files)

@app.route('/admin')
def admin_dashboard():
    if 'user' not in session or not session['user'].get('is_admin'):
        return redirect(url_for('login'))
        
    df = pd.read_excel(ACCESS_EXCEL, dtype=str)
    employees = df.to_dict(orient='records')
    
    return render_template('dashboard_admin.html', employees=employees)

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

@app.route('/logout')
def logout():
    session.clear()
    flash('Anda telah keluar dari portal.', 'info')
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True, port=5000)