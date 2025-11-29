from flask import Flask, render_template, request, redirect, url_for, flash
import pandas as pd
import os
import json

# Import fungsi logika dari spk_engine.py
# Pastikan file spk_engine.py ada di folder yang sama
from spk_engine import (
    run_spk_calculation, 
    save_matrix_values, 
    load_matrix_values, 
    KRITERIA, 
    FILE_PATH
)

app = Flask(__name__)
app.secret_key = 'kunci_rahasia_spk_kelompok_12'  # Wajib untuk fitur flash message

# ==========================================
# ROUTE 1: DASHBOARD (HALAMAN UTAMA)
# ==========================================
@app.route('/')
def dashboard():
    # Jalankan perhitungan SPK untuk mendapatkan data terbaru
    hasil_df, bobot_ahp, info_cr = run_spk_calculation()
    
    # Siapkan data statistik ringkas
    if not hasil_df.empty:
        top_desa = hasil_df.iloc[0]['Desa']     # Desa Ranking 1
        top_score = hasil_df.iloc[0]['Skor_Akhir']
        total_desa = len(hasil_df)
    else:
        top_desa = "-"
        top_score = 0
        total_desa = 0
    
    # Siapkan data untuk Grafik Chart.js
    chart_labels = KRITERIA
    chart_values = [round(float(b), 4) for b in bobot_ahp]

    return render_template(
        'dashboard.html', 
        top_desa=top_desa,
        top_score=top_score,
        total=total_desa,
        cr_status=info_cr,
        labels=chart_labels,
        values=chart_values
    )

# ==========================================
# ROUTE 2: HASIL PERANGKINGAN
# ==========================================
@app.route('/hasil')
def hasil():
    # Jalankan perhitungan
    hasil_df, _, _ = run_spk_calculation()
    
    # Ubah DataFrame ke Dictionary agar mudah ditampilkan di HTML
    data_spk = hasil_df.to_dict(orient='records')
    
    return render_template('hasil.html', data=data_spk)

# ==========================================
# ROUTE 3: INPUT DATA DESA BARU (FORM MANUAL)
# ==========================================
@app.route('/input', methods=['GET', 'POST'])
def input_data():
    if request.method == 'POST':
        try:
            # 1. Ambil data dari Form HTML
            nama_desa = request.form['desa']
            c1 = float(request.form['c1'])
            c2 = float(request.form['c2'])
            c3 = float(request.form['c3'])
            c4 = float(request.form['c4'])
            c5 = float(request.form['c5'])
            c6 = float(request.form['c6'])

            # 2. Buat DataFrame sementara (sesuai struktur di Excel)
            new_data = pd.DataFrame([{
                'Desa': nama_desa,
                'IbuHamil_Normal': c1,
                'Bayi_GiziNormal': c2,
                'IbuHamil_Periksa': c3,
                'IbuHamil_TTD': c4,
                'Anak_Terpantau_TumbuhKembang': c5,
                'Anak_GiziBuruk': c6
            }])

            # 3. Baca File Excel Lama
            if os.path.exists(FILE_PATH):
                df_lama = pd.read_excel(FILE_PATH, sheet_name=" ProsesingSAW")
                # Gabungkan data (Append)
                df_update = pd.concat([df_lama, new_data], ignore_index=True)
            else:
                df_update = new_data

            # 4. Simpan Kembali ke Excel
            with pd.ExcelWriter(FILE_PATH, engine='openpyxl', mode='w') as writer:
                df_update.to_excel(writer, sheet_name=" ProsesingSAW", index=False)

            flash('Data Desa berhasil ditambahkan dan disimpan ke Excel!', 'success')
            return redirect(url_for('input_data'))

        except Exception as e:
            flash(f'Terjadi kesalahan saat menyimpan: {e}', 'danger')
            return redirect(url_for('input_data'))

    return render_template('input.html')

# ==========================================
# ROUTE 3b: IMPORT DATA DARI FILE EXCEL/CSV
# ==========================================
@app.route('/input/upload', methods=['POST'])
def input_upload():
    file = request.files.get('file')

    if not file or file.filename == '':
        flash('Silakan pilih file terlebih dahulu.', 'warning')
        return redirect(url_for('input_data'))

    filename = file.filename.lower()

    try:
        # Baca file sesuai ekstensi
        if filename.endswith('.xlsx') or filename.endswith('.xls'):
            df_new = pd.read_excel(file)
        elif filename.endswith('.csv'):
            df_new = pd.read_csv(file)
        else:
            flash('Format file tidak didukung. Gunakan .xlsx, .xls, atau .csv.', 'danger')
            return redirect(url_for('input_data'))

        # ====== CEK FORMAT KOLOM (HARUS SAMA DENGAN EXCEL UTAMA) ======
        full_cols = [
            'Desa', 
            'IbuHamil_Normal', 
            'Bayi_GiziNormal', 
            'IbuHamil_Periksa', 
            'IbuHamil_TTD', 
            'Anak_Terpantau_TumbuhKembang', 
            'Anak_GiziBuruk'
        ]

        if not set(full_cols).issubset(set(df_new.columns)):
            flash(
                'Format kolom tidak sesuai. Header file harus persis:\n'
                'Desa, IbuHamil_Normal, Bayi_GiziNormal, IbuHamil_Periksa, '
                'IbuHamil_TTD, Anak_Terpantau_TumbuhKembang, Anak_GiziBuruk',
                'danger'
            )
            return redirect(url_for('input_data'))

        # Ambil hanya kolom yang dibutuhkan dan urutkan
        df_new_mapped = df_new[full_cols].copy()

        # Konversi kolom numerik ke float
        numeric_cols = [
            'IbuHamil_Normal',
            'Bayi_GiziNormal',
            'IbuHamil_Periksa',
            'IbuHamil_TTD',
            'Anak_Terpantau_TumbuhKembang',
            'Anak_GiziBuruk'
        ]
        for col in numeric_cols:
            df_new_mapped[col] = pd.to_numeric(df_new_mapped[col], errors='coerce')

        # Hapus baris yang semua nilai numeriknya NaN
        df_new_mapped = df_new_mapped.dropna(subset=numeric_cols, how='all')

        # ====== GABUNGKAN DENGAN EXCEL LAMA ======
        if os.path.exists(FILE_PATH):
            df_lama = pd.read_excel(FILE_PATH, sheet_name=" ProsesingSAW")
            df_update = pd.concat([df_lama, df_new_mapped], ignore_index=True)
        else:
            df_update = df_new_mapped

        # Simpan kembali ke Excel
        with pd.ExcelWriter(FILE_PATH, engine='openpyxl', mode='w') as writer:
            df_update.to_excel(writer, sheet_name=" ProsesingSAW", index=False)

        flash(f'Berhasil mengimport {len(df_new_mapped)} baris data desa dari file.', 'success')

    except Exception as e:
        print(e)
        flash(f'Terjadi kesalahan saat memproses file: {e}', 'danger')

    return redirect(url_for('input_data'))

# ==========================================
# ROUTE 4: UBAH BOBOT AHP (DINAMIS)
# ==========================================
@app.route('/bobot', methods=['GET', 'POST'])
def ubah_bobot():
    # Generate daftar pasangan kriteria otomatis untuk form
    pasangan_kriteria = []
    for i in range(len(KRITERIA)):
        for j in range(i+1, len(KRITERIA)):
            pasangan_kriteria.append({
                'id': f"{i}-{j}",  # ID unik misal "0-1"
                'kriteria1': KRITERIA[i],
                'kriteria2': KRITERIA[j]
            })

    if request.method == 'POST':
        try:
            # 1. Ambil semua input dari form
            new_values = {}
            for p in pasangan_kriteria:
                key = p['id']
                new_values[key] = float(request.form[key])
            
            # 2. Simpan ke file JSON via spk_engine
            save_matrix_values(new_values)
            
            flash('Bobot AHP berhasil diperbarui! Nilai CR dan Ranking otomatis dihitung ulang.', 'success')
            return redirect(url_for('ubah_bobot'))
            
        except Exception as e:
            flash(f'Gagal menyimpan konfigurasi bobot: {e}', 'danger')

    # Load nilai konfigurasi yang tersimpan saat ini
    current_values = load_matrix_values()
    
    return render_template('bobot.html', pasangan=pasangan_kriteria, values=current_values)

# ==========================================
# MAIN EXECUTION
# ==========================================
if __name__ == '__main__':
    print("===========================================")
    print("   SISTEM SPK GIZI BERJALAN DI PORT 5000   ")
    print("   Buka browser: http://127.0.0.1:5000/    ")
    print("===========================================")
    app.run(debug=True)
