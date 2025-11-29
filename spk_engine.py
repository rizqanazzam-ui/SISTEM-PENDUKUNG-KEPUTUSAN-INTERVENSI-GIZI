import pandas as pd
import numpy as np
import os
import json

# ==========================================
# KONFIGURASI
# ==========================================
FILE_PATH = r"D:\TEKNIK KOMPUTER\S5\Sistem Pengambilan keputusan\SPK\Artikel\AXCEL_PROJEK_SPK_KELOMPOK_12.xlsx"
SHEET_NAME = " ProsesingSAW"
CONFIG_FILE = "ahp_config.json" # File baru untuk simpan matriks

KRITERIA = [
    "IbuHamil_Normal",              # Index 0
    "Bayi_GiziNormal",              # Index 1
    "IbuHamil_Periksa",             # Index 2
    "IbuHamil_TTD",                 # Index 3
    "Anak_Terpantau_TumbuhKembang", # Index 4
    "Anak_GiziBuruk"                # Index 5
]

JENIS_KRITERIA = [False, False, False, True, True, True]

# Nilai Default (Jika belum ada settingan)
DEFAULT_MATRIX = {
    "0-1": 1.0, "0-2": 1.5, "0-3": 1.5, "0-4": 0.6, "0-5": 0.6,
    "1-2": 1.5, "1-3": 1.5, "1-4": 0.6, "1-5": 0.6,
    "2-3": 1.0, "2-4": 0.4, "2-5": 0.4,
    "3-4": 0.4, "3-5": 0.4,
    "4-5": 1.0
}

def load_matrix_values():
    """Membaca nilai perbandingan dari file JSON"""
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return DEFAULT_MATRIX

def save_matrix_values(new_values):
    """Menyimpan nilai perbandingan ke file JSON"""
    with open(CONFIG_FILE, 'w') as f:
        json.dump(new_values, f)

def get_ahp_weights():
    n = len(KRITERIA)
    matriks = np.ones((n, n))
    
    # Ambil nilai dari file konfigurasi
    values = load_matrix_values()

    # Mengisi Matriks dari Dictionary Values
    # Format key dictionary: "baris-kolom" (misal "0-1" artinya baris 0 banding kolom 1)
    for i in range(n):
        for j in range(i+1, n):
            key = f"{i}-{j}"
            val = float(values.get(key, 1.0)) # Default 1 jika error
            
            matriks[i, j] = val
            matriks[j, i] = 1 / val

    # Normalisasi AHP
    matriks_norm = matriks / matriks.sum(axis=0)
    bobot = matriks_norm.mean(axis=1)

    # Uji Konsistensi
    weighted_sum = matriks @ bobot
    lambda_max = (weighted_sum / bobot).mean()
    CI = (lambda_max - n) / (n - 1)
    RI_dict = {1: 0, 2: 0, 3: 0.58, 4: 0.90, 5: 1.12, 6: 1.24}
    RI = RI_dict.get(n, 1.24)
    CR = CI / RI if RI != 0 else 0

    konsistensi_data = {
        "Lambda Max": lambda_max,
        "CI": CI,
        "RI": RI,
        "CR": CR,
        "Status": "KONSISTEN" if CR < 0.1 else "TIDAK KONSISTEN"
    }

    return bobot, konsistensi_data, matriks

def run_spk_calculation():
    if not os.path.exists(FILE_PATH):
        # Buat dummy jika file tidak ada (untuk handling error awal)
        return pd.DataFrame(), [], {'Status': 'Error', 'CR': 0}

    df = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)
    df = df.drop_duplicates(subset=["Desa"], keep='first')

    bobot_ahp, status_konsistensi, _ = get_ahp_weights()
    
    normalisasi = df[KRITERIA].copy()
    for i, col in enumerate(KRITERIA):
        if JENIS_KRITERIA[i] == True:
            normalisasi[col] = df[col] / df[col].max()
        else:
            normalisasi[col] = df[col].min() / df[col]

    df['Skor_Akhir'] = (normalisasi * bobot_ahp).sum(axis=1)
    
    hasil = df[["Desa"] + KRITERIA + ["Skor_Akhir"]].copy()
    hasil = hasil.sort_values(by="Skor_Akhir", ascending=False).reset_index(drop=True)
    hasil["Ranking"] = hasil.index + 1

    return hasil, bobot_ahp, status_konsistensi