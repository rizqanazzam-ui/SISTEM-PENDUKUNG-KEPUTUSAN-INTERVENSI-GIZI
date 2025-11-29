import pandas as pd

cols = [
    'Desa',
    'IbuHamil_Normal',
    'Bayi_GiziNormal',
    'IbuHamil_Periksa',
    'IbuHamil_TTD',
    'Anak_Terpantau_TumbuhKembang',
    'Anak_GiziBuruk'
]

df = pd.DataFrame(columns=cols)

df.to_excel("template_input_desa.xlsx", index=False)
print("Template berhasil dibuat: template_input_desa.xlsx")
