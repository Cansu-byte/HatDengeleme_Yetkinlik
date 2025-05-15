import pulp
import pandas as pd
from pulp import LpProblem, LpStatus, value

# Excel'den verileri çekme
excel_path = r"C:\Users\cansu\OneDrive\Masaüstü/UPK_VHAT.xlsx"
personel_df = pd.read_excel(excel_path, sheet_name="Personel_Kumesi")
makine_df = pd.read_excel(excel_path, sheet_name="Makine_Kumesi")
operasyon_df = pd.read_excel(excel_path, sheet_name="Operasyon_Kumesi")
süre_df = pd.read_excel(excel_path, sheet_name="Sure_Kumesi")

# Kümeler ve parametreler
P = list(personel_df['Personel_ID'])
M = list(makine_df['Makine_ID'])
O = list(operasyon_df['Operasyon_ID'])

# Operasyon gerçekleşme adetleri
Adet = {row['Operasyon_ID']: row['Adet'] for _, row in operasyon_df.iterrows()}

# Çevrim süresi sınırı (her operasyon için 13.1 dakika)
Ck = 13.1
# Personel çalışma süresi (günde 430 dakika)
Ai = {p: 430 for p in P}

# Süre matrisini oluşturma
Sj = {}
for _, row in süre_df.iterrows():
    personel = row['Personel_ID']
    for operasyon_id in range(1, 29):
        süre = row.get(operasyon_id, None)
        if süre is not None:
            Sj[(personel, operasyon_id)] = süre

def convert_to_float(value):
    try:
        return float(value)
    except ValueError:
        corrected = value.replace('o', '0').replace('O', '0')
        return float(corrected)

Sj = {key: convert_to_float(value) for key, value in Sj.items()}

# Model oluşturma
model = pulp.LpProblem("Hat_Dengeleme", pulp.LpMinimize)

# Karar değişkeni
Xijk = pulp.LpVariable.dicts("Xijk", (P, O, M), cat="Binary")

# Amaç fonksiyonu
model += pulp.lpSum(Xijk[i][j][k] * Sj.get((i, j), 1) for i in P for j in O for k in M)

# Kısıtlar
# Kısıt 1: Her makine yalnızca bir operasyon gerçekleştirebilir
for k in M:
    model += pulp.lpSum(Xijk[i][j][k] for i in P for j in O) <= 1, f"Makine_tektek_operasyon_{k}"

# Kısıt 2: Her operasyon, gerçekleşme adedi kadar atama almalı
for j in O:
    model += pulp.lpSum(Xijk[i][j][k] for i in P for k in M) == Adet[j], f"Operasyon_adet_{j}"

# Kısıt 3: Her personel en fazla bir operasyon alabilir
for i in P:
    model += pulp.lpSum(Xijk[i][j][k] for j in O for k in M) <= 1, f"Personel_tektek_operasyon_{i}"
    
# Kısıt 4: Her personel en az bir operasyona atanmalı 
for i in P:
    model += pulp.lpSum(Xijk[i][j][k] for j in O for k in M) >= 1, f"Personel_en_az_bir_is_{i}"


# Kısıt 5: Her makine en az bir operasyona atanmalı
for k in M:
    model += pulp.lpSum(Xijk[i][j][k] for i in P for j in O) >= 1, f"Makine_en_az_bir_is_{k}"

# Kısıt 6: Çevrim süresi sınırı
for k in M:
    for j in O:
        model += pulp.lpSum(Xijk[i][j][k] * Sj.get((i, j), 1) for i in P) <= Ck, f"Cevrim_suresi_{j}_{k}"

# Kısıt 7: Personelin günlük toplam çalışma süresi 430 dakikayı geçemez
for i in P:
    model += pulp.lpSum(Xijk[i][j][k] * Sj.get((i, j), 0) for j in O for k in M) <= Ai[i], f"Calisma_suresi_limiti_{i}"

# Modeli çöz
model.solve()

# Çözüm çıktısı
if pulp.LpStatus[model.status] == "Optimal":
    print("Optimum çözüm bulundu:")
    for i in P:
        for j in O:
            for k in M:
                if pulp.value(Xijk[i][j][k]) > 0.5:
                    print(f"Personel {i} operasyon {j} makinada {k} üzerinde çalışacak.")
else:
    print("Optimal çözüm bulunamadı. Çözüm durumu:", pulp.LpStatus[model.status])

print("Amaç fonksiyonu değeri:", value(model.objective))
