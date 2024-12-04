


import pulp
import pandas as pd
from pulp import LpProblem

# Excel'den verileri çekme
excel_path = r"C:\Users\cansu\OneDrive\Masaüstü/YON.xlsx"
personel_df = pd.read_excel(excel_path, sheet_name="Personel Kumesi")
makine_df = pd.read_excel(excel_path, sheet_name="Makine Kumesi")
operasyon_df = pd.read_excel(excel_path, sheet_name="Operasyon Kumesi")
performans_df = pd.read_excel(excel_path, sheet_name="Performans Kumesi")
süre_df = pd.read_excel(excel_path, sheet_name="Sure Kumesi")
zorluk_df = pd.read_excel(excel_path, sheet_name="Zorluk Kumesi")


# Kümeler ve parametreler
P = list(personel_df['Personel'])
M = list(makine_df['Makine'])
O = list(operasyon_df['Operasyon_ID'])

# Çevrim süresi sınırı (her operasyon için 13.1 dakika)
Ck = 13.1
# Personel çalışma süresi (günde 430 dakika)
Ai = {p: 430 for p in P}
# Operasyon zorluk derecesi
Zorluk = {row['Operasyon_ID']: row['Zorluk'] for _, row in zorluk_df.iterrows()}

# Eşik değerin üstünde performansa sahip olanları yetkin olarak tanımlama
performans_esiği = 2
Eij = {}
for _, row in performans_df.iterrows():
    personel = row['Personel']
    for operasyon in range(1, 29):
        performans = row[operasyon]
        if performans >= performans_esiği and performans >= Zorluk[operasyon]:
            Eij[(personel, operasyon)] = 1
        else:
            Eij[(personel, operasyon)] = 0

# Süre matrisini oluşturma
Sj = {}
for _, row in süre_df.iterrows():
    personel = row['Personel']
    for operasyon_id in range(1, 29):
        süre = row.get(operasyon_id, None)
        if süre is not None:
            Sj[(personel, operasyon_id)] = süre

print(Sj)
print(Eij)


# Model oluşturma
model = pulp.LpProblem("Hat_Dengeleme", pulp.LpMinimize)

# Karar değişkeni: Xijk, i. personel j. operasyonu k. makinada yapacak mı?
Xijk = pulp.LpVariable.dicts("Xijk", (P, O, M), cat="Binary")

# Amaç fonksiyonu: Çevrim süresini minimize et
model += pulp.lpSum(Xijk[i][j][k] * Sj.get((i, j), 1) for i in P for j in O for k in M)

# Kısıtlar
# Kısıt 1: Her makine yalnızca bir operasyon gerçekleştirecek
for k in M:
    model += pulp.lpSum(Xijk[i][j][k] for i in P for j in O) <= 1, f"Makine_tektek_operasyon_{k}"

# Kısıt 2: Her operasyon en az bir kez atanmalı
for j in O:
    model += pulp.lpSum(Xijk[i][j][k] for i in P for k in M) >= 1, f"Operasyon_atama_{j}"

# Kısıt 3: Her personel yalnızca bir operasyona atanacak
for i in P:
    model += pulp.lpSum(Xijk[i][j][k] for j in O for k in M) <= 1, f"Personel_tektek_operasyon_{i}"

# Kısıt 4: Her personel en az bir operasyona atanmalı
for i in P:
    model += pulp.lpSum(Xijk[i][j][k] for j in O for k in M) >= 1, f"Personel_en_az_bir_is_{i}"

# Kısıt 5: Her makine en az bir operasyona atanmalı
for k in M:
    model += pulp.lpSum(Xijk[i][j][k] for i in P for j in O) >= 1, f"Makine_en_az_bir_is_{k}"

# Kısıt 6: Yetkinlik matrisi
for i in P:
    for j in O:
        for k in M:
            if Eij.get((i, j), 0) == 1:
                model += Xijk[i][j][k] <= Eij.get((i, j), 0), f"Yetkinlik_{i}_{j}_{k}"

# Kısıt 7: Çevrim süresi sınırı
for k in M:
    for j in O:
        model += pulp.lpSum(Xijk[i][j][k] * Sj.get((i, j), 1) for i in P) <= Ck, f"Cevrim_suresi_{j}_{k}"

# Kısıt 8: Her işçinin toplam çalışma süresi 430 dakikayı aşamaz
for i in P:
    model += pulp.lpSum(Xijk[i][j][k] * Sj.get((i, j), 0) for j in O for k in M) <= Ai[i], f"Calisma_suresi_limiti_{i}"


# Modeli çözme
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
