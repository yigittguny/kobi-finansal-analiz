import pandas as pd

data = {
    "Satis": [1000, 1500, 2000],
    "Maliyet": [600, 900, 1200]
}

df = pd.DataFrame(data)

df.to_excel("satis.xlsx", index=False)

print("satis.xlsx oluşturuldu.")