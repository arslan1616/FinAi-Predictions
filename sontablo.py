import pandas as pd

hisse_durumlari_df = pd.read_excel('hisse_teknikdurumlari.xlsx')


siralama_df = pd.read_excel('sıralama.xlsx')


merged_df = pd.merge(siralama_df, hisse_durumlari_df[['Hisse Kodu', 'Durum']], how='left', left_on='Hisse', right_on='Hisse Kodu')


merged_df = merged_df.drop(columns=['Hisse Kodu'])


merged_df.to_excel('guncel_siralama.xlsx', index=False)

print("Veriler başarıyla güncellendi ve 'guncel_siralama.xlsx' dosyasına kaydedildi.")