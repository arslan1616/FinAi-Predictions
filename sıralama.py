import pandas as pd
import os

folder_path = 'D:/dolarbilanco'


excel_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

data = []

for file in excel_files:
    
    df = pd.read_excel(os.path.join(folder_path, file))
    stock_name = os.path.splitext(file)[0]

    
    net_profit_change_row = df[(df.iloc[:, 0] == 'Net Faaliyet Kar/Zararı') |
                               (df.iloc[:, 0] == '  F-Dönem Net Karı ') |
                               (df.iloc[:, 0] == 'XVII. SÜRDÜRÜLEN FAALİYETLER DÖNEM NET K/Z (XV±XVI)')]
    if not net_profit_change_row.empty:
        net_profit_change = net_profit_change_row.iloc[0].dropna().iloc[-1]
    else:
        net_profit_change = None

    
    equity_change_row = df[(df.iloc[:, 0] == 'Özkaynaklar') |
                           (df.iloc[:, 0] == ' Özsermaye Toplamı') |
                           (df.iloc[:, 0] == 'XVI. ÖZKAYNAKLAR') |
                           (df.iloc[:, 0] == 'ÖZKAYNAK')]
    if not equity_change_row.empty:
        equity_change = equity_change_row.iloc[0].dropna().iloc[-1]
    else:
        equity_change = None

    
    if net_profit_change is not None and equity_change is not None:
        average_change = (net_profit_change + equity_change) / 2
    else:
        average_change = None

    data.append((stock_name, net_profit_change, equity_change, average_change))


df = pd.DataFrame(data, columns=['Hisse', 'Net Kar Değişimi', 'Özkaynak Değişimi', 'Ortalama Değişim'])
df = df.sort_values(by='Ortalama Değişim', ascending=False)


df = df.fillna(0)


df['Net Kar Değişimi'] = df['Net Kar Değişimi'].apply(lambda x: f"{x:.2f}")
df['Özkaynak Değişimi'] = df['Özkaynak Değişimi'].apply(lambda x: f"{x:.2f}")
df['Ortalama Değişim'] = df['Ortalama Değişim'].apply(lambda x: f"{x:.2f}")


output_file = 'sıralama.xlsx'
df.to_excel(output_file, index=False)

print(f"Veriler başarıyla sıralandı ve '{output_file}' dosyasına kaydedildi.")