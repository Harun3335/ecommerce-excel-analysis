import pandas as pd
import matplotlib.pyplot as plt
# Excel dosyasını oku
df = pd.read_excel('ecommerce_sales.xlsx')

# Negatif Quantity iptal anlamına gelir → sadece pozitif satışlar
df = df[df['Quantity'] > 0]

# Satış tutarını hesapla
df['SalesAmount'] = df['Quantity'] * df['UnitPrice']

# İlk 5 satırı göster
print(df.head())

total_sales = df['SalesAmount'].sum()
print(f"Toplam Satış Tutarı: {total_sales:,.2f}")

top_selling = (
    df.groupby('Description')['Quantity']
    .sum()
    .sort_values(ascending=False)
    .head(10)
)
print("En Çok Satan Ürünler:")
print(top_selling)

top_revenue = (
    df.groupby('Description')['SalesAmount']
    .sum()
    .sort_values(ascending=False)
    .head(10)
)
print("En Çok Kazandıran Ürünler:")
print(top_revenue)

# Satış tarihini datetime tipine çevir
df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])

# Ay bazlı grupla
df['Month'] = df['InvoiceDate'].dt.to_period('M')
monthly_sales = df.groupby('Month')['SalesAmount'].sum()

print("Aylık Satış Tutarları:")
print(monthly_sales)

# Grafik çizimi
plt.figure(figsize=(10, 6))
monthly_sales.plot(kind='line', marker='o', color='green')
plt.title('Aylık Toplam Satış Tutarı')
plt.xlabel('Ay')
plt.ylabel('Satış Tutarı (GBP)')
plt.grid(True)
plt.tight_layout()
plt.show()

country_sales = df.groupby('Country')['SalesAmount'].sum().sort_values(ascending=False)

print("Ülke Bazlı Satış Tutarları (İlk 10):")
print(country_sales.head(10))

# İlk 10 ülkeyi al
top_countries = country_sales.head(10)

# Grafik çizimi
plt.figure(figsize=(10, 6))
top_countries.plot(kind='bar', color='coral')
plt.title('En Çok Satış Yapılan 10 Ülke')
plt.xlabel('Ülke')
plt.ylabel('Satış Tutarı (GBP)')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()


# Negatif Quantity olanlar: iptaller
returns_df = df[df['Quantity'] < 0]

# Toplam iade satır sayısı
total_returns = returns_df.shape[0]

# Toplam iade tutarı
total_return_amount = returns_df['SalesAmount'].sum()

print(f"Toplam İade Sayısı: {total_returns}")
print(f"Toplam İade Tutarı: {total_return_amount:,.2f}")

#en çok iade edilen ürünler.

top_returned_products = (
    returns_df.groupby('Description')['Quantity']
    .sum()
    .sort_values()
    .head(10)
)

print("En Çok İade Edilen Ürünler:")
print(top_returned_products)


# Excel dosyasına yazmak için Writer kullanıyoruz
with pd.ExcelWriter('ecommerce_rapor.xlsx', engine='xlsxwriter') as writer:
    # 1. Tüm satışlar (temiz veri)
    df.to_excel(writer, sheet_name='Tüm Satışlar', index=False)

    # 2. En çok satan ürünler
    top_selling.to_frame(name='Adet').to_excel(writer, sheet_name='En Çok Satanlar')

    # 3. En çok kazandıran ürünler
    top_revenue.to_frame(name='Ciro').to_excel(writer, sheet_name='En Çok Kazandıranlar')

    # 4. Aylık satışlar
    monthly_sales.to_frame(name='Aylık Satış').to_excel(writer, sheet_name='Aylık Trend')

    # 5. Ülke bazlı satışlar
    country_sales.to_frame(name='Ülke Satışı').to_excel(writer, sheet_name='Ülke Satış')

    # 6. İade edilen ürünler
    top_returned_products.to_frame(name='İade Adedi').to_excel(writer, sheet_name='En Çok İade')

print("✅ Otomatik Excel raporu 'ecommerce_rapor.xlsx' oluşturuldu.")

import matplotlib.pyplot as plt

# Aylık satış grafiği çiz
plt.figure(figsize=(10, 6))
monthly_sales.plot(kind='line', marker='o', color='green')
plt.title('Aylık Toplam Satış Tutarı')
plt.xlabel('Ay')
plt.ylabel('Satış Tutarı (GBP)')
plt.grid(True)
plt.tight_layout()

# Görseli dosyaya kaydet
plt.savefig('monthly_sales_chart.png')
plt.close()

print("✅ Grafik 'monthly_sales_chart.png' olarak kaydedildi.")

