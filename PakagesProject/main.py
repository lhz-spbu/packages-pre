import pandas as pd
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

# 读取 Excel 文件中的数据
data = pd.read_excel('car.xlsx')
# Проверим структуру данных
print(data.head())

# Отобразим количество встречаемости каждой страны;显示每个国家的数量
country_counts = data['COUNTRY'].value_counts()
print(country_counts)

# Выберем 5 наиболее часто встречающихся стран;选择出现频率最高的5个国家
top_countries = country_counts.nlargest(5).index

# Фильтруем данные только по топовым странам; 仅过滤出最常见的国家的数据
filtered_data = data[data['COUNTRY'].isin(top_countries)]

# Группируем по странам и считаем средние значения цен;按国家分组并计算平均价格
mean_prices = filtered_data.groupby('COUNTRY')['PRICEEACH'].sum()

# Создаем столбчатую диаграмму;创建柱状图
plt.figure(figsize=(10, 6))
mean_prices.plot(kind='bar', color='skyblue')
plt.title('Benefit of Country')
plt.xlabel('Country')
plt.ylabel('Average Price Each')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()

# Группировка данных по компании и подсчет суммы заказов;按公司分组并计算订单总数
total_orders = data.groupby('COUNTRY')['QUANTITYORDERED'].sum()

# Выбор пяти компаний с наибольшей суммой заказов;选择订单总数最多的5个公司
top_companies = total_orders.nlargest(5)
plt.figure(figsize=(10, 6))
top_companies.plot(kind='bar', color='green')
plt.title('Total Quantity Ordered by Top 5 Country')
plt.xlabel('Company')
plt.ylabel('Total Quantity Ordered')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()
print(data.columns)
print(data['MSRP'])

# Убедимся, что столбец с датами в правильном формате datetime；确保日期列采用正确的日期格式
data['ORDERDATE'] = pd.to_datetime(data['ORDERDATE'])

# Проверка структуры данных；数据结构检查
print("Первые несколько строк данных:")
print(data.head())
print("\nСтолбцы данных:")
print(data.columns)

# Группировка данных по 'PRODUCTLINE', суммирование 'QUANTITYORDERED'；
# 按 "PRODUCTLINE "对数据进行分组，汇总 "QUANTITYORDERED "数据
grouped_data = data.groupby('PRODUCTLINE')['QUANTITYORDERED'].sum()

# Выбор пяти наиболее востребованных продуктов；选出五款最受欢迎的产品
top_products = grouped_data.nlargest(5)
print("\nТоп 5 самых востребованных машин:")
print(top_products)

# Построение графика；绘制图表
plt.figure(figsize=(10, 6))
top_products.plot(kind='bar', color='skyblue')
plt.title('Top 5 Most Demanded Cars')
plt.xlabel('Product Line')
plt.ylabel('Total Quantity Ordered')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()


# # Группировка данных по 'CUSTOMERNAME', суммирование 'QUANTITYORDERED'
# total_sales = data.groupby('CUSTOMERNAME')['QUANTITYORDERED'].sum()
#
# # Выбор пяти компаний с наибольшим количеством продаж
# top_companies = total_sales.nlargest(5)
# print("Топ 5 компаний по продажам:")
# print(top_companies)
# # Построение графика
# plt.figure(figsize=(10, 6))
# top_companies.plot(kind='bar', color='green')
# plt.title('Top 5 Companies by Sales Volume')
# plt.xlabel('Company Name')
# plt.ylabel('Total Quantity Ordered')
# plt.xticks(rotation=45)
# plt.tight_layout()
# plt.show()

# Добавляем страну к названию компании для уникальности
data['CompanyWithCountry'] = data['CUSTOMERNAME'] + ' (' + data['COUNTRY'] + ')'

# Группировка данных по новому столбцу 'CompanyWithCountry', суммирование 'QUANTITY ORDERED'；
# 按新列 "CompanyWithCountry "对数据进行分组，汇总 "QUANTITYORDERED"（数量）。
total_sales = data.groupby('CompanyWithCountry')['QUANTITYORDERED'].sum()

# Выбор пяти компаний с наибольшим количеством продаж；评选销售额最高的五家公司
top_companies = total_sales.nlargest(5)
print("Топ 5 компаний по продажам с указанием страны:")
print(top_companies)
# Построение графика‘绘制图表
plt.figure(figsize=(12, 7))
top_companies.plot(kind='bar', color='green')
plt.title('Top 5 Companies by Sales Volume with Country Information')
plt.xlabel('Company Name and Country')
plt.ylabel('Total Quantity Ordered')
plt.xticks(rotation=45)
plt.tight_layout()
plt.show()