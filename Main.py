import pandas as pd
import lxml
import chardet
import matplotlib.pyplot as plot
from displayfunction import display
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import numpy as np
import openpyxl

#Чтение файла "base.sas7bdat"
base = pd.read_sas("base.sas7bdat", encoding='iso-8859-1')
# Преобразование столбов в численный тип
base['VAR_ID'] = base['VAR_ID'].astype('int64')
base['DOR_ID'] = base['DOR_ID'].astype('int64')

#Функция преобразования столбца датафрейма из байтов в кириллицу, а также удаление дубликатов
def decode_df(df, sort_by='Column_name'):
  if sort_by in df.columns:
    df[sort_by] = df[sort_by].apply(lambda x: x.decode('cp1251') if isinstance(x, bytes) else x)
  df = df.drop_duplicates()
  return df.reset_index(drop=True)

#Чтение файла "var.sas7bdat"
var = pd.read_sas("var.sas7bdat", format='sas7bdat')
var = decode_df(var, sort_by='NAME')
var['VAR_ID'] = var['VAR_ID'].astype('int64')

#Чтение файла "dor_name.xml"
dor_file = pd.read_xml("dor_name.xml", encoding='Windows-1251')

#Чтение файла "operiod.txt"
operiod_file = pd.read_fwf("operiod.txt", encoding='utf-8')

#Словари для прочтённых файлов
dor_dict = {}
operiod_dict = {}
var_dict = {}

#Циклы занесения информации из файлов в словари
index = 1
for row in dor_file.values:
  dor_dict[index] = row[0]
  index += 1

for row in operiod_file.values:
  operiod_dict[row[0]] = row[1]

for row in var[['VAR_ID', 'NAME']].values:
  var_dict[row[0]] = row[1]


#Отчёт 4 ---------------------------------------------------------

filtered_df = base[(base['VAR_ID'] == 11370) & (base['OPERIOD'] == 'H')]

res = {}

for dor_id, dor_name in dor_dict.items():
  #Фильтруем DataFrame по текущей дороге
  dor_df = filtered_df[filtered_df['DOR_ID'] == dor_id]

  #Группируем по годам и считаем нарастающий итог затрат
  grouped_df = dor_df.groupby(dor_df['DATE'].dt.year)['fact'].sum().cumsum()

  #Добавляем результат в словарь
  for year, total_cost in grouped_df.items():
    res[(year, dor_name)] = f'{total_cost:,.0f} ₽'

#Преобразуем результаты в DataFrame
result_df = pd.DataFrame(list(res.items()), columns=['Year_Road', 'Costs'])
result_df[['Year', 'Road']] = result_df['Year_Road'].apply(pd.Series)
result_df = result_df.drop('Year_Road', axis=1)

df_res = result_df
#Группируем DataFrame по годам и дорогам и считаем сумму затрат
result_df_4 = df_res.groupby(['Year', 'Road'])['Costs'].sum().unstack().reset_index()

#Преобразуем столбец 'Costs' в числовой формат
result_df['Costs'] = result_df['Costs'].str.replace(',', '').str.replace(' ₽', '').astype(float)

print(result_df_4)

#Отчёт 5 ---------------------------------------------------------

filtered_df = base[(base['VAR_ID'] == 11370) & (base['OPERIOD'] == 'H')]

res = {}

for dor_id, dor_name in dor_dict.items():
  #Фильтруем data frame по условиям задачи
  dor_df = filtered_df[filtered_df['DOR_ID'] == dor_id]

  #Нахождение среднего, мадианного, максимального значения для каждой дороги по годам
  for year in dor_df['DATE'].dt.year.unique():
    year_df = dor_df[dor_df['DATE'].dt.year == year]
    median = year_df['fact'].median()
    total = year_df['fact'].sum()
    mean = year_df['fact'].mean()

    key = (year, dor_name)
    value = {'Total': total, 'mean': mean, 'median': median}
    res[key] = value

#Преобразуем результаты в DataFrame
result_df_5 = pd.DataFrame(list(res.items()), columns=['Year_Road', 'Statistics'])
result_df_5[['Year', 'Road']] = result_df_5['Year_Road'].apply(pd.Series)
result_df_5 = result_df_5.drop('Year_Road', axis=1)
result_df_5 = pd.concat([result_df_5, result_df_5['Statistics'].apply(pd.Series)], axis=1)
result_df_5 = result_df_5.drop('Statistics', axis=1)

#Выбираем необходимые столбцы и датафрейма
result_df_5 = result_df_5[['Year', 'Road', 'Total', 'mean', 'median']]

#Приводим столбец к целочисленному типу
result_df_5['Year'] = result_df_5['Year'].astype(int)
#Приводим столбцы к числам с плавающей запятой
result_df_5[['Total', 'mean', 'median']] = result_df_5[['Total', 'mean', 'median']].astype(float)

print(result_df_5)

#Отчёт 6 ---------------------------------------------------------
res = {}

#Показатели по варианту
var_ids = [11370, 11400, 11410]
var_names = ['Себестоимость - всего', 'Себестоимость грузовых перевозок', 'Себестоимость пассажирских перевозок - всего']

#Цикл по каждому показателю и его имени
for var_id, var_name in zip(var_ids, var_names):
  #Фильтрация базового DataFrame по текущему показателю
  var_df = base[(base['VAR_ID'] == var_id) & (base['OPERIOD'] == 'H')]

  #Цикл по каждой дороге из словаря dor_dict
  for dor_id, dor_name in dor_dict.items():
    #Фильтрация DataFrame по текущей дороге, показателю и условию OPERIOD == 'H'
    dor_df = var_df[(var_df['DOR_ID'] == dor_id)]

    #Цикл по каждому уникальному году в DataFrame
    for year in dor_df['DATE'].dt.year.unique():
      year_df = dor_df[dor_df['DATE'].dt.year == year]
      total = year_df['fact'].sum()
      key = (year, dor_name, var_name)

      #Форматирование значения
      value = f'{total:,.0f} ₽'
      res[key] = value

#Создание DataFrame на основе словаря res
result_df_6 = pd.DataFrame(list(res.items()), columns=['Year_Road_Var', 'Value'])
result_df_6[['Year', 'Road', 'Var']] = result_df_6['Year_Road_Var'].apply(pd.Series)
result_df_6 = result_df_6.drop('Year_Road_Var', axis=1)
result_df_6 = result_df_6.sort_values(['Year', 'Road', 'Var'])
result_df_6 = result_df_6[['Year', 'Road', 'Var', 'Value']]

#Вывод результатов на экран
print(result_df_6.to_string())

#Отчёт 7 ---------------------------------------------------------

#Показатели по варианту
var_ids = [240021, 240022]
var_names = ['рентабельность грузовых перевозок',
      'рентабельность пассажирских перевозок']

res = {}
for var_id, var_name in zip(var_ids, var_names):

  # Фильтруем data frame по условиям задачи
  var_df = base[(base['VAR_ID'] == var_id) & (base['OPERIOD'] == 'P') & (base['DATE'].dt.year == 2003)]

  for dor in dor_dict.items():
    dor_id = dor[0]
    dor_name = dor[1]

    dor_df = var_df[(var_df['DOR_ID'] == dor_id) & (var_df['VAR_ID'] == var_id) &(var_df['OPERIOD'] == 'P') & (var_df['DATE'].dt.year == 2003)]

    total = dor_df['fact'].sum()

    key = (dor_name, var_name)
    res[key] = total

#Преобразуем результаты в DataFrame
result_df_7 = pd.DataFrame(list(res.items()), columns=['Road_Var', 'Value'])
result_df_7[['Road', 'Var']] = result_df_7['Road_Var'].apply(pd.Series)
result_df_7 = result_df_7.drop('Road_Var', axis=1)
#Переставляем столбцы, чтобы было 4 колонки
result_df_7_pivot = result_df_7.pivot_table(index='Road', columns='Var', values='Value')
print(result_df_7_pivot.to_string())
result_df_7_res = result_df_7.pivot_table(index='Road', columns='Var',
values='Value').reset_index()
#Создаем столбчатую диаграмму
fig, ax = plot.subplots(figsize=(17, 6))

#Определяем цвета для столбцов из палитры "Set1"
colors = plot.cm.Set1(np.linspace(0, 1, len(result_df_7_pivot.columns)))
result_df_7_pivot.plot(kind='bar', stacked=False, legend=False, ax=ax, color=colors)

#Настраиваем отображение графика
plot.xlabel('Наименование дороги')
plot.ylabel('Значение показателей')
plot.title('Значение показателей за 2003 год по дорогам')
plot.xticks(rotation=90)
plot.legend(var_names,loc='center left', bbox_to_anchor=(1, 0.5))
plot.tight_layout()
plot.savefig('graph_7.png')
plot.show()

#Экпорт данных в эксель
writer = pd.ExcelWriter('Результаты.xlsx', engine='openpyxl')

result_df_4.to_excel(writer, sheet_name='4', index=False)
result_df_5.to_excel(writer, sheet_name='5', index=False)
result_df_6.to_excel(writer, sheet_name='6', index=False)
result_df_7_res.to_excel(writer, sheet_name='7', index=False)
writer.close()

wb = openpyxl.load_workbook('Результаты.xlsx')

ws = wb['4']
for column in range(ord('B'), ord('T')):
  column_letter = chr(column)
  ws.column_dimensions[column_letter].width = 20

for ws_name in ['5', '6', '7']:
  ws = wb[ws_name]
  if ws_name == '7':
    img_filename = f'graph_{ws_name}.png'
    img = Image(img_filename)
    ws.add_image(img, 'F3')
  for column in range(ord('A'), ord('E')):
    column_letter = chr(column)
    if ws_name == '6':
      ws.column_dimensions[column_letter].width = 30
    elif ws_name == '7':
      ws.column_dimensions[column_letter].width = 57
    else:
      ws.column_dimensions[column_letter].width = 20

wb.save('Результаты.xlsx')