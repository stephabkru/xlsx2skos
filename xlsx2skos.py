from openpyxl import *
import pandas as pd
import csv
from io import StringIO
from rdflib import Graph, RDF, Literal, URIRef
from rdflib.namespace import SKOS
from collections import OrderedDict

#функция для создания массива данных из файла excel.
def begin_data(df):
    arr = []
    for i in range (0,df.shape[0]):
        item = ''
        for j in range (0,df.shape[1]):
            item = item + ',' + str(df[j][i])
        arr.append(item)
    arr_fin = []
    for i in range(0,len(arr)):
        y = arr[i].replace("None", "")
        y = y[1:]
        arr_fin.append(y)
    return(arr_fin)

#функция нахождения главного языка
def findmainlang(arr):
    a = arr[0].split(':')
    b = a[1].split(',')
    return b[0]

#функция для создания массива упорядоченных библиотек, описывающих строки онтологии.
#На вход подается преобразованная в нужный формат (где разделитель - запятая) онтология
def create_rows(arr):
    arr_midterm = []
    for j in range(1,len(arr)):
        a = arr[j].split(',')
        for i in range (0,len(a)):
            if a[i] == 'None':
                a[i] = ''
        arr_midterm.append(a)
    rows = []
    for e in range(1,len(arr_midterm)):
        d = OrderedDict()
        for u in range (0,len(arr_midterm[0])):
            d.update({arr_midterm[0][u]:arr_midterm[e][u]})
        rows.append(d)
    return rows

#Прописываются имена файлов
name_of_your_file = 'finik.xlsx'
name_of_need_sheet = 'finik'
name_of_output_file = 'finalfo.rdf'

#Записываем данные excel в датафрейм
wb = load_workbook(name_of_your_file)
sheet = wb[name_of_need_sheet]
df = pd.DataFrame(sheet.values)

#Создаем объект для записи онтологии в формат SKOS, а также выгружаем данные из DataFrame в нужном формате
g = Graph()
g.bind('skos', SKOS)
arr_fin = begin_data(df)
mainlang = findmainlang(arr_fin)
data_rows = create_rows(arr_fin)

#Конвертируем данные
lang_dict = {mainlang: []}
level = 1
previous_level = 1
last_level = {}

for row in data_rows:
    uri = row['URI']
    for column in row:
        if 'level' in column and row[column]:
            level = int(column.split(':')[1])
            last_level[level] = uri
    column_level = ':'.join(['level', str(level)])
    main_item = row[column_level].strip()

    lang_dict[mainlang].append((uri, main_item))
    g.add((URIRef(uri), RDF['type'], SKOS['Concept']))

    preflabels = []
    literal = Literal(main_item.strip(), lang=mainlang)
    preflabels.append(literal)
    for element in preflabels:
        g.add((URIRef(uri), SKOS['prefLabel'], element))
            
    if not level == 1:
        if not level == previous_level:
            parent = last_level[level - 1]
        g.add((URIRef(uri), SKOS['broader'], URIRef(parent)))
        g.add((URIRef(parent), SKOS['narrower'], URIRef(uri)))
        
    previous_level = level
x = str(str(g.serialize(format='pretty-xml'), "utf-8").encode('utf-8'))

#Преобразовываем конвертированную онтологию под формат записи в файл rdf и записываем
y = x.replace("\\n", "\n")
final_file = y[2:-2]
file = open(name_of_output_file, "w")
file.write(final_file)
file.close()
