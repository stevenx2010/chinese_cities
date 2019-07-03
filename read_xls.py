import xlrd
import json
import sys

reload(sys)
sys.setdefaultencoding('utf-8')

workbook = xlrd.open_workbook('baidu_data_201904_modified.xlsx')
table = workbook.sheets()[0]
rows = table.nrows
print(rows)

def generate_data(text, value, child, unit_type):
    if child == 0:
        data = {'text': text, 'value': int(value)}
    else:
        if unit_type == 'c':    #city
            parent = int(value) / 10000 * 10000
        if unit_type == 'd':    #district
            parent = int(value) /100 * 100
        if unit_type == 't':    #town
            parent = int(value) / 1000

        data = {'text': text, 'value': int(value), 'parentVal': parent}

    #print(data)
    return data



province_list = [generate_data(table.cell(0,2).value, table.cell(0,1).value, 0, 'p')]
city_list = [generate_data(table.cell(0,4).value, table.cell(0,3).value, 1, 'c')]
district_list = [generate_data(table.cell(0,6).value, table.cell(0,5).value, 1, 'd')]
town_list = [generate_data(table.cell(0,8).value, table.cell(0,7).value, 1, 't')]

i = 1
while i < 1000:
    #build province list
    data = generate_data(table.cell(i,2).value, table.cell(i,1).value, 0, 'p')
    try:       
        province_list.index(data)
    except ValueError:
        province_list.append(data)

    #build city list
    data = generate_data(table.cell(i,4).value, table.cell(i,3).value, 1, 'c')
    try:
        city_list.index(data)
    except ValueError:
        city_list.append(data)

    #build district list
    data = generate_data(table.cell(i,6).value, table.cell(i,5).value, 1, 'd')
    try:
        district_list.index(data)
    except ValueError:
        district_list.append(data)

    #build town list
    data = generate_data(table.cell(i,8).value, table.cell(i,7).value, 1, 't')
    try:
        town_list.index(data)
    except ValueError:
        town_list.append(data)

    i = i + 1

def sortByValue(element):
    return element['value']

province_list.sort(key=sortByValue)
city_list.sort(key=sortByValue)
district_list.sort(key=sortByValue)

cityColumns = [
    {
        'options': province_list
    },
    {
        'options': city_list
    },
    {
        'options': district_list
    }
]

#print(json.dumps(cityColumns).decode('unicode_escape'))

line1 = 'export class ChineseCities {\n'
line2 = '\tstatic cities = \n'
line3 = '\t\t'
last_line = '\n}'

f = open('chinese-cities.ts', 'wb')
f.write(line1 + line2)
f.write(line3)
f.write(json.dumps(cityColumns).decode('unicode_escape'))
f.write(last_line)
f.close()