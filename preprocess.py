import openpyxl
import xlwt
import os
import csv
import json

def city_name_equa(a:str, b:str, f = 1):
    flag = 1
    if len(a) > 2:
        a = a[:-1]
    if len(b) > 2:
        b = b[:-1]
    for i in a:
        if i not in b:
            flag = 0
            break
    if flag or (f and city_name_equa(b, a, 0)):
        return 1 # 双向属于满足其一即匹配成功
    else:
        return 0 # 匹配失败

path = os.path.split(os.path.realpath(__file__))[0]
city_name_lut = {}
with open(path[:-10] + "ChinaCityList.json", encoding='utf-8') as loadf:
    ori_city_num_lut = json.load(loadf)
    for province in ori_city_num_lut:
        for city in province['cities']:
            city_name_lut[city['code']] = city['name']
print(city_name_lut)

name_wnum_lut = openpyxl.load_workbook(path[:-10] + "气象站编号对照表.xlsx") 
name_wnum_lut = name_wnum_lut.active
name_city = []
province_name = []
weather_number = []
city_number = []
for row in name_wnum_lut.iter_rows(min_row=2, min_col=1, max_row=name_wnum_lut.max_row, max_col=5, values_only=True):
    try:
        with open(path[:-10] + "气象原始数据\\" + str(row[0]) + "099999.csv", 'r') as file:
            ori_data = csv.reader(file)
            pass
    except:
        continue
    if row[3] != None:
        this_name = row[3]
    elif row[4] != None:
        this_name = row[4]
    else:
        this_name = row[1]
    if this_name not in name_city:
        name_city.append(this_name)
        province_name.append(row[2])
        weather_number.append(row[0])
        for key in city_name_lut:
            if city_name_equa(city_name_lut[key], this_name):
                city_number.append(key)
                name_city[-1] = city_name_lut[key]
                break
        else:
            city_number.append('error')
# print(len(name_city))
# print(name_city)
# print(city_number)

esc30 = []
esc35 = []
esc40 = []
bel10 = []
bel20 = []
amtd = []
amws = []
atp = []
for city in weather_number:
    with open(path[:-10] + "气象原始数据\\" + str(city) + "099999.csv", 'r') as file:
            ori_data = csv.reader(file)
            # 计算各个变量