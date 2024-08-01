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
# print(city_name_lut)

name_wnum_lut = openpyxl.load_workbook(path[:-10] + "气象站编号对照表.xlsx") 
name_wnum_lut = name_wnum_lut.active
name_city = []
province_name = []
weather_number = []
city_number = []
for row in name_wnum_lut.iter_rows(min_row=2, min_col=1, max_row=name_wnum_lut.max_row, max_col=5, values_only=True):
    try:
        with open(path[:-10] + "2023ori_data\\" + str(row[0]) + "099999.csv", 'r') as file:
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

esc30 = [] # 最高气温86℉以上的累计天数
esc35 = [] # 最高气温95℉以上的累计天数
esc40 = [] # 最高气温104℉以上的累计天数
bel10 = [] # 最低气温14华氏度以下的累计天数
bel20 = [] # 最低气温-4华氏度以下的累计天数
amtd = [] # 年均昼夜温差 ℃
amws = [] # 年均风速 m/s
atp = [] # 年总降水量 mm
mem = [] # 平均温度最适区间[50℉,68℉]隶属度和 单日隶属度=1,50<x<68;-0.003324x^2+0.3922x-10.3711,other 求年和 x为日平均气温
for i in range(len(weather_number)):
    esc30.append(0)
    esc35.append(0)
    esc40.append(0)
    bel10.append(0)
    bel20.append(0)
    amtd.append(0)
    amws.append(0)
    atp.append(0)
    mem.append(0)
    for year in range(2019, 2024):
        with open(path[:-10] + str(year) + "ori_data\\" + str(weather_number[i]) + "099999.csv", 'r') as file:
            ori_data = csv.reader(file)
            next(ori_data)
            for day in ori_data:
                if float(day[20]) > 86:
                    esc30[i] += 1
                if float(day[20]) > 95:
                    esc35[i] += 1
                if float(day[20]) > 104:
                    esc40[i] += 1
                if float(day[22]) < 14:
                    bel10[i] += 1
                if float(day[22]) < -4:
                    bel20[i] += 1
                amtd[i] += float(day[20]) - float(day[22])
                amws[i] += 0.514 * float(day[16])
                atp[i] += (25.4 * float(day[24])) if day[24] != '99.99' else 0
                if float(day[6]) < 68 and float(day[6]) > 50:
                    mem[i] += 1
                else:
                    mem[i] += max(-0.003324 * (float(day[6])**2) + 0.3922 * float(day[6]) - 10.3711, 0)
    amtd[i] = (amtd[i] - 32) / 1.8 / 365 / 5
    amws[i] /= 365 * 5
    mem[i] /= 5
    esc30[i] /= 5
    esc35[i] /= 5
    esc40[i] /= 5
    bel10[i] /= 5
    bel20[i] /= 5
    atp[i] /= 5
            
res = xlwt.Workbook(encoding = 'utf-8')
sheet = res.add_sheet('sheet1')
title = ['城市', '省', '行政区代码', '气象站代码', '超30', '超35', '超40', '低10', '低20', '均温差', '风速', '降水量', '舒适天数']
for i in range(len(title)):
    sheet.write(0, i, label = title[i])
for i in range(len(name_city)):
    sheet.write(i+1, 0, label = name_city[i])
    sheet.write(i+1, 1, label = province_name[i])
    sheet.write(i+1, 2, label = city_number[i])
    sheet.write(i+1, 3, label = weather_number[i])
    sheet.write(i+1, 4, label = esc30[i])
    sheet.write(i+1, 5, label = esc35[i])
    sheet.write(i+1, 6, label = esc40[i])
    sheet.write(i+1, 7, label = bel10[i])
    sheet.write(i+1, 8, label = bel20[i])
    sheet.write(i+1, 9, label = amtd[i])
    sheet.write(i+1, 10, label = amws[i])
    sheet.write(i+1, 11, label = atp[i])
    sheet.write(i+1, 12, label = mem[i])
res.save('气象数据.xls')
# print(name_city)
# print(esc30)