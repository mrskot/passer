#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import openpyxl
import MySQLdb


#Constante
DIR = 'SPEC'

#функция создания списка файлов переданной в функцию дирректории
def read_dir(dir):
    #создание списка дирректории
    file_list = os.listdir(dir)
    #возврат списка дирректории
    return file_list

#функция открывает файл и возвращает номенклатурный список
def read_file(file_name):
    file_name = file_name
    spec_list = []

    wb = openpyxl.load_workbook(DIR + '/' + file_name)
    sheet = wb.get_active_sheet()

    # цикл создаёт лист спецификации из excel файла (столбец 'B' ислючения пустые ячейки и ясейки со словом 'Номенклатура')
    for spec in sheet.columns[2]:
        if spec.value is not None and spec.value != 'Номенклатура':
            spec_list.append(spec.value)
    return spec_list

# функция получает спецификацию и имя файла фильтрует и возвращает список столбцов
def filter_spec(spec_list, file_name):
    spec_list = spec_list
    file_name = file_name
    filter_name = [file_name, '', '', '', '', '']

    for spec in spec_list:
        if spec[0:3] in '02.0':
            korpus_number = spec[0:8]
            korpus_name = filter_korpusa(korpus_number)

            filter_name[1] = korpus_name
            #break
#        elif spec[0:8] in 'Шкаф ШПТ':
#            korpus_name = spec[5:12]
#            filter_name[1] = korpus_name

#        elif spec[0:3] in '04.0':
#            detal_1_number = spec[0:12]
#            detal_1_name = filter_detal_1(detal_1_number)
#
#            filter_name[2] = detal_1_name
            #break

    return filter_name

#функция получает номер детали корпуса из filter_spec и определяет название корпуса и передаёт
def filter_korpusa(korpus_number):
    korpus_number = korpus_number
    korpus_tip = ['Д', 'К', 'M', 'Щ', 'А', 'П']
    korpus_num = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]

    tip = int(korpus_number[4])
    num = int(korpus_number[7])
    #korpus_name = korpus_number
    korpus_name = str(korpus_tip[tip - 1]) + '.' + str(korpus_num[num - 1])

    return korpus_name

def filter_detal_1(detal_1_number):
    detal_1_number = detal_1_number
    detal_1_name = 'detal_1_number' + detal_1_number
    return detal_1_name



# функция получает список столбцов и добавляет их в базу данных
def add_spec(filter_list):
    add_list = filter_list
    db = MySQLdb.connect("localhost", "mrskot", "mrskot", "for_pyqt", charset='utf8', init_command='SET NAMES UTF8')
    cursor = db.cursor()
    cursor.execute("SELECT * FROM test")
    sql = "INSERT INTO test(PROJECT, TIP_KORPUSA, USTANOVKA) VALUES ('%s','%s','%s')" % (add_list[0], add_list[1], add_list[2])
    try:
        cursor.execute(sql)
        db.commit()
    except:
        db.rollback()
    return add_list

# главная функция
def my_main():
    # принять список файлов из указанной в параметрах функции read_dir дирректории
    file_list = read_dir(DIR)
    print(file_list)
    print(len(file_list))

    # цикл поочерёдно предаёт в функцию read_file имена файлов и получает список спецификации
    for file_name in file_list:
        # получение спицификации по имени файла
        spek_list = read_file(file_name)
        #print(spek_list)

        #получение отфильтрованной спецификации по переданной спецификации и имени файла
        filter_name = filter_spec(spek_list, file_name)
        print(filter_name)

        #добавление в базу данных созданной отфильтрованной спецификации
        add_spec(filter_name)



my_main()
































