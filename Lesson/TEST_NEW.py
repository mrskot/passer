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
    filter_name = [file_name, None, None, None, None, None, None, None, None,
                   None, None, None, None, None, None, None, None, None, None, None]

    for spec in spec_list:
        if spec[0:3] in '02.0':
            korpus_number = spec[0:8]
            korpus_name = filter_korpusa(korpus_number)

            filter_name[1] = korpus_name

        elif spec[0:8] in 'Шкаф ШПТ':
            korpus_name = spec[5:12]
            filter_name[1] = korpus_name



        elif spec[0:6] in '04.001':
            filter_name[8] = spec

        elif spec[0:6] in '04.004':
            filter_name[9] = spec

        elif spec[0:6] in '04.005':
            filter_name[10] = spec

        elif spec[0:6] in '04.008':
            filter_name[11] = spec

        elif spec[0:6] in '04.009':
            filter_name[12] = spec

        elif spec[0:6] in '04.010':
            filter_name[13] = spec

        elif spec[0:6] in '04.011':
            filter_name[14] = spec

        elif spec[0:6] in '04.014':
            filter_name[15] = spec

        elif spec[0:6] in '04.015':
            filter_name[16] = spec

        elif spec[0:6] in '04.016':
            filter_name[17] = spec

        elif spec[0:6] in '04.018':
            filter_name[18] = spec

        elif spec[0:6] in '04.022':
            filter_name[19] = spec

    return filter_name

#функция получает номер детали корпуса из filter_spec и определяет название корпуса и передаёт
def filter_korpusa(korpus_number):
    korpus_number = korpus_number
    korpus_tip = ['Д', 'К', 'M', 'Щ', 'А', 'П']
    korpus_num = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]

    tip = int(korpus_number[4])
    num = int(korpus_number[7])

    korpus_name = str(korpus_tip[tip - 1]) + '.' + str(korpus_num[num - 1])

    return korpus_name

# функция получает список столбцов и добавляет их в базу данных
def add_spec(filter_list):
    add_list = filter_list
    db = MySQLdb.connect("localhost", "mrskot", "mrskot", "for_pyqt",
                         charset='utf8', init_command='SET NAMES UTF8')
    cursor = db.cursor()
    cursor.execute("SELECT * FROM test")
    sql = "INSERT INTO test(PROJECT, TIP_KORPUSA, OBOGREV, ADAPTOR_TRUBNIY, OSNOLAYN_KVS, KAPILYARI," \
          " PILT, TERMOSTAT, 04_001_ADAPTOR,04_004_PLITA, 04_005_PLANKA, 04_008_STOYKA, 04_009_USEL_NA_BOB," \
          " 04_010_USEL_NA_TRUB, 04_011_USEL_NA_FLAN, 04_014_KRONSHTEYN_NA_POV, 04_015_USTROYSTVO_KAPILAR," \
          " 04_016_STEKLOPAKET, 04_018_OPORA_MONTAJNAYA, 04_022_KRONSHTEYNI) " \
          "VALUES ('%s','%s','%s','%s','%s','%s', '%s','%s','%s','%s'," \
          "'%s','%s','%s','%s', '%s','%s','%s','%s','%s','%s')" \
          % (add_list[0], add_list[1], add_list[2], add_list[3], add_list[4], add_list[5], add_list[6],
             add_list[7], add_list[8], add_list[9], add_list[10], add_list[11], add_list[12], add_list[13],
             add_list[14], add_list[15], add_list[16], add_list[17], add_list[18], add_list[19])
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
































