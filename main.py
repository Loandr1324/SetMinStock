# Author Loik Andrey 7034@balancedv.ru
# TODO
#  Часть 1
#  1. Создать отдельный файл программы для чтения файлов эксель. Передавать туда маску наименования файла, который требуется прочитать.
#  2. Прочитать все файлы по продажам и объединить
#  3. Прочитать все файлы МО и объединить.
#  4. Прочиттаь старый итоговый файл и взять от туда старые данные по продажам
#  5. Занести все данные в новый файл и найти отличия новых данных по продажам по всей компании со старыми и записать в отдельную колонку
#  6. Отформатировать все строки, выделить те строки в которых в колонке Отличие значение не равно нулю
#  Часть 2
#  7. Считать значения МО из эксель файл полученного после редактирования отвественным заказчиком итого вого файла из Части 1
#  8. Создать файл для загрузки значений МО, которые были считаны в п.7

import pandas as pd
import os

# import readfile

FOLDER = 'Установка МО ответственным'
salesName = 'продажи'
minStockName = 'МО'


# sales = readxls(FOLDER, 'Продажи')

def search_file(name):
    """
    :param name: Поиск всех файлов, в наименовании которых, содержится name
    :return: filelist список с наименованиями фалов
    """
    filelist = []
    for item in os.listdir(FOLDER):  # для каждого файла в папке FOLDER
        if name in item and item.endswith('.xlsx'):
            filelist.append(FOLDER + "/" + item)
        else:
            pass
    return filelist


def read_xlsx(file_list):
    """
    :param file_list: Загружаем в DataFrame файлы из file_list
    :return: df_result Дата фрэйм с данными из файлов
    """

    df_result = None
    for filename in file_list:
        print(filename)
        df = pd.read_excel(filename, sheet_name='TDSheet', header=0, usecols="A:I",
                           skipfooter=1, engine='openpyxl')
        df_search_header = df.iloc[:15, :1]
        mask = df_search_header.replace('.*Номенклатура.*', True, regex=True).eq(True)
        f = df_search_header[mask].dropna(axis=0,
                                          how='all').index.values  # Удаление пустых колонок, если axis=0, то строк
        df.columns = df.iloc[int(f)]
        df = df.iloc[int(f) + 2:, :]
        df = df.dropna(axis=1, how='all')  # Убрали пустые колонки
        df.set_index(['Номенклатура.Код', 'Номенклатура'], inplace=True)
        print(df.iloc[:15, :2])

        df_result = pd.concat([df_result, df], axis=1, ignore_index=False, levels=['Номенклатура.Код'])

    df_result['Итого по компании'] = df_result.sum(axis=1)
    df_result.to_excel('test.xlsx')
    return


if __name__ == '__main__':
    salesFilelist = search_file(salesName)
    minStockFilelist = search_file(minStockName)
    # print ( salesFilelist )
    # print ( minStockFilelist )
    read_xlsx(salesFilelist)

    '''
    Блок №1
    Чтение данных за прошедший месяц, сортировка и запись суммовых значений в итоговый файл с данными

    print("Начало Блока №1")
    # Поиск списков файлов для чтения и распределение по типам файлов
    file_custom_order, file_supp_order, file_supp_receipt, file_sms = search_file()
    # Определение старых и новых имён файлов с хранимыми данным и переименование
    file_custom, file_supp = rename_out_file()
    # Чтение данных за прошедший месяц по заказам клиента и запись в DataFrame + получения общего количетсва строк
    custom_row, sms_row, quantity_row_custom = read_xlsx_custom(file_custom_order, file_sms)
    # Сортировка DataFrame заказов клиента по типам продаж
    your_warehouse, another_warehouse, by_order = sorting_custom_row(custom_row)
    # Сортировка DataFrame СМС. Добавляем, если доставлено
    sms = sorting_sms(sms_row)
    # Чтение данных за прошедший месяц по заказам поставщикам и запись в DataFrame + получения общего количетсва строк
    supp_ord_row, supp_rec_row, quantity_row_supp = read_xlsx_supp(file_supp_order, file_supp_receipt)
    # Получение итоговых значений по типам продаж и объединение итоговых значений в один DataFrame
    total_custom = total_df_custom(another_warehouse, your_warehouse, by_order, sms)
    # Получение итоговых значений по типам дкументов и объединение итоговых значений в один DataFrame
    total_supp = total_df_supp(supp_ord_row, supp_rec_row)
    # Добавляем в файл данных по клиентам, данные по Заказам клиента
    append_file_data(file_custom, total_custom)
    # Добавляем в файл данных по поставщикам, данные по Заказам поставщикам
    append_file_data(file_supp, total_supp)
    print("Блок №1 завершён!")

    Блок №2
    Чтение данных из файлов данных, построение графиков и запись итоговых файлов для отпрвки по почте

    # Считываем данные за всё время и формируем итоговый файл по Заказам для наличия
    print("Начало Блока №2")
    # print(file_custom, file_supp) # Используем при тестах
    custom_category = 'Тип продажи'
    custom_set_cat = ['Продажи с других складов', 'Продажи со своего склада',
                      'Заказное', 'Отправка СМС', 'Итого за год']
    custom_index = ['Год', 'Тип продажи', 'Месяц']
    supp_category = 'Тип документа'
    supp_set_cat = ['Заказы внешним поставщикам', 'Поступления от МХ Комсомольск', 'Итого за год']
    supp_index = ['Год', 'Тип документа', 'Месяц']
    supp_index_out = ['Год', 'Месяц', 'Тип документа']
    caption_custom = 'Carbaz Заказы клиентов. Все заказы кроме наших заказов внешним поставщикам на ' \
                     'постоянное наличие. Включая индивидуальные заказы.'
    caption_supp = 'Наши заказы внешним поставщикам на постоянное наличие. (Заказы Клиентов не входят)'
    sec_level_custom = ['Год', 'Тип продажи']
    sec_level_supp = ['Год', 'Месяц']
    data_pt_custom = pivot_table(file_custom, custom_category, custom_set_cat, custom_index)
    data_pt_supp = pivot_table(file_supp, supp_category, supp_set_cat, supp_index)
    result_to_xlsx(out_file_custom, data_pt_custom, custom_category, caption_custom, sec_level_custom, custom_index)
    result_to_xlsx(out_file_supp, data_pt_supp, supp_category, caption_supp, sec_level_supp, supp_index_out)
    print("Блок №2 завершён!")

    Блок №3
    Отправка файлов по почте и выдержка по количеству строк в теле письма

    # Подготавливаем и отправляем данные по почте Юре
    print("Начало Блока №3")
    send_mail()
    print("Блок №3 завершён!")
    '''
