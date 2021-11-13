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

def search_file(name):
    """
    :param name: Поиск всех файлов в папке FOLDER, в наименовании которых, содержится name
    :return: filelist список с наименованиями фалов
    """
    filelist = []
    for item in os.listdir(FOLDER):
        if name in item and item.endswith('.xlsx'): # если файл содержит name и с расширенитем .xlsx, то выполняем
            filelist.append(FOLDER + "/" + item) # добаляем в список папку и имя файла для последующего обращения из списка
        else:
            pass
    return filelist


def read_xlsx(file_list, add_name):
    """
    :param file_list: Загружаем в DataFrame файлы из file_list
    :param add_name: Добавляем add_name в наименование колонок DataFrame
    :return: df_result Дата фрэйм с данными из файлов
    """

    df_result = None # объявляем перменную

    for filename in file_list: # проходим по каждому элементу списка файлов
        print(filename) # для тестов выводим в консоль наименование файла с которым проходит работа
        df = pd.read_excel(filename, sheet_name='TDSheet', header=None, usecols="A:I",
                           skipfooter=1, engine='openpyxl')
        df_search_header = df.iloc[:15, :2] # для ускорения работы выбираем из DataFrame первую колонку и 15 строк
        # print (df_search_header)
        # создаём маску и отмечаем True строку где есть слово "Номенклатура", остальные False
        mask = df_search_header.replace('.*Номенклатура.*', True, regex=True).eq(True)
        # Преобразуем Dataframe согласно маски. После обработки все значения будут NaN кроме нужного нам.
        # В этой же строке кода удаляем все строки со значением NaN и далее получаем индекс оставшейся строки
        f = df_search_header[mask].dropna(axis=0, how='all').index.values # Удаление пустых колонок, если axis=0, то строк
        print ( f )
        # print (df.iloc[:15, :2])
        df = df.iloc[int(f):, :] # Убираем все строки с верха DF до заголовков
        df = df.dropna(axis=1, how='all')  # Убираем пустые колонки
        df.iloc[0, :] = df.iloc[0, :] + ' ' + add_name # Добавляем в наименование тип данных
        df.iloc[0, 0] = 'Код'
        df.iloc[0, 1] = 'Номенклатура'
        df.columns = df.iloc[0] # Значения из найденной строки переносим в заголовки DataFrame для простоты дальнейшего обращения
        df = df.iloc[2:, :] # Убираем две строки с верха DF
        df.set_index(['Код', 'Номенклатура'], inplace=True) # переносим колонки в индекс, для упрощения дальнейшей работы
        print(df.iloc[:15, :2]) # Для тестов выводим в консоль 15 строк и два столбца полученного DF
        # Добавляем преобразованный DF в результирующий DF
        #df_result = pd.concat([df_result, df], axis=1, ignore_index=False, levels=['Код'])
        df_result = concat_df(df_result, df)


    # Добавляем в результирующий DF колоку и подставляем в неё сумму всех колонок
    if add_name == 'продажи':
        df_result['Компания MaCar продажи'] = df_result.sum(axis=1)

    return df_result

def concat_df(df1, df2):
    df = pd.concat([df1, df2], axis=1, ignore_index=False, levels=['Код'])
    return df

def sort_df (df):
    sort_list = ['01 Кирова продажи', '01 Кирова МО', '02 Автолюбитель продажи', '02 Автолюбитель МО',
                 '03 Интер продажи', '03 Интер МО',	'04 Победа продажи', '04 Победа МО',
                 '08 Центр продажи', '08 Центр МО',	'09 Вокзалка продажи', '09 Вокзалка МО',
                 '05 Павловский продажи',	'05 Павловский МО',	'Компания MaCar продажи', 'Компания MaCar МО']
    df = df[sort_list]
    return df

if __name__ == '__main__':
    salesFilelist = search_file(salesName) # запускаем функцию по поиску файлов и получаем список файлов
    minStockFilelist = search_file(minStockName) # запускаем функцию по поиску файлов и получаем список файлов
    df_sales = read_xlsx (salesFilelist, salesName)
    df_minStock = read_xlsx (minStockFilelist, minStockName)
    df_general = concat_df (df_sales, df_minStock)
    df_general = sort_df(df_general) # сортируем столбцы
    df_general.to_excel ('test.xlsx')  # записываем полученные джанные в эксель.
