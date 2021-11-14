# Author Loik Andrey 7034@balancedv.ru
# TODO
#  Часть 1
#  1. Создать отдельный файл программы для чтения файлов эксель. Передавать туда маску наименования файла, который требуется прочитать.
#  Готово 2. Прочитать все файлы по продажам и объединить
#  Готово 3. Прочитать все файлы МО и объединить.
#  Не делаем 4. Прочитать старый итоговый файл и взять от туда старые данные по продажам
#  5. Занести все данные в новый файл и
#  6. найти отличия новых данных по продажам по всей компании со старыми и записать в отдельную колонку
#  7. Отформатировать все строки, выделить те строки в которых в колонке Отличие значение не равно нулю
#  Часть 2
#  8. Считать значения МО из эксель файл полученного после редактирования отвественным заказчиком итого вого файла из Части 1
#  9. Создать файл для загрузки значений МО, которые были считаны в п.7

import pandas as pd
import numpy as np
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

    df_result = None

    for filename in file_list: # проходим по каждому элементу списка файлов
        print (filename) # для тестов выводим в консоль наименование файла с которым проходит работа
        df = read_excel(filename)
        df_search_header = df.iloc[:15, :2] # для ускорения работы выбираем из DataFrame первую колонку и 15 строк
        # print (df_search_header)
        # создаём маску и отмечаем True строку где есть слово "Номенклатура", остальные False
        mask = df_search_header.replace('.*Номенклатура.*', True, regex=True).eq(True)
        # Преобразуем Dataframe согласно маски. После обработки все значения будут NaN кроме нужного нам.
        # В этой же строке кода удаляем все строки со значением NaN и далее получаем индекс оставшейся строки
        f = df_search_header[mask].dropna(axis=0, how='all').index.values # Удаление пустых колонок, если axis=0, то строк
        print ( 'Номер строки с заколовком:' + f )
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
        df_result = concat_df(df_result, df)
    # Добавляем в результирующий DF по продажам колоку и подставляем в неё сумму всех колонок
    if add_name == 'продажи':
        df_result = df_result.astype('Int32')
        df_result = df_result / 4
        df_result = df_result.apply(np.ceil)
        df_result['Компания MaCar продажи'] = df_result.sum(axis=1)

    return df_result

def read_excel (file_name):
    """
    Пытаемся прочитать файл xlxs, если не получается, то исправляем ошибку и опять читаем файл
    :param file_name: Имя файла для чтения
    :return: DataFrame
    """
    print ('Попытка загрузки файла:'+file_name)
    try:
        df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=1, engine='openpyxl')
        return (df)
    except KeyError as Error:
        print (Error)
        df = None
        if str(Error) == "\"There is no item named 'xl/sharedStrings.xml' in the archive\"":
            bug_fix (file_name)
            print('Исправлена ошибка: ', Error, f'в файле: \"{file_name}\"\n')
            df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=1, engine='openpyxl')
            return df
        else:
            print('Ошибка: >>' + str(Error) + '<<')

def bug_fix (file_name):
    """
    Переименовываем не корректное имя файла в архиве excel
    :param file_name: Имя excel файла
    """
    import shutil
    from zipfile import ZipFile

    # Создаем временную папку
    tmp_folder = '/temp/'
    os.makedirs(tmp_folder, exist_ok=True)

    # Распаковываем excel как zip в нашу временную папку и удаляем excel
    with ZipFile(file_name) as excel_container:
        excel_container.extractall(tmp_folder)
    os.remove(file_name)

    # Переименовываем файл с неверным названием
    wrong_file_path = os.path.join(tmp_folder, 'xl', 'SharedStrings.xml')
    correct_file_path = os.path.join(tmp_folder, 'xl', 'sharedStrings.xml')
    os.rename(wrong_file_path, correct_file_path)

    # Запаковываем excel обратно в zip и переименовываем в исходный файл
    shutil.make_archive('Установка МО ответственным/correct_file', 'zip', tmp_folder)
    os.rename('Установка МО ответственным/correct_file.zip', file_name)

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
    df_general = df_general.reset_index()
    df_general.to_excel ('test.xlsx', index=False)  # записываем полученные джанные в эксель.
