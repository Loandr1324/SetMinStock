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


FOLDER = 'Исходные данные'
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

def create_df (file_list, add_name):
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
        # print (df.iloc[:15, :2])
        df = df.iloc[int(f):, :] # Убираем все строки с верха DF до заголовков
        df = df.dropna(axis=1, how='all')  # Убираем пустые колонки
        df.iloc[0, :] = df.iloc[0, :] + ' ' + add_name # Добавляем в наименование тип данных
        df.iloc[0, 0] = 'Код'
        df.iloc[0, 1] = 'Номенклатура'
        df.columns = df.iloc[0] # Значения из найденной строки переносим в заголовки DataFrame для простоты дальнейшего обращения
        df = df.iloc[2:, :] # Убираем две строки с верха DF
        df['Номенклатура'] = df['Номенклатура'].str.strip() # Удалить пробелы с обоих концов строки в ячейке
        df.set_index(['Код', 'Номенклатура'], inplace=True) # переносим колонки в индекс, для упрощения дальнейшей работы
        # print(df.iloc[:15, :2]) # Для тестов выводим в консоль 15 строк и два столбца полученного DF
        # Добавляем преобразованный DF в результирующий DF
        df_result = concat_df(df_result, df)
    # Добавляем в результирующий DF по продажам расчётные данные
    if add_name == 'продажи':
        df_result = payment(df_result)
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
    shutil.make_archive(f'{FOLDER}/correct_file', 'zip', tmp_folder)
    os.rename(f'{FOLDER}/correct_file.zip', file_name)

def concat_df(df1, df2):
    df = pd.concat([df1, df2], axis=1, ignore_index=False, levels=['Код'])
    return df

def sort_df (df):
    sort_list = ['01 Кирова продажи', '01 Кирова МО', '01 Кирова МО расчёт',
                 '02 Автолюбитель продажи', '02 Автолюбитель МО', '02 Автолюб. МО расчёт',
                 '03 Интер продажи', '03 Интер МО',	'03 Интер МО расчёт',
                 '04 Победа продажи', '04 Победа МО', '04 Победа МО расчёт',
                 '08 Центр продажи', '08 Центр МО',	'08 Центр МО расчёт',
                 '09 Вокзалка продажи', '09 Вокзалка МО', '09 Вокзалка МО расчёт',
                 '05 Павловский продажи', '05 Павловский МО', '05 Павловский МО расчёт',
                 'Компания MaCar продажи', 'Компания MaCar МО', 'Компания MaCar МО расчёт',
                 'Компания MaCar МО техническое']
    df = df[sort_list]
    return df

def payment (df_payment):
    df_payment = df_payment.astype('Int32')
    df_payment['Компания MaCar продажи'] = df_payment.sum(axis=1)

    list_colums_payment = ['01 Кирова МО расчёт', '02 Автолюб. МО расчёт', '03 Интер МО расчёт',
                           '04 Победа МО расчёт', '08 Центр МО расчёт', '09 Вокзалка МО расчёт',
                           '05 Павловский МО расчёт']

    list_colums_sales = ['01 Кирова продажи', '02 Автолюбитель продажи', '03 Интер продажи',
                         '04 Победа продажи', '08 Центр продажи', '09 Вокзалка продажи',
                         '05 Павловский продажи']

    for i in range(len(list_colums_payment)):
        df_payment[list_colums_payment[i]] = (df_payment[list_colums_sales[i]] / 4).apply(np.ceil)

    df_payment['Компания MaCar МО расчёт'] = df_payment[list_colums_payment].sum(axis=1)

    return df_payment

def df_write_xlsx(df):
    # Сохраняем в переменные значения конечных строк и столбцов
    row_end, col_end = len(df), len(df.columns)
    row_end_str, col_end_str = str(row_end), str(col_end)

    # Сбрасываем встроенный формат заголовков pandas
    pd.io.formats.excel.ExcelFormatter.header_style = None

    # Создаём эксельи сохраняем данные
    name_file = 'Анализ МО по компании.xlsx'
    sheet_name = 'Данные'  # Наименование вкладки для сводной таблицы
    writer = pd.ExcelWriter(name_file, engine='xlsxwriter')
    workbook = writer.book
    df.to_excel(writer, sheet_name=sheet_name)
    wks1 = writer.sheets[sheet_name]  # Сохраняем в переменную вкладку для форматирования

    # Получаем словари форматов для эксель
    header_format, con_format, border_storage_format_left, border_storage_format_right, \
    name_format, MO_format, data_format = format_custom(workbook)

    # Форматируем таблицу
    wks1.set_default_row(12)
    wks1.set_row(0, 40, header_format)
    wks1.set_column('A:A', 12, name_format)
    wks1.set_column('B:B', 32, name_format)
    wks1.set_column('C:AA', 10, data_format)

    # Делаем жирным рамку между складами и форматируем колонку с МО по всем складам
    i = 2
    while i < col_end+1:
        wks1.set_column(i, i, None, border_storage_format_left)
        wks1.set_column(i+1, i+1, None, MO_format)
        wks1.set_column(i+2, i+2, None, border_storage_format_right)
        i += 3

    # Подставляем формулу в колонку с МО по всей компании
    f = 2
    while f-1 <= row_end:
        wks1.write_formula(f'Y{f}', f'=IF(OR(AA{f}>1,AA{f}=0),SUM(D{f},G{f},J{f},M{f},P{f},S{f},V{f}),AA{f})')
        f += 1

    # Добавляем выделение цветом строки при МО=0 по всей компании
    wks1.conditional_format(f'A2:Z{row_end_str}', {'type': 'formula',
                                                   'criteria': '=AND($Y2=0,$X2<>0)',
                                                   'format': con_format})


    # Добавляем фильтр в первую колонку
    wks1.autofilter(0, 0, row_end+1, col_end)
    wks1.set_column(col_end+1, col_end+1, None, None, {'hidden': 1})
    writer.save() # Сохраняем файл
    return

def format_custom(workbook):
    header_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': '7',
        'align': 'center',
        'valign': 'top',
        'text_wrap': True,
        'bold': True,
        'bg_color': '#F4ECC5',
        'border': True,
        'border_color': '#CCC085'
    })

    border_storage_format_left = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'left': 2,
        'left_color': '#000000',
        'bottom': True,
        'bottom_color': '#CCC085',
        'top': True,
        'top_color': '#CCC085',
        'right': True,
        'right_color': '#CCC085',
    })
    border_storage_format_right = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'right': 2,
        'right_color': '#000000',
        'bottom': True,
        'bottom_color': '#CCC085',
        'top': True,
        'top_color': '#CCC085',
        'left': True,
        'left_color': '#CCC085',
    })

    name_format = workbook.add_format({
        'font_name': 'Arial',
        'font_size': '8',
        'align': 'left',
        'valign': 'top',
        'text_wrap': True,
        'bold': False,
        'border': True,
        'border_color': '#CCC085'
    })

    MO_format = workbook.add_format({
        'num_format': '# ### ##0.00;;',
        'bold': True,
        'font_name': 'Arial',
        'font_size': '8',
        'font_color': '#FF0000',
        # 'text_wrap': True,
        'border': True,
        'border_color': '#CCC085'
    })
    data_format = workbook.add_format({
        'num_format': '# ### ##0.00',
        'font_name': 'Arial',
        'font_size': '8',
        'text_wrap': True,
        'border': True,
        'border_color': '#CCC085'
    })
    con_format = workbook.add_format({
        'bg_color': '#FED69C',
    })

    return header_format, con_format, border_storage_format_left, border_storage_format_right, \
           name_format, MO_format, data_format

if __name__ == '__main__':
    salesFilelist = search_file(salesName) # запускаем функцию по поиску файлов и получаем список файлов
    minStockFilelist = search_file(minStockName) # запускаем функцию по поиску файлов и получаем список файлов
    df_sales = create_df (salesFilelist, salesName)
    df_minStock = create_df (minStockFilelist, minStockName)
    df_general = concat_df (df_sales, df_minStock)
    df_general['Компания MaCar МО техническое'] = df_general['Компания MaCar МО']
    df_general = sort_df(df_general) # сортируем столбцы
    # df_general.to_excel ('test.xlsx')  # записываем полученные джанные в эксель.
    df_write_xlsx(df_general)
