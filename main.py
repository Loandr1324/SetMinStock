# Author Loik Andrey 7034@balancedv.ru
# TODO
#  Часть 1
#  1. Создать отдельный файл программы для чтения файлов эксель. Передавать туда маску наименования файла,
#  который требуется прочитать.
#  Готово 2. Прочитать все файлы по продажам и объединить
#  Готово 3. Прочитать все файлы МО и объединить.
#  Не делаем 4. Прочитать старый итоговый файл и взять от туда старые данные по продажам
#  5. Занести все данные в новый файл и
#  6. найти отличия новых данных по продажам по всей компании со старыми и записать в отдельную колонку
#  7. Отформатировать все строки, выделить те строки в которых в колонке Отличие значение не равно нулю
#  Часть 2.
#  8. Считать значения МО из эксель файл полученного после редактирования
#  отвественным заказчиком итого вого файла из Части 1
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
        if name in item and item.endswith('.xlsx'):  # если файл содержит name и с расширением .xlsx, то выполняем
            # Добавляем в список папку и имя файла для последующего обращения из списка
            filelist.append(FOLDER + "/" + item)
        else:
            pass
    return filelist


def create_df(file_list, add_name):
    """
    :param file_list: Загружаем в DataFrame файлы из file_list
    :param add_name: Добавляем add_name в наименование колонок DataFrame
    :return: df_result Дата фрэйм с данными из файлов
    """

    df_result = None

    for filename in file_list:  # проходим по каждому элементу списка файлов
        print(filename)  # для тестов выводим в консоль наименование файла с которым проходит работа
        df = read_excel(filename)
        df_search_header = df.iloc[:15, :2]  # для ускорения работы выбираем из DataFrame первую колонку и 15 строк
        # print(df_search_header)
        # создаём маску и отмечаем True строку где есть слово "Номенклатура", остальные False
        mask = df_search_header.replace('.*Номенклатура.*', True, regex=True).eq(True)
        # Преобразуем Dataframe согласно маске. После обработки все значения будут NaN кроме нужного нам.
        # В этой же строке кода удаляем все строки со значением NaN и далее получаем индекс оставшейся строки
        f = df_search_header[mask].dropna(axis=0, how='all').index.values # Удаление пустых строк
        df = df.iloc[int(f):, :]  # Убираем все строки с верха DF до заголовков
        df = df.dropna(axis=1, how='all')  # Убираем пустые колонки
        df.iloc[0, :] = df.iloc[0, :] + ' ' + add_name  # Добавляем в наименование тип данных
        df.iloc[0, 0] = 'Код'
        df.iloc[0, 1] = 'Номенклатура'
        # Значения из найденной строки переносим в заголовки DataFrame для простоты дальнейшего обращения
        df.columns = df.iloc[0].values
        df = df.iloc[2:, :]  # Убираем две строки с верха DF
        df['Номенклатура'] = df['Номенклатура'].str.strip()  # Удалить пробелы с обоих концов строки в ячейке
        # Переносим колонки в индекс, для упрощения дальнейшей работы
        df.set_index(['Код', 'Номенклатура'], inplace=True)

        # Добавляем преобразованный DF в результирующий DF
        if df_result is None:
            df_result = df
        else:
            df_result = concat_df(df_result, df)
    # Добавляем в результирующий DF по продажам расчётные данные
    if add_name == 'продажи':
        df_result = payment(df_result)
    return df_result


def read_excel(file_name):
    """
    Пытаемся прочитать файл xlxs, если не получается, то исправляем ошибку и опять читаем файл
    :param file_name: Имя файла для чтения
    :return: DataFrame
    """
    print ('Попытка загрузки файла:'+file_name)
    try:
        if 'продажи' in file_name:
            df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=1, engine='openpyxl')
        else:
            df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=0, engine='openpyxl')
        return (df)
    except KeyError as Error:
        print (Error)
        df = None
        if str(Error) == "\"There is no item named 'xl/sharedStrings.xml' in the archive\"":
            bug_fix (file_name)
            print('Исправлена ошибка: ', Error, f'в файле: \"{file_name}\"\n')
            if 'продажи' in file_name:
                df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=1, engine='openpyxl')
            else:
                df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=0, engine='openpyxl')
            return df
        else:
            print('Ошибка: >>' + str(Error) + '<<')


def bug_fix(file_name):
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
    # df = pd.concat([df1, df2], axis=1, ignore_index=False, levels=['Код'])
    df = pd.concat([df1, df2], axis=1, ignore_index=False)
    return df


def sort_df(df):
    sort_list = ['01 Кирова продажи', '01 Кирова МО', '01 Кирова МО расчёт',
                 '02 Автолюбитель продажи', '02 Автолюбитель МО', '02 Автолюбитель МО расчёт',
                 '03 Интер продажи', '03 Интер МО',	'03 Интер МО расчёт',
                 '04 Победа продажи', '04 Победа МО', '04 Победа МО расчёт',
                 '08 Центр продажи', '08 Центр МО',	'08 Центр МО расчёт',
                 '09 Вокзалка продажи', '09 Вокзалка МО', '09 Вокзалка МО расчёт',
                 '05 Павловский МО',  '05 Павловский МО расчёт',
                 'Компания MaCar продажи', 'Компания MaCar МО', 'Компания MaCar МО расчёт',
                 'Компания MaCar МО техническое']
    df = df[sort_list]
    return df


def payment(df_payment):
    df_payment = df_payment.astype('Int32')
    df_payment['Компания MaCar продажи'] = df_payment.sum(axis=1)

    list_colums_payment = ['01 Кирова МО расчёт', '02 Автолюбитель МО расчёт', '03 Интер МО расчёт',
                           '04 Победа МО расчёт', '08 Центр МО расчёт', '09 Вокзалка МО расчёт']

    list_colums_sales = ['01 Кирова продажи', '02 Автолюбитель продажи', '03 Интер продажи',
                         '04 Победа продажи', '08 Центр продажи', '09 Вокзалка продажи']

    for i in range(len(list_colums_payment)):
        df_payment[list_colums_payment[i]] = df_payment[list_colums_sales[i]].fillna(0).apply(calc)
    return df_payment


def calc(val, x=int(input('Введите кол-во дней анализа продаж: ')), y=int(input('Введите кол-во дней запаса: '))):
# def calc(val, x=365, y=90):
    quan_prod = val / x * y
    if 0 < quan_prod < 1:
        val = 1
    else:
        val = int(quan_prod)
    return val


def df_write_xlsx(df):
    # Сохраняем в переменные значения конечных строк и столбцов
    row_end, col_end = len(df), len(df.columns)
    row_end_str, col_end_str = str(row_end), str(col_end)

    # Для простоты форматирования переводим индекс колонки
    df.reset_index(inplace=True)

    # Создаём эксельи сохраняем данные
    name_file = 'Анализ МО по компании.xlsx'
    sheet_name = 'Данные'  # Наименование вкладки для сводной таблицы
    # writer = pd.ExcelWriter(name_file, engine='xlsxwriter')
    with pd.ExcelWriter(name_file, engine='xlsxwriter') as writer:
        workbook = writer.book
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        wks1 = writer.sheets[sheet_name]  # Сохраняем в переменную вкладку для форматирования

        # Получаем словари форматов для эксель
        header_format, con_format, border_storage_format_left, border_storage_format_right, \
        name_format, MO_format, data_format = format_custom(workbook)

        for col_num, value in enumerate(df.columns.values):
            wks1.write(0, col_num, value, header_format)

        # Форматируем таблицу
        wks1.set_default_row(12)
        wks1.set_row(0, 40, None)
        wks1.set_column('A:A', 12, name_format)
        wks1.set_column('B:B', 32, name_format)
        wks1.set_column('C:AA', 10, data_format)

        # Делаем жирным рамку между складами и форматируем колонку с МО по всем складам
        i = 2
        while i < col_end+1:
            if i < 18:
                wks1.set_column(i, i, None, border_storage_format_left)
                wks1.set_column(i+1, i+1, None, MO_format)
                wks1.set_column(i+2, i+2, None, border_storage_format_right)
            elif 18 < i < 21:
                wks1.set_column(i, i, None, border_storage_format_left)
                wks1.set_column(i, i, None, MO_format)
                # wks1.set_column(i + 2, i + 2, None, border_storage_format_right)
            else:
                wks1.set_column(i-1, i-1, None, border_storage_format_left)
                wks1.set_column(i, i, None, MO_format)
                # wks1.set_column(i + 2, i + 2, None, border_storage_format_right)
            i += 3

        # Подставляем формулу в колонку с МО по всей компании
        f = 2
        while f-1 <= row_end:
            wks1.write_formula(f'X{f}', f'=IF(OR(Z{f}>=1,Z{f}=0),SUM(INT(D{f}),INT(G{f}),'
                                        f'INT(J{f}),INT(M{f}),INT(P{f}),INT(S{f}),INT(U{f})),Z{f})')
            f += 1

        # Добавляем выделение цветом строки при МО=0 по всей компании
        wks1.conditional_format(f'A2:Z{row_end_str}', {'type': 'formula',
                                                       'criteria': '=AND($X2=0,$W2<>0)',
                                                       'format': con_format})

        # Добавляем фильтр в первую колонку
        wks1.autofilter(0, 0, row_end+1, col_end)
        wks1.set_column(col_end+1, col_end+1, None, None, {'hidden': 1})
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


def final_calc(row):
    """
    На основании данных продаж из колонки "Компания MaCar продажи" корректируем значения в колонках "... МО расчёт"
    Так же на основании этих данных добавляем колонку "05 Павловский МО расчёт" и подставляем значения.
    Добавляем колонку "Компания MaCar МО расчёт", которая является суммой целых чисел всех колонок с
    наименованием "... МО расчёт"
    Если МО на соответствующем складе > 0, но < 1, то значение в колонки "... МО расчёт" устанавливаем равное этому МО.
    Если "Компания MaCar продажи" = 0, то устанавливаем во все колонки "... МО расчёт" = 0.5 и в колонку
    "05 Павловский МО расчёт" = 0.5.
    Если "Компания MaCar продажи" > 0, но < 25, то устанавливаем в колонку "05 Павловский МО расчёт" = 0.
    Если "Компания MaCar продажи" >=25, то устанавливаем в колонку
    "05 Павловский МО расчёт" = "Компания MaCar продажи" // 10.
    Если "Компания MaCar продажи" = 1, то устанавливаем в колонке 02 Автолюбитель МО расчет = 1.
    Если "Компания MaCar продажи" = 2, то устанавливаем в колонке 02 Автолюбитель МО расчет = 1.
    В других колонках где есть 1 убираем одну ихз них.
    Если "Компания MaCar продажи" > 0, но < 10, и значение в колонке "... МО расчёт" = 0, то устанавливаем 0.33.
    Если "Компания MaCar продажи" >= 10, и значение в колонке "... МО расчёт" = 0, то устанавливаем 1.
    Суммируем значения из всех колонок "... МО расчёт" и подставляем в колонку "Компания MaCar МО расчёт".
    Если "Компания MaCar продажи" = 0, то устанавливаем "Компания MaCar МО расчёт" = 0.5.
    Если значение в колонке "Компания MaCar МО" > 0, но < 1, то устанавливаем в колонку "Компания MaCar МО расчёт"
    значение из колонки "Компания MaCar МО".

    :param row: Series pandas
        наименование колонок:
        02 Автолюбитель продажи
        08 Центр продажи
        04 Победа продажи
        09 Вокзалка продажи
        01 Кирова продажи
        03 Интер продажи
        Компания MaCar продажи
        01 Кирова МО расчёт
        02 Автолюбитель МО расчёт
        03 Интер МО расчёт
        04 Победа МО расчёт
        08 Центр МО расчёт
        09 Вокзалка МО расчёт
        05 Павловский МО
        Компания MaCar МО
        02 Автолюбитель МО
        04 Победа МО
        08 Центр МО
        01 Кирова МО
        03 Интер МО
        09 Вокзалка МО
        Компания MaCar МО техническое
    :return: Series pandas
        добавлено к полученным колонкам ещё две:
        05 Павловский МО расчёт
        Компания MaCar МО расчёт
    """
    list_sol = ['01 Кирова МО расчёт', '02 Автолюбитель МО расчёт', '03 Интер МО расчёт',
                '04 Победа МО расчёт', '08 Центр МО расчёт', '09 Вокзалка МО расчёт']

    if row['Компания MaCar продажи'] == 0:
        for i in list_sol:
            row[i] = set_value_mo(row[i[:-7]], 0.5)
            row['05 Павловский МО расчёт'] = set_value_mo(row['05 Павловский МО'], 0.5)
    elif row['Компания MaCar продажи'] == 1:
        if 0 < row['02 Автолюбитель МО'] < 1:
            for i in list_sol:
                if row[i] == 0:
                    row[i] = set_value_mo(row[i[:-7]], 0.33)
                else:
                    row[i] = set_value_mo(row[i[:-7]], row[i])
        else:
            for i in list_sol:
                row[i] = set_value_mo(row[i[:-7]], 0.33)
            row['02 Автолюбитель МО расчёт'] = set_value_mo(row['02 Автолюбитель МО'], 1)
    elif row['Компания MaCar продажи'] == 2:
        if 0 < row['02 Автолюбитель МО'] < 1 or row['02 Автолюбитель МО расчёт'] >= 1:
            for i in list_sol:
                if row[i] == 0:
                    row[i] = set_value_mo(row[i[:-7]], 0.33)
                else:
                    row[i] = set_value_mo(row[i[:-7]], row[i])
        else:
            a = 0
            for i in reversed(list_sol):
                if row[i] == 0:
                    row[i] = set_value_mo(row[i[:-7]], 0.33)
                elif row[i] > 0 and a == 0:
                    row[i] = set_value_mo(row[i[:-7]], 0.33)
                    a += 1
                else:
                    row[i] = set_value_mo(row[i[:-7]], row[i])
            row['02 Автолюбитель МО расчёт'] = set_value_mo(row['02 Автолюбитель МО'], 1)

    elif 2 < row['Компания MaCar продажи'] < 10:
        for i in list_sol:
            if row[i] == 0:
                row[i] = set_value_mo(row[i[:-7]], 0.33)
            else:
                row[i] = set_value_mo(row[i[:-7]], row[i])
    elif row['Компания MaCar продажи'] >= 10:
        for i in list_sol:
            if row[i] == 0:
                row[i] = set_value_mo(row[i[:-7]], 1)
            else:
                row[i] = set_value_mo(row[i[:-7]], row[i])
    if row['Компания MaCar продажи'] >= 25:
        val = row['Компания MaCar продажи'] // 10
        row['05 Павловский МО расчёт'] = set_value_mo(row['05 Павловский МО'], val)
    elif row['Компания MaCar продажи'] > 0:
        row['05 Павловский МО расчёт'] = set_value_mo(row['05 Павловский МО'], 0)

    row['Компания MaCar МО расчёт'] = 0
    for i in list_sol:
        row['Компания MaCar МО расчёт'] += int(row[i])
    row['Компания MaCar МО расчёт'] += int(row['05 Павловский МО расчёт'])
    if row['Компания MaCar продажи'] == 0:
        row['Компания MaCar МО расчёт'] = 0.5
    row['Компания MaCar МО расчёт'] = set_value_mo(row['Компания MaCar МО техническое'], row['Компания MaCar МО расчёт'])
    return row


def set_value_mo(mo_old: float, value: float) -> float:
    """Если mo_old от 0 до 1, то возвращаем его. В противном случае возвращаем значение value"""
    return mo_old if 0 < mo_old < 1 else value


if __name__ == '__main__':
    salesFilelist = search_file(salesName)  # запускаем функцию по поиску файлов и получаем список файлов
    minStockFilelist = search_file(minStockName)  # запускаем функцию по поиску файлов и получаем список файлов
    df_sales = create_df(salesFilelist, salesName)
    df_minStock = create_df(minStockFilelist, minStockName)
    df_general = concat_df(df_sales, df_minStock)
    df_general['Компания MaCar МО техническое'] = df_general['Компания MaCar МО']
    df_general = df_general.fillna(0).apply(final_calc, axis=1)  # осуществляем последние расчёты по каждой строке
    df_general = sort_df(df_general)  # сортируем столбцы
    # df_general.to_excel ('test.xlsx')  # записываем полученные данные в эксель для тестов.
    df_write_xlsx(df_general)
