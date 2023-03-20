# Author Loik Andrey 7034@balancedv.ru

import pandas as pd
import os

FOLDER = 'Исходные данные'
salesName = 'продажи'
minStockName = 'МО'
DF_COMP_CRAT = pd.DataFrame()
LIST_WH = ['01 Кирова', '02 Автолюбитель', '03 Интер', '04 Победа', '08 Центр', '09 Вокзалка']


def input_prior_wh():
    """Запрашиваем приоритетный склад у пользователя"""
    print('Введите приоритетный склад из списка ' + ', '.join(LIST_WH))
    prior = input('Если приоритетного склада нет, то нажмите Enter: ')

    if prior not in LIST_WH:
        print(f'Вы не ввели приоритетный склад, либо наименования {prior} нет в списке.')
        prior = False
    return prior


def input_max_crat_comp():
    """Запрашиваем режим использования Кратности и Комплектности"""
    print('Введите 1, если необходимо учитывать Кратность или Комплектность')
    crat, comp, const_crat = 1, 1, False
    try:
        result = int(input('Если учитывать не нужно, то нажмите Enter: '))
    except ValueError:
        result = 0

    try:
        if result:
            print('Установите режим работы с Кратностью и Комплекностью.')
            print('1 - Присваиваем значение кратности программой. Комплектность не учитываем.')
            print('2 - Ограничиваем максимальное значение Кратности и Комплектности.')
            result = int(input('Введите вариант: '))
        if result == 1:
            print('Вы выбрали режим установки кратности программой.')
            global CONST_CRAT
            const_crat = int(input('Введите значение Кратности: '))
            crat = const_crat
        elif result == 2:
            crat = int(input('Введите максимальное значение Кратности: '))
            comp = int(input('Введите максимальное значение Комплектности: '))
    except ValueError:
        print('Вы ввели не корректное значение. Это должно быть целое число. Попробуйте ещё раз.')
        input_max_crat_comp()
    return crat, comp, const_crat


def input_day_sales_stock():
    """Запрашиваем кол-во анализируемых дней и дней запаса"""
    print('Введите 1, если необходимо изменить кол-во дней анализируемых продаж и дней запаса')
    day_sales, day_stock = 365, 90
    try:
        result = int(input('Если используем дни по умолчанию (365 и 90), то нажмите Enter: '))
    except ValueError:
        result = 0
    if result:
        try:
            day_sales = int(input('Введите кол-во дней анализа продаж: : '))
            day_stock = int(input('Введите кол-во дней запаса: '))
        except ValueError:
            print('Вы ввели не корректное значение. Это должно быть целое число. Попробуйте ещё раз.')
            input_day_sales_stock()

    return day_sales, day_stock


def input_use_opt_store():
    """Запрашиваем использование оптового склада"""
    print('Введите 1, если требуется учитывать продажи оптового склада')
    use_opt_store = False
    try:
        result = int(input('Если учитывать не нужно, то нажмите Enter: '))
    except ValueError:
        result = 0
    if result:
        print('Включен режим использования оптового склада')
        use_opt_store = True
    return use_opt_store


def search_file(name):
    """
    :param name: Поиск всех файлов в папке FOLDER, в наименовании которых, содержится name
    :return: filelist список с наименованиями фалов
    """
    filelist = []
    for item in os.listdir(FOLDER):
        condition_opt = True
        if not USE_OPT_STORE and name == 'продажи':
            condition_opt = 'техснаб' not in item.lower()
        # если файл содержит name и с расширением .xlsx, то выполняем
        if name in item and item.endswith('.xlsx') and condition_opt:
            # Добавляем в список папку и имя файла для последующего обращения из списка
            filelist.append(FOLDER + "/" + item)
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
        f = df_search_header[mask].dropna(axis=0, how='all').index.values  # Удаление пустых строк
        df = df.iloc[int(f):, :]  # Убираем все строки с верха DF до заголовков
        df = df.dropna(axis=1, how='all')  # Убираем пустые колонки
        df.iloc[0, :] = df.iloc[0, :] + ' ' + add_name  # Добавляем в наименование тип данных
        if add_name == 'продажи':
            df.iloc[0, :4] = ['Код', 'Комплектность', 'Кратность', 'Номенклатура']
        else:
            df.iloc[0, :2] = ['Код', 'Номенклатура']
        # df.iloc[0, 1] = 'Номенклатура'
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
    Пытаемся прочитать файл xlxs, если не получается, то исправляем ошибку и опять читаем файл.
    :param file_name: Имя файла для чтения.
    :return: DataFrame
    """
    print('Попытка загрузки файла:' + file_name)
    try:
        if 'продажи' in file_name:
            df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=1, engine='openpyxl')
        else:
            df = pd.read_excel(file_name, sheet_name='TDSheet', header=None, skipfooter=0, engine='openpyxl')
        return df
    except KeyError as Error:
        print(Error)
        df = None
        if str(Error) == "\"There is no item named 'xl/sharedStrings.xml' in the archive\"":
            bug_fix(file_name)
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
    df = pd.concat([df1, df2], axis=1, ignore_index=False)
    return df


def sort_df(df):
    sort_list = ['01 Кирова продажи', '01 Кирова МО', '01 Кирова МО расчёт',
                 '02 Автолюбитель продажи', '02 Автолюбитель МО', '02 Автолюбитель МО расчёт',
                 '03 Интер продажи', '03 Интер МО', '03 Интер МО расчёт',
                 '04 Победа продажи', '04 Победа МО', '04 Победа МО расчёт',
                 '08 Центр продажи', '08 Центр МО', '08 Центр МО расчёт',
                 '09 Вокзалка продажи', '09 Вокзалка МО', '09 Вокзалка МО расчёт',
                 '05 Павловский МО', '05 Павловский МО расчёт',
                 'Розн. продажи MaCar', 'Компания MaCar МО', 'Компания MaCar МО расчёт',
                 'Компания MaCar МО техническое']
    if USE_OPT_STORE:
        sort_list = ['01 Кирова продажи', '01 Кирова МО', '01 Кирова МО расчёт',
                 '02 Автолюбитель продажи', '02 Автолюбитель МО', '02 Автолюбитель МО расчёт',
                 '03 Интер продажи', '03 Интер МО', '03 Интер МО расчёт',
                 '04 Победа продажи', '04 Победа МО', '04 Победа МО расчёт',
                 '08 Центр продажи', '08 Центр МО', '08 Центр МО расчёт',
                 '09 Вокзалка продажи', '09 Вокзалка МО', '09 Вокзалка МО расчёт',
                 '05 Павловский продажи', '05 Павловский МО', '05 Павловский МО расчёт',
                 'Розн. продажи MaCar', 'Компания MaCar МО', 'Компания MaCar МО расчёт',
                 'Компания MaCar МО техническое']
    df = df[sort_list]
    return df


def payment(df_payment):
    """Расчитываем значения МО расчет с учётом Кратности и Комплектности"""

    columns_payment = ['01 Кирова МО расчёт', '02 Автолюбитель МО расчёт', '03 Интер МО расчёт',
                       '04 Победа МО расчёт', '08 Центр МО расчёт', '09 Вокзалка МО расчёт']

    columns_sales = '01 Кирова|02 Автолюбитель|03 Интер|04 Победа|08 Центр|09 Вокзалка'

    if USE_OPT_STORE:
        columns_payment += ['05 Павловский МО расчёт']
        columns_sales += '|05 Павловский'

    DF_COMP_CRAT['Комплектность'] = df_payment['Комплектность'].fillna(1).max(axis=1)
    DF_COMP_CRAT['Кратность'] = df_payment['Кратность'].fillna(1).max(axis=1)

    if CONST_CRAT:
        print(f'Устанавливаем значение кратности равное {CONST_CRAT}')
        DF_COMP_CRAT['Кратность'] = CONST_CRAT
    else:
        DF_COMP_CRAT['Кратность'][DF_COMP_CRAT['Кратность'] >= MAX_CRAT] = MAX_CRAT
    DF_COMP_CRAT['Комплектность'][DF_COMP_CRAT['Комплектность'] >= MAX_COMP] = MAX_COMP

    if USE_OPT_STORE:
        df_payment['Розн. продажи MaCar'] = df_payment.filter(regex=columns_sales[:-14]).sum(axis=1)
    else:
        df_payment['Розн. продажи MaCar'] = df_payment.filter(regex=columns_sales).sum(axis=1)

    for i in range(len(columns_payment)):
        # Суммируем продажи по всем складам подразделения
        df_payment[columns_payment[i][:-10] + ' продажи'] = df_payment.filter(like=columns_payment[i][:-10]).sum(axis=1)

        # Рассчитываем значения исходя из периода анализируемых продажи и периода запасов
        df_payment[columns_payment[i]] = df_payment[columns_payment[i][:-10] + ' продажи'].fillna(0).apply(calc)

        # Обнуляем все рассчитанные МО меньше 0
        mask_lz = df_payment[columns_payment[i]] < 0
        df_payment.loc[mask_lz, columns_payment[i]] = 0

        if columns_payment[i] != '05 Павловский МО расчёт':
            # Приравниваем расчётные значения от 0 до 1 к значению кратности
            mask_min_crat = (df_payment[columns_payment[i]] > 0) & \
                            (df_payment[columns_payment[i]] < DF_COMP_CRAT['Кратность'])
            df_payment.loc[mask_min_crat, columns_payment[i]] = DF_COMP_CRAT.loc[mask_min_crat, 'Кратность']

            # Если есть расчётные значения больше 0 до значения Комплектности, то устанавливаем значения комплектности
            mask_min_comp = (df_payment[columns_payment[i]] > 0) & \
                            (df_payment[columns_payment[i]] < DF_COMP_CRAT['Комплектность'])
            df_payment.loc[mask_min_comp, columns_payment[i]] = DF_COMP_CRAT.loc[mask_min_comp, 'Комплектность']

            # Все расчётные значения больше комплектности устанавливаем согласно кратности
            df_payment.loc[~mask_min_comp, columns_payment[i]] = round(
                (df_payment.loc[~mask_min_comp, columns_payment[i]] + 0.000001) /
                DF_COMP_CRAT.loc[~mask_min_comp, 'Кратность']) * DF_COMP_CRAT.loc[~mask_min_comp, 'Кратность']
    return df_payment


def calc(val):
    quan_prod = val / DAY_SALES * DAY_STOCK
    return int(quan_prod) if quan_prod >= 1 else quan_prod


def df_write_xlsx(df):
    # Сохраняем в переменные значения конечных строк и столбцов
    row_end, col_end = len(df), len(df.columns)
    row_end_str, col_end_str = str(row_end), str(col_end)

    # Для простоты форматирования переводим индекс колонки
    df.reset_index(inplace=True)

    # Создаём эксель и сохраняем данные
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
        while i < col_end + 1:
            if USE_OPT_STORE:
                wks1.set_column(i, i, None, border_storage_format_left)
                wks1.set_column(i + 1, i + 1, None, MO_format)
                wks1.set_column(i + 2, i + 2, None, border_storage_format_right)
            else:
                if i < 18:
                    wks1.set_column(i, i, None, border_storage_format_left)
                    wks1.set_column(i + 1, i + 1, None, MO_format)
                    wks1.set_column(i + 2, i + 2, None, border_storage_format_right)
                elif 18 < i < 21:
                    wks1.set_column(i, i, None, border_storage_format_left)
                    wks1.set_column(i, i, None, MO_format)
                    # wks1.set_column(i + 2, i + 2, None, border_storage_format_right)
                else:
                    wks1.set_column(i - 1, i - 1, None, border_storage_format_left)
                    wks1.set_column(i, i, None, MO_format)
                    # wks1.set_column(i + 2, i + 2, None, border_storage_format_right)
            i += 3

        # Подставляем формулу в колонку с МО по всей компании
        f = 2
        while f - 1 <= row_end:
            if USE_OPT_STORE:
                wks1.write_formula(f'Y{f}', f'=IF(OR(AA{f}>=1,AA{f}=0),SUM(INT(D{f}),INT(G{f}),'
                                            f'INT(J{f}),INT(M{f}),INT(P{f}),INT(S{f}),INT(V{f})),AA{f})')
            else:
                wks1.write_formula(f'X{f}', f'=IF(OR(Z{f}>=1,Z{f}=0),SUM(INT(D{f}),INT(G{f}),'
                                            f'INT(J{f}),INT(M{f}),INT(P{f}),INT(S{f}),INT(U{f})),Z{f})')
            f += 1

        # Добавляем выделение цветом строки при МО=0 по всей компании
        if USE_OPT_STORE:
            wks1.conditional_format(f'A2:Z{row_end_str}', {'type': 'formula',
                                                           'criteria': '=AND($Y2=0,$X2<>0)',
                                                           'format': con_format})
        else:
            wks1.conditional_format(f'A2:Z{row_end_str}', {'type': 'formula',
                                                           'criteria': '=AND($X2=0,$W2<>0)',
                                                           'format': con_format})

        # Добавляем фильтр в первую колонку
        wks1.autofilter(0, 0, row_end + 1, col_end)
        wks1.set_column(col_end + 1, col_end + 1, None, None, {'hidden': 1})
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
    На основании данных продаж из колонки "Розн. продажи MaCar" корректируем значения в колонках "... МО расчёт"
    Так же на основании этих данных добавляем колонку "05 Павловский МО расчёт" и подставляем значения.
    Добавляем колонку "Компания MaCar МО расчёт", которая является суммой целых чисел всех колонок с
    наименованием "... МО расчёт"
    Если МО на соответствующем складе > 0, но < 1, то значение в колонки "... МО расчёт" устанавливаем равное этому МО.
    Если "Розн. продажи MaCar" = 0, то устанавливаем во все колонки "... МО расчёт" = 0.5 и в колонку
    "05 Павловский МО расчёт" = 0.5.
    Если "Розн. продажи MaCar" > 0, но < 25, то устанавливаем в колонку "05 Павловский МО расчёт" = 0.
    Если "Розн. продажи MaCar" >=25, то устанавливаем в колонку
    "05 Павловский МО расчёт" = "Розн. продажи MaCar" // 10.
    Если "Розн. продажи MaCar" = 1, то устанавливаем в колонке 02 Автолюбитель МО расчет = 1.
    Если "Розн. продажи MaCar" = 2, и они были на разных складах, то устанавливаем в колонке 02 Автолюбитель
    МО расчет = 1. Если они были на одном складе, то устанавливаем МО=1 на складе где была продажа.
    В других колонках где есть 1 убираем одну ихз них.
    Если "Розн. продажи MaCar" > 0, но < 10, и значение в колонке "... МО расчёт" = 0, то устанавливаем 0.33.
    Если "Розн. продажи MaCar" >= 10, и значение в колонке "... МО расчёт" = 0, то устанавливаем 1.
    Суммируем значения из всех колонок "... МО расчёт" и подставляем в колонку "Компания MaCar МО расчёт".
    Если "Розн. продажи MaCar" = 0, то устанавливаем "Компания MaCar МО расчёт" = 0.5.
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
        Розн. продажи MaCar
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
        Компания MaCar МО техническое.
    :return: Series pandas
        добавлено к полученным колонкам ещё две:
        05 Павловский МО расчёт
        Компания MaCar МО расчёт
    """
    list_col = ['01 Кирова МО расчёт', '02 Автолюбитель МО расчёт', '03 Интер МО расчёт',
                '04 Победа МО расчёт', '08 Центр МО расчёт', '09 Вокзалка МО расчёт']
    prior_wh_mo, prior_wh_calc = None, None
    if PRIOR_WH:
        prior_wh_mo = PRIOR_WH + ' МО'
        prior_wh_calc = PRIOR_WH + ' МО расчёт'

    if row['Розн. продажи MaCar'] <= 0:
        for i in list_col:
            row[i] = set_value_mo(row[i[:-7]], 0.5)
    elif row['Розн. продажи MaCar'] == 1:
        for i in list_col:
            if PRIOR_WH:
                # Устанавливаем ... МО расчёт согласно приоритетного склада
                if 0 < row[prior_wh_mo] < 1:
                    row[i] = set_value_mo(row[i[:-7]], 0.33 if row[i] == 0 else row[i])
                else:
                    row[prior_wh_calc] = \
                        set_value_mo(row[prior_wh_mo], row[i]) if row[i] > row[prior_wh_calc] else row[prior_wh_calc]
                    if i != prior_wh_calc:
                        row[i] = set_value_mo(row[i[:-7]], 0.33)
            else:
                # Устанавливаем ... МО расчёт согласно приоритетного склада
                row[i] = set_value_mo(row[i[:-7]], 0.33 if row[i] == 0 else row[i])
    elif row['Розн. продажи MaCar'] == 2:
        if PRIOR_WH:
            if 0 < row[prior_wh_mo] < 1 or row[prior_wh_calc] >= 1:
                for i in list_col:
                    row[i] = set_value_mo(row[i[:-7]], 0.33 if row[i] == 0 else row[i])
            else:
                # Переставляем приоритетный склад вперёд списка и переворачиваем остальной список
                list_col1 = [prior_wh_calc] + [col for col in reversed(list_col) if col != prior_wh_calc]
                first_pass = True
                for i in list_col1:
                    if row[i[:-9] + 'продажи'] == 0:
                        row[i] = set_value_mo(row[i[:-7]], 0.33)
                    elif row[i[:-9] + 'продажи'] == 1 and first_pass:
                        if row[i[:-7]] >= 1 or row[i[:-7]] == 0:
                            row[prior_wh_calc] = set_value_mo(row[prior_wh_mo], row[i])
                            row[i] = set_value_mo(row[i[:-7]], 0.33)
                            first_pass = False
                        else:
                            row[i] = set_value_mo(row[i[:-7]], row[i])
                    else:
                        row[i] = set_value_mo(row[i[:-7]], row[i])
        else:
            for i in list_col:
                row[i] = set_value_mo(row[i[:-7]], 0.33 if row[i] == 0 else row[i])

    elif 2 < row['Розн. продажи MaCar'] < 10:
        for i in list_col:
            row[i] = set_value_mo(row[i[:-7]], 0.33 if row[i] == 0 else row[i])
    elif row['Розн. продажи MaCar'] >= 10:
        for i in list_col:
            val_default = row[['Комплектность', 'Кратность']].max()
            row[i] = set_value_mo(row[i[:-7]], val_default if row[i] == 0 else row[i])

    if not USE_OPT_STORE:
        row['05 Павловский МО расчёт'] = 0

    if row['Розн. продажи MaCar'] <= 0:
        row['05 Павловский МО расчёт'] = 0.5
    elif row['Розн. продажи MaCar'] >= 25:
        val = row['Розн. продажи MaCar'] // 10
        row['05 Павловский МО расчёт'] = val + row['05 Павловский МО расчёт']
    elif row['Розн. продажи MaCar'] > 0:
        row['05 Павловский МО расчёт'] = 0 + row['05 Павловский МО расчёт']

    # Приводим рассчитанные значение в колонке '05 Павловский МО расчёт' к значению Кратности или Комплектности
    if (row['05 Павловский МО расчёт'] > 0) & (row['05 Павловский МО расчёт'] < row['Кратность']):
        row['05 Павловский МО расчёт'] = set_value_mo(row['05 Павловский МО'], row['Кратность'])

    if (row['05 Павловский МО расчёт'] > 0) & (row['05 Павловский МО расчёт'] < row['Комплектность']):
        row['05 Павловский МО расчёт'] = set_value_mo(row['05 Павловский МО'], row['Комплектность'])

    if row['05 Павловский МО расчёт'] > row['Комплектность']:
        opt_val = round(
            (row['05 Павловский МО расчёт'] + 0.000001) / row['Кратность']
        ) * row['Кратность']
        row['05 Павловский МО расчёт'] = set_value_mo(row['05 Павловский МО'], opt_val)

    row['05 Павловский МО расчёт'] = set_value_mo(row['05 Павловский МО'], row['05 Павловский МО расчёт'])

    row['Компания MaCar МО расчёт'] = 0
    for i in list_col:
        row['Компания MaCar МО расчёт'] += int(row[i])
    row['Компания MaCar МО расчёт'] += int(row['05 Павловский МО расчёт'])
    if row['Розн. продажи MaCar'] == 0:
        row['Компания MaCar МО расчёт'] = 0.5
    row['Компания MaCar МО расчёт'] = set_value_mo(row['Компания MaCar МО техническое'],
                                                   row['Компания MaCar МО расчёт'])
    return row


def set_value_mo(mo_old: float, value: float) -> float:
    """Если mo_old от 0 до 1, то возвращаем его. В противном случае возвращаем значение value"""
    return mo_old if 0 < mo_old < 1 else value


def create_final_df(df1: object, df2: object):
    """Создаём финальный DataFrame
    :param df1:
    :param df2:

    """
    df_final = concat_df(df1, df2)
    df_final.drop(columns=['Комплектность', 'Кратность'], inplace=True)
    df_final['Компания MaCar МО техническое'] = df_final['Компания MaCar МО']
    df_final = concat_df(df_final, DF_COMP_CRAT)
    df_final[['Комплектность', 'Кратность']] = df_final[['Комплектность', 'Кратность']].fillna(1)
    df_final = df_final.fillna(0).apply(final_calc, axis=1)  # осуществляем последние расчёты по каждой строке
    df_final = sort_df(df_final)  # сортируем столбцы
    return df_final


if __name__ == '__main__':
    MAX_CRAT, MAX_COMP, CONST_CRAT = input_max_crat_comp()  # Задаём максимальное значение Кратности и Комплектности
    PRIOR_WH = input_prior_wh()  # Задаём приоритетный склад
    DAY_SALES, DAY_STOCK = input_day_sales_stock()  # Задаём кол-во дней для расчётов
    USE_OPT_STORE = input_use_opt_store()  # Задаём использование оптового склада
    salesFilelist = search_file(salesName)  # Запускаем функцию по поиску файлов и получаем список файлов
    minStockFilelist = search_file(minStockName)  # Запускаем функцию по поиску файлов и получаем список файлов
    df_sales = create_df(salesFilelist, salesName)

    df_minStock = create_df(minStockFilelist, minStockName)

    df_general = create_final_df(df_sales, df_minStock)
    # df_general.to_excel ('test.xlsx')  # Записываем полученные данные в эксель для тестов.
    df_write_xlsx(df_general)
