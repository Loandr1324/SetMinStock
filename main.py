# Author Loik Andrey 7034@balancedv.ru
import math
import pandas as pd
import os

FOLDER = 'Исходные данные'
salesName = 'продажи'
minStockName = 'МО'
DF_COMP_CRAT = pd.DataFrame()
PERIOD_LONG = '1 год'
PERIOD_SHOT = '3 мес'
DEFAULT_COLUMNS_NAMES: list = ['01 Кирова', '02 Автолюбитель', '03 Интер', '04 Победа', '08 Центр', '09 Вокзалка']
VALUE_MO_SKIP = [0.24, 0.25, 0.26, 0.27, 0.28, 0.29]


def input_list_wh():
    """Запрашиваем список складов, которые необходимо отрабатывать согласно приоритету"""
    dict_list_wh = {
        1: '01 Кирова',
        2: '02 Автолюбитель',
        3: '03 Интер',
        4: '04 Победа',
        8: '08 Центр',
        9: '09 Вокзалка'
    }
    list_wh = [value for value in dict_list_wh.values()]
    print('Введите список складов в порядке убывания приоритета.')
    print(f'По умолчанию список складов расположен по следующему порядку: {list_wh}.')
    print('Изменить список можно, введя через запятую номера складов. Например: 8, 3, 2, 9, 1, 4')
    list_key = input('Если порядок по умолчанию менять не нужно, то нажмите Enter:')
    try:
        list_key = list_key.split(',')
        # Проверяем что введены все номера складов, все номера разные и все номера есть в словаре
        list_wh = [dict_list_wh[int(k)] for k in list_key]
        if len(set(list_wh)) != len(dict_list_wh):
            raise ValueError
        print(f'Порядок складов изменён: {list_wh}')
    except (ValueError, KeyError):
        print(f'Ошибка при вводе складов!!! Порядок складов установлен по умолчанию: {list_wh}')

    return list_wh


def input_wh_skip_list():
    """Запрашиваем список складов, которые необходимо пропускать при анализе"""
    dict_list_wh = {
        1: '01 Кирова',
        2: '02 Автолюбитель',
        3: '03 Интер',
        4: '04 Победа',
        8: '08 Центр',
        9: '09 Вокзалка'
    }
    list_wh = [value for value in dict_list_wh.values()]
    print('Введите список складов, которые надо исключить из анализа и не менять на нём МО.')
    print(f'По умолчанию список складов пуст, но Вы можете добавить любой из списка: {list_wh}.')
    print('Выбор можно сделать, вводя через запятую номера складов. Например: 1, 4')
    list_key = input('Если никакие склады исключать не нужно, то нажмите Enter:')
    try:
        list_key = list_key.split(',')
        # Проверяем что введены все номера складов, все номера разные и все номера есть в словаре
        list_wh = [dict_list_wh[int(k)] for k in list_key]
        print(f'Данные склады будут исключены из списка: {list_wh}')
    except (ValueError, KeyError):
        print(f'Ошибка при вводе складов!!! Никакие склады исключаться не будут: {list_wh}')
        list_wh = []

    return list_wh


def input_prior_wh():
    """Запрашиваем приоритетный склад у пользователя"""
    print('Введите цифру приоритетного склада из списка ' + ', '.join(LIST_WH))
    key = input('Если приоритетного склада нет, то нажмите Enter: ')
    dict_list_wh = {
        1: '01 Кирова',
        2: '02 Автолюбитель',
        3: '03 Интер',
        4: '04 Победа',
        8: '08 Центр',
        9: '09 Вокзалка'
    }
    prior = False
    try:
        prior = dict_list_wh[int(key)]
        print(f'Установлен приоритетный склад {prior}')
    except ValueError:
        print(f'Вы не ввели приоритетный склад, либо {key} не является номером склада.')
    except KeyError:
        print(f'В списке нет склада с номером {key}. Попробуйте ещё раз.')
        input_prior_wh()
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
            # global CONST_CRAT
            const_crat = int(input('Введите значение Кратности: '))
            # crat = const_crat
        elif result == 2:
            crat = int(input('Введите максимальное значение Кратности: '))
            comp = int(input('Введите максимальное значение Комплектности: '))
    except ValueError:
        print('Вы ввели не корректное значение. Это должно быть целое число. Попробуйте ещё раз.')
        input_max_crat_comp()
    return crat, comp, const_crat


def input_day_sales(day_sales, period):
    """Запрашиваем кол-во анализируемых дней и дней запаса"""
    print(f'Введите 1, если необходимо изменить кол-во дней анализируемых продаж за {period}')
    try:
        result = int(input(f'Если используем дни по умолчанию ({day_sales}), то нажмите Enter: '))
    except ValueError:
        result = 0
    if result:
        try:
            day_sales = int(input('Введите кол-во дней анализа продаж: : '))
        except ValueError:
            print('Вы ввели не корректное значение. Это должно быть целое число. Попробуйте ещё раз.')
            input_day_sales(day_sales, period)
    return day_sales


def input_day_stock():
    """Устанавливаем количество дней запаса для всех периодов"""
    print('Введите 1, если необходимо изменить кол-во дней запаса для всех периодов')
    day_stock = 91
    try:
        result = int(input(f'Если используем дни по умолчанию ({day_stock}), то нажмите Enter: '))
    except ValueError:
        result = 0
    if result:
        try:
            day_stock = int(input('Введите кол-во дней запаса: '))
        except ValueError:
            print('Вы ввели не корректное значение. Это должно быть целое число. Попробуйте ещё раз.')
            input_day_stock()

    return day_stock


def input_use_opt_store():
    """Запрашиваем необходимость учитывать продажи оптового склада для расчёта МО на оптовом складе"""
    print('Введите 1, если требуется учитывать продажи оптового склада для установки МО оптового склада')
    use_opt_store = False

    try:
        result = int(input('Если учитывать не нужно, то нажмите Enter: '))
    except ValueError:
        result = 0
    if result:
        print('Включен режим использования оптового склада')
        use_opt_store = True
    return use_opt_store


def input_day_opt_stock():
    """Запрашиваем использование оптового склада"""
    print('Введите 1, если требуется держать дополнительный запас на оптовом складе')
    day_opt_stock = DAY_STOCK
    min_value_opt_stock = 4
    name_columns = ''
    try:
        result = int(input('Если не нужно, то нажмите Enter: '))
    except ValueError:
        result = 0
    if result:
        # Получаем наименование колонки на основании которой будет производиться расчёт
        name_columns = input_name_period_opt_stock()
        try:
            print()
            print('Введите минимальную сумму расчётных значений продаж при котором необходимо держать '
                  'дополнительный запас на оптовом складе: ')
            min_value_opt_stock = int(input(f'По умолчанию используется ({min_value_opt_stock}):'))
            if min_value_opt_stock < 1:
                print('Минимальная сумму расчётных значений МО не может быть меньше 1. Попробуйте ещё раз.')
                input_day_opt_stock()
        except ValueError:
            print(f'Используем значение по умолчанию: {min_value_opt_stock}')

        try:
            print()
            day_opt_stock = int(input('Введите кол-во дней запаса для оптового склада: '))
            if day_opt_stock < DAY_STOCK:
                print('Кол-во дней запаса оптового склада не может быть меньше кол-ва дней запаса розничного склада. '
                      'Попробуйте ещё раз.')
                input_day_opt_stock()
        except ValueError:
            print('Вы ввели не корректное значение. Это должно быть целое число. Попробуйте ещё раз.')
            input_day_opt_stock()

    return day_opt_stock, min_value_opt_stock, name_columns


def input_name_period_opt_stock():
    """Устанавливаем по какому периоду будет производиться расчёт дополнительного запаса на оптовом складе"""
    print('Выберите период анализа на основании которого будет рассчитываться дополнительный запас на Оптовом складе')
    list_name_period = [f'Розн. MaCar МО расчёт за {PERIOD_LONG}',
                        f'Розн. MaCar МО расчёт за {PERIOD_SHOT}',
                        'Розн. MaCar МО среднее значение']
    try:
        number_period = int(input('1 - Длинный период, 2 - Короткий период, 3 - Среднее значение, '
                                  'по умолчанию Длинный период: '))
        if number_period not in [1, 2, 3]:
            raise ValueError
    except ValueError:
        number_period = 1
        print('Вы ввели не корректное значение. Используем значение по умолчанию: Длинный период')
    name_period = list_name_period[number_period - 1]
    return name_period


def search_file(name):
    """
    :param name: Поиск всех файлов в папке FOLDER, в наименовании которых, содержится name
    :return: filelist список с наименованиями фалов
    """
    filelist = []
    for item in os.listdir(FOLDER):
        condition_opt = True
        if not USE_OPT_STORE and name == 'продажи':
            keywords = ['техснаб', 'новотрейд']
            condition_opt = not any(keyword in item.lower() for keyword in keywords)
        # если файл содержит name и с расширением .xlsx, то выполняем
        if name in item and item.endswith('.xlsx') and condition_opt:
            # Добавляем в список папку и имя файла для последующего обращения из списка
            filelist.append(FOLDER + "/" + item)
    print(filelist)
    return filelist


def create_filelist_one_tow_year(filelist: list) -> tuple:
    """Получаем на вход список с наименованием файлов и возвращаем три списка."""
    print(filelist)
    filelist_one = [file for file in filelist if "1 год" in file]
    filelist_two = [file for file in filelist if "2 года" in file]
    filelist_three_month = [file for file in filelist if "3 мес" in file]

    return filelist_one, filelist_two, filelist_three_month


def create_df(file_list, add_name=''):
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
        # Добавляем в наименование тип данных
        if add_name:
            df.iloc[0, :] = df.iloc[0, :] + ' ' + add_name

        # if add_name == 'продажи':
        #     df.iloc[0, :] = df.iloc[0, :] + '1'
        #     df.iloc[0, :4] = ['Код', 'Комплектность', 'Кратность', 'Номенклатура']
        # else:
        #     df.iloc[0, :2] = ['Код', 'Номенклатура']
        df.iloc[0, :2] = ['Код', 'Номенклатура']
        # df.iloc[0, 1] = 'Номенклатура'
        # Значения из найденной строки переносим в заголовки DataFrame для простоты дальнейшего обращения
        df.columns = df.iloc[0].values
        if not add_name:
            df = df.iloc[1:, :]  # Убираем одну с верха DF
        else:
            df = df.iloc[2:, :]  # Убираем две строки с верха DF
        df['Номенклатура'] = df['Номенклатура'].str.strip()  # Удалить пробелы с обоих концов строки в ячейке
        # Переносим колонки в индекс, для упрощения дальнейшей работы
        df.set_index(['Код', 'Номенклатура'], inplace=True)

        # Добавляем преобразованный DF в результирующий DF
        if df_result is None:
            df_result = df
        else:
            df_result = concat_df(df_result, df)
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
            df = pd.read_excel(file_name, header=None, skipfooter=1, engine='openpyxl')
        else:
            df = pd.read_excel(file_name, header=None, skipfooter=0, engine='openpyxl')
        return df
    except KeyError as Error:
        print(Error)
        df = None
        if str(Error) == "\"There is no item named 'xl/sharedStrings.xml' in the archive\"":
            bug_fix(file_name)
            print('Исправлена ошибка: ', Error, f'в файле: \"{file_name}\"\n')
            if 'продажи' in file_name:
                df = pd.read_excel(file_name, header=None, skipfooter=1, engine='openpyxl')
            else:
                df = pd.read_excel(file_name, header=None, skipfooter=0, engine='openpyxl')
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
    # df1 = df1.reset_index(inplace=True, level=['Номенклатура'])
    df1.reset_index(inplace=True, level=['Номенклатура'])
    df2 = df2.reset_index(level=['Номенклатура'])
    df2 = df2.drop(['Номенклатура'], axis=1)
    df = pd.concat([df1, df2], axis=1, ignore_index=False)
    df.reset_index(inplace=True)
    df.set_index(['Код', 'Номенклатура'], inplace=True)
    return df


def sort_df(df: pd.DataFrame) -> pd.DataFrame:
    """Сортируем солонки для дальнейшей работы алгоритма программы"""
    sort_list = []
    if not USE_OPT_STORE:
        sort_list = [
            'Комплектность', 'Кратность',
            f'01 Кирова продажи за {PERIOD_LONG}', f'01 Кирова продажи за {PERIOD_SHOT}', '01 Кирова МО',
            f'01 Кирова МО расчёт за {PERIOD_LONG}', f'01 Кирова МО расчёт за {PERIOD_SHOT}',
            f'02 Автолюбитель продажи за {PERIOD_LONG}', f'02 Автолюбитель продажи за {PERIOD_SHOT}',
            '02 Автолюбитель МО',
            f'02 Автолюбитель МО расчёт за {PERIOD_LONG}', f'02 Автолюбитель МО расчёт за {PERIOD_SHOT}',
            f'03 Интер продажи за {PERIOD_LONG}', f'03 Интер продажи за {PERIOD_SHOT}', '03 Интер МО',
            f'03 Интер МО расчёт за {PERIOD_LONG}', f'03 Интер МО расчёт за {PERIOD_SHOT}',
            f'04 Победа продажи за {PERIOD_LONG}', f'04 Победа продажи за {PERIOD_SHOT}', '04 Победа МО',
            f'04 Победа МО расчёт за {PERIOD_LONG}', f'04 Победа МО расчёт за {PERIOD_SHOT}',
            f'08 Центр продажи за {PERIOD_LONG}', f'08 Центр продажи за {PERIOD_SHOT}', '08 Центр МО',
            f'08 Центр МО расчёт за {PERIOD_LONG}', f'08 Центр МО расчёт за {PERIOD_SHOT}',
            f'09 Вокзалка продажи за {PERIOD_LONG}', f'09 Вокзалка продажи за {PERIOD_SHOT}', '09 Вокзалка МО',
            f'09 Вокзалка МО расчёт за {PERIOD_LONG}', f'09 Вокзалка МО расчёт за {PERIOD_SHOT}',
            '05 Павловский МО', 'Розн. продажи MaCar за 2 года', 'Компания MaCar МО',
            'Компания MaCar МО из 1С техническое'
        ]
    if USE_OPT_STORE:
        sort_list = [
            'Комплектность', 'Кратность',
            f'01 Кирова продажи за {PERIOD_LONG}', f'01 Кирова продажи за {PERIOD_SHOT}', '01 Кирова МО',
            f'01 Кирова МО расчёт за {PERIOD_LONG}', f'01 Кирова МО расчёт за {PERIOD_SHOT}',
            f'02 Автолюбитель продажи за {PERIOD_LONG}', f'02 Автолюбитель продажи за {PERIOD_SHOT}',
            '02 Автолюбитель МО',
            f'02 Автолюбитель МО расчёт за {PERIOD_LONG}', f'02 Автолюбитель МО расчёт за {PERIOD_SHOT}',
            f'03 Интер продажи за {PERIOD_LONG}', f'03 Интер продажи за {PERIOD_SHOT}', '03 Интер МО',
            f'03 Интер МО расчёт за {PERIOD_LONG}', f'03 Интер МО расчёт за {PERIOD_SHOT}',
            f'04 Победа продажи за {PERIOD_LONG}', f'04 Победа продажи за {PERIOD_SHOT}', '04 Победа МО',
            f'04 Победа МО расчёт за {PERIOD_LONG}', f'04 Победа МО расчёт за {PERIOD_SHOT}',
            f'08 Центр продажи за {PERIOD_LONG}', f'08 Центр продажи за {PERIOD_SHOT}', '08 Центр МО',
            f'08 Центр МО расчёт за {PERIOD_LONG}', f'08 Центр МО расчёт за {PERIOD_SHOT}',
            f'09 Вокзалка продажи за {PERIOD_LONG}', f'09 Вокзалка продажи за {PERIOD_SHOT}', '09 Вокзалка МО',
            f'09 Вокзалка МО расчёт за {PERIOD_LONG}', f'09 Вокзалка МО расчёт за {PERIOD_SHOT}',
            f'05 Павловский продажи за {PERIOD_LONG}', f'05 Павловский продажи за {PERIOD_SHOT}',
            '05 Павловский МО',
            f'05 Павловский МО расчёт за {PERIOD_LONG}', f'05 Павловский МО расчёт за {PERIOD_SHOT}',
            'Розн. продажи MaCar за 2 года', 'Компания MaCar МО', 'Компания MaCar МО из 1С техническое'
        ]
    df = df[sort_list]
    return df


def payment(df_payment: pd.DataFrame, period: str) -> pd.DataFrame:
    """Рассчитываем значения МО расчет с учётом Кратности и Комплектности"""
    columns_payment = ['01 Кирова МО расчёт', '02 Автолюбитель МО расчёт',
                       '03 Интер МО расчёт', '04 Победа МО расчёт',
                       '08 Центр МО расчёт', '09 Вокзалка МО расчёт']

    if USE_OPT_STORE:
        columns_payment += ['05 Павловский МО расчёт']

    # Добавляем колонки сроков анализа и запаса
    df_payment[f'Кол-во дней анализа за {period}'] = data_day(period)
    df_payment[f'Кол-во дней запаса на {period}'] = DAY_STOCK

    for i in range(len(columns_payment)):
        # Суммируем продажи по всем складам подразделения
        df_payment[columns_payment[i][:-10] + f' продажи за {period}'] = \
            df_payment.filter(like=columns_payment[i][:-10]).sum(axis=1)

        # Рассчитываем значения исходя из периода анализируемых продажи и периода запасов
        # df_payment[columns_payment[i] + f' за {period}'] = \
        #     df_payment[columns_payment[i][:-10] + f' продажи за {period}'].fillna(0).apply(calc, period=period)
        # Обнуляем все не определённые значения
        df_payment[columns_payment[i] + f' за {period}'] = \
            df_payment[columns_payment[i][:-10] + f' продажи за {period}'].fillna(0)

        # Выполняем расчёт значения МО исходя из дней анализа и запаса
        df_payment[columns_payment[i] + f' за {period}'] = \
            df_payment[columns_payment[i] + f' за {period}'] / \
            df_payment[f'Кол-во дней анализа за {period}'] * df_payment[f'Кол-во дней запаса на {period}']

        # Обнуляем все рассчитанные МО меньше 0
        mask_lz = df_payment[columns_payment[i] + f' за {period}'] < 0
        df_payment.loc[mask_lz, columns_payment[i] + f' за {period}'] = 0

    return df_payment


def sum_retail_sales(df, name):
    """Суммируем продажи по всем розничным складам"""
    columns_sales = '01 Кирова|02 Автолюбитель|03 Интер|04 Победа|08 Центр|09 Вокзалка'
    df[name] = df.filter(regex=columns_sales).sum(axis=1)
    return df


def data_day(period):
    """Возвращаем количество дней анализа и запаса в зависимости от периода анализа
    :param: period -> Длинный период или другое
    :return: day_sales -> Количество дней анализа
    """
    if period == PERIOD_LONG:
        day_sales = YEAR_DAY_SALES
    else:
        day_sales = MONTH_DAY_SALES
    return day_sales


def set_comp_crat() -> pd.DataFrame:
    """Считываем данные из файла Комплектность, Кратность.xlsx и устанавливаем значения Комплектности и Кратности"""
    print('Устанавливаем значения Комплектности и Кратности')

    file = search_file('Комплектность, Кратность.xlsx')
    df_temp = create_df(file)
    df_temp = df_temp.fillna(1)

    if CONST_CRAT:
        print(f'Устанавливаем значение кратности равное {CONST_CRAT}')
        df_temp['Кратность'] = CONST_CRAT
    else:
        # Ограничивает значение Кратности и Комплектности введённым пользователем максимальным значением
        df_temp['Кратность'][df_temp['Кратность'] >= MAX_CRAT] = MAX_CRAT

    df_temp['Комплектность'][df_temp['Комплектность'] >= MAX_COMP] = MAX_COMP
    return df_temp


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
            name_format, MO_format, data_format, MO_format_1c = format_custom(workbook)

        # Перезаписываем заголовок таблицы с их форматированием
        for col_num, value in enumerate(df.columns.values):
            wks1.write(0, col_num, value, header_format)

        # Форматируем таблицу.
        # Устанавливаем высоту ячеек по умолчанию
        wks1.set_default_row(12)
        # Устанавливаем высоту первой ячейки с заголовками без форматироания
        wks1.set_row(0, 40, None)
        # Устанавливаем ширину и форматирование колонки с Кодом
        wks1.set_column('A:A', 12, name_format)
        # Устанавливаем ширину и форматирование колонки с Номенклатурой
        wks1.set_column('B:B', 32, name_format)
        # Устанавливаем ширину и форматирование колонки с данными по складам
        wks1.set_column('C:BK', 6, data_format)

        # Делаем жирную рамку и форматирование колонок с Кратностью и Комплектностью
        wks1.set_column(2, 2, None, border_storage_format_left, {'level': 1, 'hidden': True})
        wks1.set_column(3, 3, None, border_storage_format_right, {'level': 1, 'hidden': True})

        # Форматируем таблицу с данными
        i = 4
        while i < col_end + 1:
            if USE_OPT_STORE:
                if i < 53:
                    wks1.set_column(i, i, None, border_storage_format_left)
                    wks1.set_column(i + 2, i + 2, None, MO_format)
                    wks1.set_column(i + 3, i + 5, None, data_format, {'level': 1, 'hidden': True})
                    wks1.set_column(i + 6, i + 6, None, border_storage_format_right)
                elif i == 53:
                    wks1.set_column(i, i, None, border_storage_format_left)
                    wks1.set_column(i + 3, i + 5, None, data_format, {'level': 1, 'hidden': True})
                    wks1.set_column(i + 6, i + 6, None, border_storage_format_right)
                else:
                    wks1.set_column(i + 1, i + 1, None, MO_format)
                    wks1.set_column(i + 2, i + 2, None, MO_format_1c)
                i += 7
            else:
                if i < 46:
                    wks1.set_column(i, i, None, border_storage_format_left)
                    wks1.set_column(i + 2, i + 2, None, MO_format)
                    wks1.set_column(i + 3, i + 5, None, data_format, {'level': 1, 'hidden': True})
                    wks1.set_column(i + 6, i + 6, None, border_storage_format_right)
                    i += 7
                elif i == 46:
                    wks1.set_column(i, i, None, MO_format)
                    wks1.set_column(i + 1, i + 1, None, border_storage_format_right)
                    i += 2
                elif i == 48:
                    wks1.set_column(i, i, None, border_storage_format_left)
                    wks1.set_column(i + 3, i + 5, None, data_format, {'level': 1, 'hidden': True})
                    wks1.set_column(i + 6, i + 6, None, border_storage_format_right)
                    i += 7
                else:
                    wks1.set_column(i + 1, i + 1, None, MO_format)
                    wks1.set_column(i + 2, i + 2, None, MO_format_1c)
                    i += 7

        # Подставляем формулу в колонку с МО по всей компании
        # TODO Пока не делаем. Договариваемся что должно входить в МО по компании
        f = 2
        while f - 1 <= row_end:
            if USE_OPT_STORE:
                wks1.write_formula(
                    f'BJ{f}', f'=IF(OR(BK{f}>=1,BK{f}=0),'
                              f'SUM(INT(G{f}),INT(N{f}),INT(U{f}),INT(AB{f}),INT(AI{f}),INT(AP{f}),INT(AW{f})),'
                              f'BK{f})'
                )
            else:
                wks1.write_formula(
                    f'BE{f}', f'=IF(OR(BF{f}>=1,BF{f}=0),'
                              f'SUM(INT(G{f}),INT(N{f}),INT(U{f}),INT(AB{f}),INT(AI{f}),INT(AP{f}),INT(AU{f})),'
                              f'BF{f})'
                )
            f += 1

        # Добавляем выделение цветом строки при МО=0 по всей компании
        if USE_OPT_STORE:
            wks1.conditional_format(
                f'A2:BK{row_end_str}', {
                    'type': 'formula',
                    'criteria': '=AND($BJ2=0,$BI2>=1)',
                    'format': con_format
                }
            )
        else:
            wks1.conditional_format(
                f'A2:BF{row_end_str}', {
                    'type': 'formula',
                    'criteria': '=AND($BE2=0,$BD2>=1)',
                    'format': con_format
                }
            )

        # Добавление отображение итогов группировок сверху
        wks1.outline_settings(True, False, False, False)

        # Добавляем фильтр в первую колонку
        wks1.autofilter(0, 0, row_end + 1, col_end + 1)

        # Скрываем последнюю колонку
        # wks1.set_column(col_end + 1, col_end + 1, None, data_format, {'hidden': 1})
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

    mo_format = workbook.add_format({
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
    mo_format_1c = workbook.add_format({
        'num_format': '# ### ##0.00;;',
        'bold': False,
        'font_name': 'Arial',
        'font_size': '8',
        'font_color': '#FF0000',
        'border': True,
        'border_color': '#CCC085'
    })

    return header_format, con_format, border_storage_format_left, border_storage_format_right, \
        name_format, mo_format, data_format, mo_format_1c


def final_calc(row):
    """
    На основании данных продаж из колонки "Розн. продажи MaCar" корректируем значения в колонках "... МО расчёт"
    Так же на основании этих данных добавляем колонку "05 Павловский МО расчёт" и подставляем значения.
    Добавляем колонку "Компания MaCar МО расчёт", которая является суммой целых чисел всех колонок с
    наименованием "... МО расчёт"

    :param row: Series pandas
    :return: Series pandas
    """

    # Сортируем список складов согласно их среднего значения
    list_wh_sorted = sorted(LIST_WH, key=lambda x: row[f'{x} МО среднее значение'], reverse=True)
    # Составляем список наименования колонок МО Расчёт
    # list_col = [wh + ' МО расчёт' for wh in list_wh_sorted if wh not in WH_SKIP_LIST]
    # list_mo = [wh + ' МО' for wh in LIST_WH]

    prior_wh = PRIOR_WH
    list_col = []
    skip_list_col = []
    for wh in list_wh_sorted:
        not_skip_mo = round(row[wh + ' МО'] - int(row[wh + ' МО']), 2) not in VALUE_MO_SKIP
        not_skip_wh = wh not in WH_SKIP_LIST
        if not_skip_mo and not_skip_wh:
            list_col += [wh + ' МО расчёт']
        else:
            skip_list_col += [wh]
            row[wh + ' МО расчёт'] = row[wh + ' МО']

    # Определяем переменную для списка складов с установленными первыми значениями
    list_col_set = []

    # Определяем количество для установки в каждом проходе.
    # Определяем кол-во для установки за один проход
    # value_default = row['Кратность']
    value_default = math.ceil(row['Комплектность'] / row['Кратность']) * row['Кратность']
    val_def_crat = row['Кратность']
    # Определяем кол-во проходов согласно расчётного значения по компании.
    # Определяем количество складов для установки первого значения
    count_pass = int(row['Розн. MaCar МО расчёт (окр.)'] // value_default)
    # Определяем количество установок Кратности на установленные склады
    count_pass_crat = int((row['Розн. MaCar МО расчёт (окр.)'] % value_default) / val_def_crat)

    if count_pass > len(list_col):
        count_pass = len(list_col)
        count_pass_crat = (row['Розн. MaCar МО расчёт (окр.)'] - count_pass * value_default) / val_def_crat

    # Определяем кол-во возможных установок согласно значениям по каждому складу
    count_set = 0
    count_set_crat = 0
    for i in list_col:
        row[i + ' тех'] = 0
        if i in list_col[:count_pass] and row[i] / value_default >= 1:
            count_set += 1
            count_set_crat += (row[i] - value_default) / val_def_crat

    # Приравниваем кол-во проходов к кол-во возможных установок, если кол-во возможных установок меньше
    count_pass = min(count_pass, count_set)
    count_pass_crat = min(count_pass_crat, count_set_crat)
    # if count_pass > count_set:
    #     count_pass = count_set

    # Определяем наименование колонок для приоритетного склада в случае его использования
    prior_wh_mo, prior_wh_calc = None, None
    if prior_wh and prior_wh not in skip_list_col:
        prior_wh_mo = prior_wh + ' МО'
        prior_wh_calc = prior_wh + ' МО расчёт'
    else:
        prior_wh = False

    # Если кол-во проходов равно 0, то устанавливаем 0,8 если не было продаж за 2 года.
    # Если продажи за два года были, то пытаемся установить value_default только складах, где МО >= 1, но не более 2-х
    if count_pass == 0:
        if row['Розн. продажи MaCar за 2 года'] <= 0:
            for i in list_col:
                row[i] = set_value_mo_0_9(row[i[:-7]], 0.8)
        else:
            # Считаем кол-во установленных МО >= 1, но не больше 2
            count_pass = min(sum(row[i[:-7]] >= 1 for i in list_col), 2)

            # if PRIOR_WH and (row[prior_wh_mo] == 0 or row[prior_wh_mo] >= 1: # Если надо устанавливать МО даже при 0
            if prior_wh and row[prior_wh_mo] >= 1:
                # Устанавливаем ... МО расчёт согласно приоритетного склада.
                # Переставляем приоритетный склад вперёд списка и переворачиваем остальной список
                list_col1 = [prior_wh_calc] + [col for col in reversed(list_col) if col != prior_wh_calc]
                list_col2 = list_col.copy()
                while count_pass > 0:
                    for i1 in list_col1:
                        if row[i1[:-7]] >= 1 and count_pass > 0:
                            row[prior_wh_calc + ' тех'] += value_default
                            list_col_set += [prior_wh_calc]
                            count_pass -= 1
                            list_col2.remove(i1)
                            if i1 != prior_wh_calc:
                                list_col2.remove(prior_wh_calc)
                            break

                    for i2 in list_col2:
                        if row[i2[:-7]] >= 1 and count_pass > 0:
                            row[i2 + ' тех'] += value_default
                            list_col_set += [i2]
                            count_pass -= 1
            else:
                while count_pass > 0:
                    for i in list_col:
                        if row[i[:-7]] >= 1 and count_pass > 0:
                            row[i + ' тех'] += value_default
                            list_col_set += [i]
                            count_pass -= 1

            # Устанавливаем МО согласно расчётному значению
            set_mo_value(row, list_col, count_pass=0)

    # Если кол-во проходов равно 1, то:
    # Если есть приоритетный склад и его МО позволяет установить МО, то устанавливаем;
    # Если приоритетного склада нет, то устанавливаем согласно расчётному значению
    elif count_pass == 1:
        for i in list_col:
            while count_pass > 0:
                if prior_wh and (row[prior_wh_mo] == 0 or row[prior_wh_mo] >= 0.9):
                    # Устанавливаем ... МО расчёт согласно приоритетного склада
                    row[prior_wh_calc + ' тех'] += value_default
                    list_col_set += [prior_wh_calc]
                    count_pass -= 1
                else:
                    if row[i] >= 1:
                        row[i + ' тех'] += value_default
                        list_col_set += [i]
                        count_pass -= 1

        # Добавляем значения Кратности к установленным значениям
        if count_pass_crat:
            set_def_val_crat(row, list_col_set, count_pass_crat, val_def_crat)

        set_mo_value(row, list_col, prior_wh_calc)

    # Если кол-во проходов равно 2, то:
    # Если есть приоритетный склад и его МО позволяет установить МО,
    # то убираем расчётное значение с одного склада, где были продажи и устанавливаем на приоритетный склад;
    # Если приоритетного склада нет, то устанавливаем согласно расчётному значению
    elif count_pass == 2:

        if prior_wh and (row[prior_wh_mo] == 0 or row[prior_wh_mo] >= 0.9):
            # Устанавливаем ... МО расчёт согласно приоритетного склада.
            # Переставляем приоритетный склад вперёд списка и переворачиваем остальной список
            list_col1 = [prior_wh_calc] + [col for col in reversed(list_col) if col != prior_wh_calc]
            list_col2 = list_col.copy()
            first_pass = True
            for i1 in list_col1:
                if row[i1] >= 1 and row[i1] > row[i1 + ' тех'] and count_pass > 0 and first_pass:
                    row[prior_wh_calc + ' тех'] += value_default
                    list_col_set += [prior_wh_calc]
                    first_pass = False
                    count_pass -= 1
                    list_col2.remove(i1)
                    if i1 != prior_wh_calc:
                        list_col2.remove(prior_wh_calc)

            for i2 in list_col2:
                if row[i2] >= 1 and row[i2] > row[i2 + ' тех'] and count_pass > 0:
                    row[i2 + ' тех'] += value_default
                    list_col_set += [i2]
                    count_pass -= 1

        else:
            while count_pass > 0:
                for i in list_col:
                    if row[i] >= 1 and row[i] > row[i + ' тех'] and count_pass > 0:
                        row[i + ' тех'] += value_default
                        list_col_set += [i]
                        count_pass -= 1

        if count_pass_crat:
            set_def_val_crat(row, list_col_set, count_pass_crat, val_def_crat)

        set_mo_value(row, list_col, prior_wh_calc)

    # Если кол-во проходов больше 2, то устанавливаем согласно расчётному значению
    elif count_pass > 2:
        while count_pass > 0:
            for i in list_col:
                if row[i] >= 1 and row[i] > row[i + ' тех'] and count_pass > 0:
                    row[i + ' тех'] += value_default
                    list_col_set += [i]
                    count_pass -= 1

        # Добавляем значения Кратности к установленным значениям
        if count_pass_crat:
            set_def_val_crat(row, list_col_set, count_pass_crat, val_def_crat)

        set_mo_value(row, list_col, prior_wh_calc)

    # На всякий случай выводим сообщение, если расчётное значение не устанавливалось
    else:
        print('Не попадает в условия простановки МО. Сообщите разработчику!!!')
        print(row)

    # Обнуляем расчётные значения оптового склада, если их не надо учитывать
    if not USE_OPT_STORE:
        row['05 Павловский МО расчёт'] = 0

    if DAY_OPT_STOCK > DAY_STOCK:
        # Получаем значение для установки дополнительного запаса на Оптовом складе
        val_all_sales = row[ADD_NAME_PERIOD_OPT_STOCK]

        if 0 < val_all_sales < 1:
            val_all_sales = 1
        else:
            val_all_sales = round(val_all_sales)

        if val_all_sales >= MIN_VALUE_ADD_OPT_SALES:
            # Определяем необходимое количество сверх запасов розничных складов и устанавливаем значение
            val_opt = val_all_sales * (DAY_OPT_STOCK / DAY_STOCK)
            val_opt = val_opt - val_all_sales

            # Добавляем полученное значения дополнительного запаса к рассчитанному значению Оптового склада
            row['05 Павловский МО расчёт'] = row['05 Павловский МО расчёт'] + val_opt if val_opt > 0 else 0

    # Устанавливаем расчётные значения оптового склада согласно правилам и МО на этом складе
    if row['05 Павловский МО расчёт'] == 0:
        row['05 Павловский МО расчёт'] = set_value_mo_0_9(row['05 Павловский МО'], 0.96)
    # Приводим рассчитанные значение в колонке '05 Павловский МО расчёт' к значению Кратности или Комплектности
    elif (row['05 Павловский МО расчёт'] > 0) & (row['05 Павловский МО расчёт'] < row['Кратность']):
        row['05 Павловский МО расчёт'] = set_value_mo_0_9(row['05 Павловский МО'], row['Кратность'])

    # # Определяем значение Комплектности с учётом Кратности TODO Убираем расчёт комплектности на оптовом складе
    # comp_into_crat = math.ceil(row['Комплектность'] / row['Кратность']) * row['Кратность']
    # if (row['05 Павловский МО расчёт'] > 0) & (row['05 Павловский МО расчёт'] < comp_into_crat):
    #     row['05 Павловский МО расчёт'] = set_value_mo_0_9(row['05 Павловский МО'], comp_into_crat)

    else:  # row['05 Павловский МО расчёт'] > row['Комплектность']:
        opt_val = round(
            (row['05 Павловский МО расчёт'] + 0.000001) / row['Кратность']
        ) * row['Кратность']
        row['05 Павловский МО расчёт'] = set_value_mo_0_9(row['05 Павловский МО'], opt_val)

    if row['Розн. MaCar МО расчёт (окр.)'] / value_default <= 0:
        if row['Розн. продажи MaCar за 2 года'] <= 0:
            row['05 Павловский МО расчёт'] = set_value_mo_0_9(row['05 Павловский МО'], 0.8)

    row['Компания MaCar МО расчёт'] = 0
    for i in list_col + ['05 Павловский МО расчёт']:
        row['Компания MaCar МО расчёт'] += int(row[i])
    if row['Розн. MaCar МО расчёт (окр.)'] / value_default == 0:
        if row['Розн. продажи MaCar за 2 года'] <= 0:
            row['Компания MaCar МО расчёт'] = 0.8
        else:
            if row['Компания MaCar МО расчёт'] == 0:
                row['Компания MaCar МО расчёт'] = 0.91

    row['Компания MaCar МО расчёт'] = set_value_mo_0_9(row['Компания MaCar МО из 1С техническое'],
                                                       row['Компания MaCar МО расчёт'])
    return row


def set_value_mo_0_9(mo_old: float, value: float) -> float:
    """Если mo_old от 0 до 0.9 не включительно, то возвращаем его. В противном случае возвращаем значение value"""
    return mo_old if 0 < mo_old < 0.9 else value


def set_def_val_crat(row: pd.Series, list_col: list, count_pass_crat: float, value: int) -> pd.Series:
    """Добавляем значение кратности к расчётному значению
    :param row: Строка pd.Series с данными по продажам, МО и т.д.
    :param list_col: Список колонок для добавления кратности
    :param count_pass_crat: Количество проходов по кратности
    :param value: Значение кратности
    :return: Строка pd.Series с данными по продажам, МО и т.д.
    """
    while count_pass_crat > 0:
        for i in list_col:
            if count_pass_crat > 0 and row[i] > row[i + ' тех']:
                row[i + ' тех'] += value
                count_pass_crat -= 1
    return row


def set_mo_value(row: pd.Series, list_col: list, prior_wh_calc: str = '', count_pass: int = 1) -> pd.Series:
    """
    Устанавливаем расчётные значения оптового склада согласно правилам и МО на этом складе
    :param row: Строка pd.Series с данными по продажам, МО и т.д.
    :param list_col: Список колонок для установки МО
    :param prior_wh_calc: Приоритетный склад для установки МО
    :param count_pass: Выбор режима установки в зависимости от Количество проходов по МО
    :return: Строка pd.Series с данными по продажам, МО и т.д.
    """
    if count_pass:
        # Устанавливаем значения на основании Расчётного значения и приоритетного склада
        for i in list_col:
            if i == prior_wh_calc and row[i + ' тех'] >= 1:
                row[i] = set_value_mo_0_9(row[i[:-7]], row[i + ' тех'])
            elif row[i] >= 1 and row[i + ' тех'] >= 1:
                row[i] = set_value_mo_0_9(row[i[:-7]], row[i + ' тех'])
            elif row[i] >= 1 and row[i + ' тех'] == 0:
                row[i] = set_value_mo_0_9(row[i[:-7]], 0.93)
            elif row[i] == 0:
                row[i] = set_value_mo_0_9(row[i[:-7]], 0.92)
    else:
        # Устанавливаем значения на основании значения МО
        for i in list_col:
            if row[i[:-7]] >= 1 and row[i + ' тех'] >= 1:
                row[i] = set_value_mo_0_9(row[i[:-7]], row[i + ' тех'])
            elif row[i[:-7]] >= 1 and row[i + ' тех'] >= 0:
                row[i] = set_value_mo_0_9(row[i[:-7]], 0.93)
            else:
                row[i] = set_value_mo_0_9(row[i[:-7]], 0.91)

    return row


# def set_value_mo_0_1(mo_old: float, value: float) -> float:
#     """Если mo_old от 0 до 1, то возвращаем его. Если 0, то возвращаем value.
#     В противном случае возвращаем 1"""
#     if mo_old == 0 or mo_old >= 1:
#         mo_new = value
#     else:
#         mo_new = mo_old
#
#     return mo_new


def total_value_calc(row: pd.Series) -> pd.Series:
    """Рассчитываем средние значения по каждому складу, общее значения Розн. продажи MaCar за длинный период,
    вес склада, Розн. продажи MaCar за короткий период, Розн. продажи MaCar среднее значение, Розн. расчёт MaCar.
    :param row: Строка с данными по продажам, МО и расчётным значениям.
    :return: Строка row с добавленными колонками.
    """
    list_sales_long = [wh + f' продажи за {PERIOD_LONG}' for wh in LIST_WH]
    list_sales_shot = [wh + f' продажи за {PERIOD_SHOT}' for wh in LIST_WH]
    list_mo_retail = [wh + ' МО' for wh in LIST_WH]
    list_mo = [wh + ' МО' for wh in LIST_WH]
    list_calc_long = [wh + f' МО расчёт за {PERIOD_LONG}' for wh in LIST_WH]
    list_calc_shot = [wh + f' МО расчёт за {PERIOD_SHOT}' for wh in LIST_WH]
    list_calc_merge = [wh + f' МО среднее значение' for wh in LIST_WH]
    list_calc = [wh + ' МО расчёт' for wh in LIST_WH]
    day_sales_long = data_day(PERIOD_LONG)
    day_sales_shot = data_day(PERIOD_SHOT)

    # Добавляем оптовый склад в списки при его использовании
    if USE_OPT_STORE:
        list_mo.append('05 Павловский МО')
        list_calc_long.append(f'05 Павловский МО расчёт за {PERIOD_LONG}')
        list_calc_shot.append(f'05 Павловский МО расчёт за {PERIOD_SHOT}')
        list_calc_merge.append(f'05 Павловский МО среднее значение')
        list_calc.append('05 Павловский МО расчёт')

    # Суммируем данные из колонок list_sales... в отдельные колонки, если МО == 0 или МО >= 0.9
    row[f'Розн. продажи MaCar за {PERIOD_LONG}'] = 0
    row[f'Розн. продажи MaCar за {PERIOD_SHOT}'] = 0
    row[f'Розн. MaCar МО расчёт за {PERIOD_LONG}'] = 0
    row[f'Розн. MaCar МО расчёт за {PERIOD_SHOT}'] = 0

    # Определяем значение Комплектности с учётом Кратности
    comp_into_crat = math.ceil(row['Комплектность'] / row['Кратность']) * row['Кратность']

    # Суммируем значения продаж за длинный и короткий периоды
    for key, item in enumerate(list_mo_retail):
        not_skip_mo = round(row[item] - int(row[item]), 2) not in VALUE_MO_SKIP
        not_skip_wh = LIST_WH[key] not in WH_SKIP_LIST
        # Суммируем, если 0 < МО <= 0.9, значение МО после запятой не надо пропускать и склад так же не надо пропускать
        if (row[item] == 0) or (row[item] >= 0.9):  # and not_skip_mo and not_skip_wh:
            row[f'Розн. продажи MaCar за {PERIOD_LONG}'] += row[list_sales_long[key]]
            row[f'Розн. продажи MaCar за {PERIOD_SHOT}'] += row[list_sales_shot[key]]

    # Рассчитываем среднее значение по каждому складу
    for key, item in enumerate(list_mo):
        not_skip_mo = round(row[item] - int(row[item]), 2) not in VALUE_MO_SKIP

        # Проверяем, что key в пределах LIST_WH
        """
        Когда учитываются продажи Оптового склада то он есть в списке list_mo, но его нет в списке LIST_WH.
        Чтобы не менять список LIST_WH при работе с Оптовым складом, просто возвращаем по последнему значению True,
        т.к. оно не может быть в списке исключённых
        
        Старая строка кода
        not_skip_wh = LIST_WH[key] not in WH_SKIP_LIST
        """
        if key >= len(LIST_WH):
            not_skip_wh = True  # или break, если нужно прервать выполнение
        else:
            not_skip_wh = LIST_WH[key] not in WH_SKIP_LIST

        if (row[item] == 0) or (row[item] >= 0.9) and not_skip_mo and not_skip_wh:
            row[list_calc_merge[key]] = \
                (row[list_calc_long[key]] + row[list_calc_shot[key]]) / 2
        else:
            row[list_calc_merge[key]] = 0

        # Рассчитываем расчётное значение МО на каждый склад с учетом кратности и комплектности
        row[list_calc[key]] = row[list_calc_merge[key]]
        if row[list_calc[key]] < 0:
            row[list_calc[key]] = 0
        if 0 < row[list_calc[key]] <= row['Кратность']:
            row[list_calc[key]] = row['Кратность']

        # Приравниванием к значению Комплектности с учётом Кратности
        if 0 < row[list_calc[key]] <= comp_into_crat:
            row[list_calc[key]] = comp_into_crat
        else:
            # Все расчётные значения больше комплектности устанавливаем согласно кратности
            # с округлением в меньшую сторону
            row[list_calc[key]] = \
                round(row[list_calc[key]] / row['Кратность']) * row['Кратность']

    # Рассчитываем значение Розн. MaCar МО расчёт за длинный период
    row[f'Розн. MaCar МО расчёт за {PERIOD_LONG}'] = \
        row[f'Розн. продажи MaCar за {PERIOD_LONG}'] / day_sales_long * DAY_STOCK
    row[f'Розн. MaCar МО расчёт за {PERIOD_LONG}'] = \
        row[f'Розн. MaCar МО расчёт за {PERIOD_LONG}'] if row[f'Розн. MaCar МО расчёт за {PERIOD_LONG}'] >= 0 else 0

    # Рассчитываем значение Розн. MaCar МО расчёт за короткий период
    row[f'Розн. MaCar МО расчёт за {PERIOD_SHOT}'] = \
        row[f'Розн. продажи MaCar за {PERIOD_SHOT}'] / day_sales_shot * DAY_STOCK
    row[f'Розн. MaCar МО расчёт за {PERIOD_SHOT}'] = \
        row[f'Розн. MaCar МО расчёт за {PERIOD_SHOT}'] if row[f'Розн. MaCar МО расчёт за {PERIOD_SHOT}'] >= 0 else 0

    # Рассчитываем среднее значение Розн. MaCar МО расчёт
    row[f'Розн. MaCar МО среднее значение'] = \
        (row[f'Розн. MaCar МО расчёт за {PERIOD_LONG}'] + row[f'Розн. MaCar МО расчёт за {PERIOD_SHOT}']) / 2

    # Приравниваем все расчетные значения меньше нуля к нулю
    if row['Розн. MaCar МО среднее значение'] < 0:
        row['Розн. MaCar МО среднее значение'] = 0

    # Округляем все остальные значения в большую сторону
    else:
        row['Розн. MaCar МО расчёт (окр.)'] = math.ceil(row['Розн. MaCar МО среднее значение'])

    # Если расчётное значение больше 0, но меньше Кратности, то приводим к значению Кратности
    if 0 < row['Розн. MaCar МО расчёт (окр.)'] <= row['Кратность']:
        row['Розн. MaCar МО расчёт (окр.)'] = row['Кратность']

    # Если расчётное значение больше 0, но меньше Комплектности, то приводим к значению Комплектности с учётом Кратности
    if 0 < row['Розн. MaCar МО расчёт (окр.)'] <= comp_into_crat:
        row['Розн. MaCar МО расчёт (окр.)'] = comp_into_crat
    # Все значения больше Комплектности приводим к значению Кратности с округлением в большую сторону
    elif row['Розн. MaCar МО расчёт (окр.)'] > comp_into_crat:
        row['Розн. MaCar МО расчёт (окр.)'] = math.ceil(
            row['Розн. MaCar МО расчёт (окр.)'] / row['Кратность']
        ) * row['Кратность']

    return row


def create_final_df(df1: pd.DataFrame, df2: pd.DataFrame, df3: pd.DataFrame, df4: pd.DataFrame) -> pd.DataFrame:
    """Создаём финальный DataFrame
    :param df1:
    :param df2:
    :param df3:
    :param df4:
    :return :
    """
    df_final = concat_df(df1, df2)
    df_final = concat_df(df_final, df3)
    df_final = concat_df(df_final, df4)
    df_comp_crat = set_comp_crat()
    # df_final.drop(columns=['Комплектность', 'Кратность'], inplace=True)
    df_final = concat_df(df_final, df_comp_crat)
    df_final['Компания MaCar МО из 1С техническое'] = df_final['Компания MaCar МО']
    # df_final[['Комплектность', 'Кратность']] = df_final[['Комплектность', 'Кратность']].fillna(1)
    df_final = sort_df(df_final)  # Оставляем только необходимые столбцы в нужном порядке
    return df_final.fillna(0)


def sort_df_final(df):
    """Сортируем столбцы в нужном порядке"""
    sort_list = ['Комплектность', 'Кратность',
                 f'01 Кирова продажи за {PERIOD_LONG}', f'01 Кирова продажи за {PERIOD_SHOT}', '01 Кирова МО',
                 f'01 Кирова МО расчёт за {PERIOD_LONG}', f'01 Кирова МО расчёт за {PERIOD_SHOT}',
                 f'01 Кирова МО среднее значение', f'01 Кирова МО расчёт',
                 f'02 Автолюбитель продажи за {PERIOD_LONG}', f'02 Автолюбитель продажи за {PERIOD_SHOT}',
                 '02 Автолюбитель МО',
                 f'02 Автолюбитель МО расчёт за {PERIOD_LONG}', f'02 Автолюбитель МО расчёт за {PERIOD_SHOT}',
                 f'02 Автолюбитель МО среднее значение', f'02 Автолюбитель МО расчёт',
                 f'03 Интер продажи за {PERIOD_LONG}', f'03 Интер продажи за {PERIOD_SHOT}', '03 Интер МО',
                 f'03 Интер МО расчёт за {PERIOD_LONG}', f'03 Интер МО расчёт за {PERIOD_SHOT}',
                 f'03 Интер МО среднее значение', f'03 Интер МО расчёт',
                 f'04 Победа продажи за {PERIOD_LONG}', f'04 Победа продажи за {PERIOD_SHOT}', '04 Победа МО',
                 f'04 Победа МО расчёт за {PERIOD_LONG}', f'04 Победа МО расчёт за {PERIOD_SHOT}',
                 f'04 Победа МО среднее значение', f'04 Победа МО расчёт',
                 f'08 Центр продажи за {PERIOD_LONG}', f'08 Центр продажи за {PERIOD_SHOT}', '08 Центр МО',
                 f'08 Центр МО расчёт за {PERIOD_LONG}', f'08 Центр МО расчёт за {PERIOD_SHOT}',
                 f'08 Центр МО среднее значение', f'08 Центр МО расчёт',
                 f'09 Вокзалка продажи за {PERIOD_LONG}', f'09 Вокзалка продажи за {PERIOD_SHOT}', '09 Вокзалка МО',
                 f'09 Вокзалка МО расчёт за {PERIOD_LONG}', f'09 Вокзалка МО расчёт за {PERIOD_SHOT}',
                 f'09 Вокзалка МО среднее значение', f'09 Вокзалка МО расчёт',
                 '05 Павловский МО', '05 Павловский МО расчёт',
                 'Розн. продажи MaCar за 2 года',
                 f'Розн. продажи MaCar за {PERIOD_LONG}', f'Розн. продажи MaCar за {PERIOD_SHOT}',
                 f'Розн. MaCar МО расчёт за {PERIOD_LONG}', f'Розн. MaCar МО расчёт за {PERIOD_SHOT}',
                 'Розн. MaCar МО среднее значение', 'Розн. MaCar МО расчёт (окр.)',
                 'Компания MaCar МО расчёт', 'Компания MaCar МО', 'Компания MaCar МО из 1С техническое']
    if USE_OPT_STORE:
        sort_list = ['Комплектность', 'Кратность',
                     f'01 Кирова продажи за {PERIOD_LONG}', f'01 Кирова продажи за {PERIOD_SHOT}', '01 Кирова МО',
                     f'01 Кирова МО расчёт за {PERIOD_LONG}', f'01 Кирова МО расчёт за {PERIOD_SHOT}',
                     f'01 Кирова МО среднее значение', f'01 Кирова МО расчёт',
                     f'02 Автолюбитель продажи за {PERIOD_LONG}', f'02 Автолюбитель продажи за {PERIOD_SHOT}',
                     '02 Автолюбитель МО',
                     f'02 Автолюбитель МО расчёт за {PERIOD_LONG}', f'02 Автолюбитель МО расчёт за {PERIOD_SHOT}',
                     f'02 Автолюбитель МО среднее значение', f'02 Автолюбитель МО расчёт',
                     f'03 Интер продажи за {PERIOD_LONG}', f'03 Интер продажи за {PERIOD_SHOT}', '03 Интер МО',
                     f'03 Интер МО расчёт за {PERIOD_LONG}', f'03 Интер МО расчёт за {PERIOD_SHOT}',
                     f'03 Интер МО среднее значение', f'03 Интер МО расчёт',
                     f'04 Победа продажи за {PERIOD_LONG}', f'04 Победа продажи за {PERIOD_SHOT}', '04 Победа МО',
                     f'04 Победа МО расчёт за {PERIOD_LONG}', f'04 Победа МО расчёт за {PERIOD_SHOT}',
                     f'04 Победа МО среднее значение', f'04 Победа МО расчёт',
                     f'08 Центр продажи за {PERIOD_LONG}', f'08 Центр продажи за {PERIOD_SHOT}', '08 Центр МО',
                     f'08 Центр МО расчёт за {PERIOD_LONG}', f'08 Центр МО расчёт за {PERIOD_SHOT}',
                     f'08 Центр МО среднее значение', f'08 Центр МО расчёт',
                     f'09 Вокзалка продажи за {PERIOD_LONG}', f'09 Вокзалка продажи за {PERIOD_SHOT}', '09 Вокзалка МО',
                     f'09 Вокзалка МО расчёт за {PERIOD_LONG}', f'09 Вокзалка МО расчёт за {PERIOD_SHOT}',
                     f'09 Вокзалка МО среднее значение', f'09 Вокзалка МО расчёт',
                     f'05 Павловский продажи за {PERIOD_LONG}', f'05 Павловский продажи за {PERIOD_SHOT}',
                     '05 Павловский МО',
                     f'05 Павловский МО расчёт за {PERIOD_LONG}', f'05 Павловский МО расчёт за {PERIOD_SHOT}',
                     f'05 Павловский МО среднее значение', f'05 Павловский МО расчёт',
                     'Розн. продажи MaCar за 2 года',
                     f'Розн. продажи MaCar за {PERIOD_LONG}', f'Розн. продажи MaCar за {PERIOD_SHOT}',
                     f'Розн. MaCar МО расчёт за {PERIOD_LONG}', f'Розн. MaCar МО расчёт за {PERIOD_SHOT}',
                     'Розн. MaCar МО среднее значение', 'Розн. MaCar МО расчёт (окр.)',
                     'Компания MaCar МО расчёт', 'Компания MaCar МО', 'Компания MaCar МО из 1С техническое']
    df = df[sort_list]
    return df


def run():
    sales_filelist = search_file(salesName)  # Запускаем метод по поиску файлов и получаем список файлов
    minstock_filelist = search_file(minStockName)  # Запускаем метод по поиску файлов и получаем список файлов

    # Разделяем файлы на три списка: один год, два года и три месяца
    sales_filelist_one_year, sales_filelist_two_year, sales_filelist_three_month = \
        create_filelist_one_tow_year(sales_filelist)

    # Создаём DataFrame с продажами за 2 года
    df_sales_two_year = create_df(sales_filelist_two_year, salesName)
    # Суммируем продажи за 2 года
    print(f'Суммируем продажи за 2 года...')
    df_sales_two_year = sum_retail_sales(df_sales_two_year, name='Розн. продажи MaCar за 2 года')
    df_sales_two_year = df_sales_two_year['Розн. продажи MaCar за 2 года']

    # Создаём DataFrame с продажами за длинный период
    df_sales_one_year = create_df(sales_filelist_one_year, salesName)
    # Рассчитываем значения МО расчет с учётом Кратности и Комплектности
    print(f'Рассчитываем значения МО расчет с учётом Кратности и Комплектности за {PERIOD_LONG}...')
    df_sales_one_year = payment(df_sales_one_year, PERIOD_LONG)

    # Создаём DataFrame с продажами за короткий период
    df_sales_three_month = create_df(sales_filelist_three_month, salesName)
    # Рассчитываем значения МО расчет с учётом Кратности и Комплектности
    print(f'Рассчитываем значения МО расчет с учётом Кратности и Комплектности за {PERIOD_SHOT}...')
    df_sales_three_month = payment(df_sales_three_month, PERIOD_SHOT)

    # Создаём DataFrame с минимальными остатками
    print('Создаём DataFrame с минимальными остатками...')
    df_minstock = create_df(minstock_filelist, minStockName)

    # Объединяем полученные данные в единый DataFrame
    print('Объединяем полученные данные в единый DataFrame...')
    df_final = create_final_df(df_minstock, df_sales_one_year, df_sales_two_year, df_sales_three_month)

    # Рассчитываем итоговые значения по компании
    print('Рассчитываем итоговые значения по компании...')
    df_final = df_final.apply(total_value_calc, axis=1)

    # Осуществляем последние расчёты по каждой строке
    print('Осуществляем последние расчёты по каждой строке...')
    df_final = df_final.apply(final_calc, axis=1)

    # Сортируем столбцы
    print('Сортируем столбцы...')
    df_final = sort_df_final(df_final)

    # Сохраняем конечный результат в эксель
    print('Сохраняем конечный результат в эксель...')
    df_write_xlsx(df_final)
    return


if __name__ == '__main__':
    LIST_WH = input_list_wh()  # Запрашиваем список складов
    print()
    WH_SKIP_LIST = input_wh_skip_list()  # Запрашиваем список не учитываемых складов
    print()
    MAX_CRAT, MAX_COMP, CONST_CRAT = input_max_crat_comp()  # Задаём максимальное значение Кратности и Комплектности
    print()
    PRIOR_WH = input_prior_wh()  # Задаём приоритетный склад
    print()
    # Задаём кол-во дней для расчётов Розничных складов за длинный период
    YEAR_DAY_SALES = input_day_sales(365, PERIOD_LONG)
    print()
    # Задаём кол-во дней для расчётов Розничных складов за короткий период
    MONTH_DAY_SALES = input_day_sales(91, PERIOD_SHOT)
    print()
    # Задаём кол-во дней запаса
    DAY_STOCK = input_day_stock()
    print()
    USE_OPT_STORE = input_use_opt_store()  # Задаём учёт продаж оптового склада для установки МО
    print()
    # Задаём кол-во дней запаса для оптового склада и минимальная сумма расчётных значений розничных складов
    DAY_OPT_STOCK, MIN_VALUE_ADD_OPT_SALES, ADD_NAME_PERIOD_OPT_STOCK = input_day_opt_stock()
    print()

    # Задаём значение продаж для применения всех приоритетных складов
    # VALUE_ALL_PRIOR_WH = input_value_all_prior_wh() TODO Пока решили не использовать
    # Задаём значение продаж и список дополнительных приоритетных складов
    # VALUE_ADD_PRIOR_WH, LIST_ADD_PRIOR_WH = input_value_add_prior_wh() TODO Пока решили не использовать
    # Запускаем программу
    run()
