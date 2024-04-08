def kotzdorov(file_name):
    # Получаем имя листа
    wb = load_workbook(file_name, data_only=True)
    sheet = wb.sheetnames[0]
    # Инициализируем книгу
    sheet_df_dictonary = pd.read_excel(file_name, engine='openpyxl', sheet_name=[sheet], skiprows=0, header=None)
    sheet_Name = sheet_df_dictonary[sheet]
    # Добавляем новую колонку
    sheet_Name.insert(6, 'new-column-twfs', 0)
    # Меняем названия прошлых колонок
    sheet_Name.rename(columns={1: 'ФИО плательщика',
                               3: 'Номер заказа у партнера',
                               4: 'Дата рождения плательщика',
                               'new-column-twfs': 'Пол плательщика',
                               6: 'Дата начала действия',
                               8: 'Продукт в системе партнера',
                               10: 'Номер телефона плательщика',
                               11: 'Адрес электронной почты плательщика',
                               12: 'Номер сертификата'}, inplace=True)
    # Меняем формат пола
    sheet_Name['Пол плательщика'] = IF(TYPE(sheet_Name[5]) != 'NaN',
                                       SUBSTITUTE(SUBSTITUTE(sheet_Name[5], 'Мужской', 'М'), 'Женский', 'Ж'), None)
    # Фильтруем, что-бы убрать лишние строки
    sheet_Name = sheet_Name[(sheet_Name[5].notnull()) & (~sheet_Name[5].str.contains('Пол', na=False))]
    # Оборачиваем все в словарь, для дальнейшей работы с ним
    df = pd.DataFrame.to_dict(sheet_Name)
    # Возвращаем отморфированные данные
    return df



def mts(file_name):
    # Получаем имя листа
    wb = load_workbook(file_name, data_only=True)
    sheet = wb.sheetnames[0]
    # Инициализируем книгу
    sheet_df_dictonary = pd.read_excel(file_name, engine='openpyxl', sheet_name=[sheet], skiprows=0, header=None)
    sheet_Name = sheet_df_dictonary[sheet]
    # Добавляем новый колонки
    sheet_Name.insert(10, 'new-column-4p5a', 0)
    sheet_Name.insert(12, 'new-column-vv6o', 0)
    # Меняем названия прошлых колонок
    sheet_Name.rename(columns={7: 'Продукт в системе партнера',
                               8: 'Номер сертификата',
                               'new-column-4p5a': 'Дата начала действия',
                               11: 'Стоимость',
                               'new-column-vv6o': 'ФИО плательщика'}, inplace=True)
    # Применяем формулы
    sheet_Name['ФИО плательщика'] = IF(TYPE(sheet_Name[10]) != 'NaN', PROPER(sheet_Name[10]), None)
    sheet_Name['Дата начала действия'] = IF(TYPE(sheet_Name[9]) != 'NaN',
                                            LEFT(sheet_Name[9], INT(FIND(sheet_Name[9], ' '))), None)
    # Фильтруем, что-бы убрать лишние строки
    sheet_Name = sheet_Name[
        (sheet_Name['ФИО плательщика'].notnull()) & (
            ~sheet_Name['ФИО плательщика'].str.contains('Покупатель', na=False))]
    # Меняем на нужный формат
    sheet_Name['Стоимость'] = to_float_series(sheet_Name['Стоимость'])
    # Оборачиваем все в словарь, для дальнейшей работы с ним
    df = pd.DataFrame.to_dict(sheet_Name)
    # Возвращаем отморфированные данные
    return df


def smp_stahovanie(file_name):
    # Получаем имя листа
    wb = load_workbook(file_name, data_only=True)
    sheet = wb.sheetnames[0]
    # Инициализируем книгу
    sheet_df_dictonary = pd.read_excel(file_name, engine='openpyxl', sheet_name=[sheet], skiprows=0, header=None)
    sheet_Name = sheet_df_dictonary[sheet]
    # Меняем названия прошлых колонок
    sheet_Name.rename(columns={0: 'Дата начала действия',
                               2: 'Номер сертификата',
                               3: 'Номер заказа у партнера'}, inplace=True)
    # Фильтруем, что-бы убрать лишние строки
    sheet_Name = sheet_Name[
        (sheet_Name['Номер сертификата'].notnull()) & (
            ~sheet_Name['Номер сертификата'].str.contains('Промокод', na=False))]
    # Оборачиваем все в словарь, для дальнейшей работы с ним
    df = pd.DataFrame.to_dict(sheet_Name)
    # Возвращаем отморфированные данные
    return df

def skb(file_name):
    # Получаем имя листа
    wb = load_workbook(file_name, data_only=True)
    sheet = wb.sheetnames[0]
    # Инициализируем книгу
    sheet_df_dictonary = pd.read_excel(file_name, engine='openpyxl', sheet_name=[sheet], skiprows=0, header=None)
    sheet_Name = sheet_df_dictonary[sheet]
    # Добавляем новый колонки
    sheet_Name.insert(4, 'new-column-8awm', 0)
    sheet_Name.insert(3, 'new-column-g96e', 0)
    # Меняем названия прошлых колонок
    sheet_Name.rename(columns={'new-column-g96e': 'Продукт в системе партнера',
                               'new-column-8awm': 'ФИО плательщика',
                               4: 'Номер телефона плательщика',
                               5: 'Номер сертификата',
                               6: 'Стоимость',
                               7: 'Дата начала действия',
                               8: 'Дата окончания действия',
                               9: 'Наименование ДО'}, inplace=True)
    # Применяем формулы
    sheet_Name['ФИО плательщика'] = IF(TYPE(sheet_Name[3]) != 'NaN', PROPER(sheet_Name[3]), None)
    sheet_Name['Продукт в системе партнера'] = IF(TYPE(sheet_Name[2]) != 'NaN', SUBSTITUTE(sheet_Name[2],
                                                                                           LEFT(sheet_Name[2],
                                                                                                INT(FIND(sheet_Name[2],
                                                                                                         '"') - 1)),
                                                                                           ''), None)
    # Фильтруем, что-бы убрать лишние строки
    sheet_Name = sheet_Name[
        (sheet_Name['ФИО плательщика'].notnull()) & (~sheet_Name['ФИО плательщика'].str.contains('Фио', na=False))]
    # Меняем на нужный формат
    sheet_Name['Стоимость'] = to_float_series(sheet_Name['Стоимость'])
    # Оборачиваем все в словарь, для дальнейшей работы с ним
    df = pd.DataFrame.to_dict(sheet_Name)
    # Возвращаем отморфированные данные
    return df