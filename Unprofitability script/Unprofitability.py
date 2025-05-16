import os
import numpy as np
import pandas as pd
import re #регулярны

class Unprofitability():
    error = ''

    def __init__(self, folder):
        self.folder = folder
        self.form()

    def form(self):
        if ('Убыточность.csv' not in os.listdir(self.folder)):
            self.error = f'В директории не хватает файлов: Убыточность.csv'
            raise Exception

        filename= self.folder + '/Убыточность.csv'
        df_una = pd.read_csv(filename, delimiter=';', encoding='1251', low_memory=False)

        columns = df_una.columns.tolist()
        required_col = ['Сумма начисленной премии руб', 'Сумма начисленной комиссии руб', 'Заработанная премия', 'Заработанная комиссия',
                        'Сумма заявленных убытков', 'Сумма урегулированных убытков', 'Выплаты по неустойкам', 'Сумма убытка или ЗНУ',
                        'Сумма убытка или ЗНУ без неустоек', 'ДВОУ', 'Расходы на урегулирование убытков', 'Расходы на урегулирование убытков (ОФР 25203)',
                        'Сумма начисления по регрессу', 'Сумма поступления по регрессу', 'Год изготовления', 'Номер полиса', 'Дата договора',
                        'Дата начала ответственности', 'Дата окончания ответственности', 'Дата начисления премии']

        if (sorted(list(set(required_col))) != sorted(list(set(required_col) & set(columns)))):
            differences = list(set(required_col) - set(columns))
            append_or_crush = ", ".join(str(element) for element in differences)
            self.error = f'В таблице Убыточность не хватает столбцов:\n{append_or_crush}'
            raise Exception

        num_column = ['Сумма начисленной премии руб', 'Сумма начисленной комиссии руб', 'Сумма расторжения',
        'Заработанная премия', 'Заработанная комиссия', 'Сумма заявленных убытков', 'Сумма урегулированных убытков',
        'Выплаты по неустойкам', 'Сумма убытка или ЗНУ', 'Сумма убытка или ЗНУ без неустоек', 'ДВОУ',
        'Расходы на урегулирование убытков', 'Расходы на урегулирование убытков (ОФР 25203)',
        'Сумма начисления по регрессу', 'Сумма поступления по регрессу', 'Год изготовления', 'Доля действия договора', 'КБМ по последнему полису']

        #удалить пробелы и преобразовать в числа
        for columns in num_column:
            df_una[columns] = df_una[columns].replace({r'\s+': '', r'\,': '.'}, regex=True)
            df_una[columns] = pd.to_numeric(df_una[columns], errors='coerce')


        #бумага электрон
        df_una.insert(2, 'Бумага/электронный', np.where(df_una['Номер полиса'].str.match(r'ХХХ.*'), 'Электронный', 'Бумага'))

        #преобразовать в даты
        date_column = ['Дата договора', 'Дата начала ответственности', 'Дата окончания ответственности','Дата начисления премии']
        for columns in date_column:
            df_una[columns] = df_una[columns].apply(pd.to_datetime, format='%d.%m.%Y', errors='coerce')

        def conclusion_date(row):
            if row is not None:
                value = row.strftime('%Y') + '_' + row.strftime('%m')
                return value
            return None
        df_una.insert(4, 'Год_месяц заключения', df_una['Дата договора'].apply(conclusion_date))

        # Заполнение филиалов
        def get_filial(row):
            if row['Филиал'] == 'ЦО Казань':
                return 'Е-ОСАГО'
            elif row['Филиал'] == 'Казань, филиал':
                if row['Подразделение'] == 'Департамент агентских продаж' or row[
                    'Подразделение'] == 'Департамент прямых продаж':
                    return 'ДАПП'
                elif row['Подразделение'] == 'Департамент партнерских продаж' or row[
                    'Подразделение'] == 'Департамент корпоративных продаж':
                    return 'ДКПП'
                else:
                    return 'Казань'
            else:
                return row['Филиал']

        df_una.insert(12, 'Филиалы', df_una.apply(get_filial, axis=1))

        max_date = pd.to_datetime(df_una['Дата начисления премии'].max())
        df_una.insert(15, 'Плавающий год', np.where(df_una['Дата начисления премии'] >= (max_date - pd.Timedelta(days=365)), 1,
                                                    np.where(df_una['Дата начисления премии'] >= (max_date - pd.Timedelta(days=730)), 2,
                                                     pd.NA)))


        df_una.insert(16, 'Период', df_una['Дата начисления премии'].dt.year)
        df_una.insert(17, 'Год и квартал договора', df_una['Дата начисления премии'].apply(lambda x: f"{x.year}_{(x.month - 1) // 3 + 1}" if pd.notna(x) else pd.NA))
        df_una.insert(20, "Субагент/кто ввел", "")
        df_una.insert(31, 'Возраст авто', (df_una['Период'] - df_una['Год изготовления']), 0)

        foreing_model = [
            "Acura", "Aito", "Alfa Romeo", "Asia", "Audi", "BAIC", "BAW", "Belgee", "Bentley", "BMW",
            "Brilliance", "BUELL", "Buick", "BYD", "Cadillac", "Caterpillar", "Changan", "Chery", "CheryExeed",
            "Chevrolet", "Chrysler", "Citroen", "Claas", "Dacia", "Dadi", "Daewoo", "DAF", "Daihatsu", "Datsun",
            "Derways", "DFM (Dongfeng)", "Dodge", "Dongfeng (DFM)", "DONINVEST", "DS", "Ducati", "Exeed", "FAW",
            "Fiat", "Ford", "Forthing", "Foton", "Fuego", "GAC", "Gaz-Mahindra", "Geely", "Genesis", "Geo", "GMC",
            "Great Wall", "HAFEI", "Haima", "Hangcha", "Harley-Davidson", "Haval", "Hawtai", "Hidromek", "HINO",
            "Honda", "Hongqi", "Howo", "Hummer", "Hyundai", "Infiniti", "Iran Khodro", "Isuzu", "IVECO", "JAC",
            "Jaecoo", "Jaguar", "JCB", "Jeep", "Jetour", "Jetta", "JMC", "John Deere", "Kaiyi", "Kawasaki", "Kia",
            "KTM", "Lancia", "Land Rover", "Landwind", "LDV", "Lexus", "LIFAN", "Lincoln", "Livan", "LiXiang", "Luxgen",
            "Mahindra", "MAN", "Maserati", "Maybach", "Mazda", "Mercedes-Benz", "Mercury", "MG", "Mini", "Mini (BMW)",
            "Mitsubishi", "Mitsubishi FUSO", "Nissan", "Oldsmobile", "OMODA", "Opel", "Perodua", "Peugeot", "Plymouth",
            "Pontiac", "Porsche", "Proton", "RAM", "Ravon", "Renault", "Rolls-Royce", "Rover", "Saab", "Samsung",
            "Saturn",
            "Scania", "Scion", "Seat", "Shacman", "Shanghai Maple", "Shuanghuan", "Skoda", "Skywell", "Smart",
            "Solaris (AGR)",
            "SsangYong", "Subaru", "Suzuki", "SWM", "Sym", "Talbot", "Tank", "Tata", "Terex", "Tesla", "Tianma",
            "Tianye",
            "Toyota", "Volkswagen", "Volvo", "Vortex", "Voyah", "Xiaomi", "Xin Kai", "Xpeng", "Yamaha", "Yulon",
            "Zeekr",
            "Zotye", "ZX", 'Другая марка (иностранная спецтехника)', 'Другая марка (иностранные мотоциклы и мотороллеры)'
            'Другая марка (иностранный автобус)', 'Другая марка (иностранный грузовой)', 'Другая марка (иностранный легковой)',
            'HUMMER', 'Dieci', 'Lifan', 'SITRAK', 'Bajaj', 'BAW', 'Denza', 'Doninvest', 'Fendt', 'Freightliner', 'Hafei',
            'Hamm', 'Komatsu', 'LONKING', 'Mack', 'SEAT', 'SYM']
        rus_model = [
            "Belarus", "IRBIS", "LADA", "Sollers", "Stels", "UAZ", "Xcite", "ZAZ",
            "Амкодор", "Богдан", "ВАЗ", "ВАЗ (LADA)", "ВАЗ/Lada", "ВИС", "Волжанин",
            "ГАЗ", "Другая марка (отечественная спецтехника)", "Другая марка (отечественные мотоциклы и мотороллеры)",
            "Другая марка (отечественный автобус)", "Другая марка (отечественный грузовой)",
            "Другая марка (отечественный легковой)", "ЗАЗ", "ЗИЛ", "ИВЕКО-УралАЗ", "ИЖ",
            "КамАЗ", "Кировец", "ЛиАЗ", "ЛТЗ", "ЛуАЗ", "МАЗ", "Москвич", "МТЗ",
            "НефАЗ", "ПАЗ", "Ростсельмаш", "СеАЗ", "Тагаз", "УАЗ", "Урал", "ЮМЗ", 'MAZ', 'Evolute', 'Minsk', 'КАВЗ', 'ХТЗ'
        ]
        df_una.insert(31, 'Иномарка/отечественная',
                      np.where(df_una['Марка ТС'].isin(foreing_model), 'Иностранного производства',
                               np.where(df_una['Марка ТС'].isin(rus_model), 'Отечественного производства', "")))

        def get_age(row):
            if row['Возраст авто'] <= 5:
                return '0-5 лет'
            elif row['Возраст авто'] <= 10:
                    return '6-10 лет'
            elif row['Возраст авто'] <= 14:
                    return '11-14 лет'
            else:
                return '15 лет и более'

        df_una.insert(33, 'Диапазон возраста авто', df_una.apply(get_age, axis=1))

        def get_segment(row):
            if row['Цель ТС'] == 'Учебная езда':
                return 'Учебная езда'
            elif "D" in row['Тип и назначение категория ТС']:
                    return 'Автобусы'
            elif row['Цель ТС'] == 'Такси':
                    return 'Такси'
            elif 'Мотоцикл' in row['Тип и назначение категория ТС']:
                    return 'Мотоциклы'
            elif 'Тракторы' in row['Тип и назначение категория ТС']:
                    return 'Тракторы'
            elif 'Трамв' in row['Тип и назначение категория ТС']:
                    return 'Трамваи'
            elif 'Троллейбус' in row['Тип и назначение категория ТС']:
                    return 'Троллейбусы'
            elif 'СЕ' in row['Тип и назначение категория ТС'] and row['Тип страхователя'] == 2:
                    return 'Грузовик, ФЛ'
            elif 'СЕ' in row['Тип и назначение категория ТС'] and row['Тип страхователя'] == 1 and row['КБМ по последнему полису'] >= 1.17:
                return 'Грузовик, ЮЛ, КБМ 1,17 и более'
            elif (row['Марка ТС'] == 'Mercedes-Benz' or row['Марка ТС'] == 'BMW') and row['Возраст авто'] > 10:
                return 'БМВ и Мерс старше 10 лет'
            elif row['Марка ТС'] == 'ГАЗ':
                return 'ГАЗ'
            elif row['Марка ТС'] == 'Daewoo' and row['Тип страхователя'] == 2 and row['КБМ по последнему полису'] >= 1.17:
                return 'Дэу ФЛ, КБМ 1,17 и более'
            elif row['Марка ТС'] == 'Daewoo' and row['Тип страхователя'] == 2 and row['КБМ по последнему полису'] == 1 and row['Год изготовления'] < 2015:
                return 'Дэу ФЛ, КБМ 1, старше 2015'
            elif row['Марка ТС'] == 'Renault' and 'Logan' in row['Модель ТС'] and row['Тип страхователя'] == 2 and row['КБМ по последнему полису'] == 1 and row['Год изготовления'] < 2015:
                return 'Логан ФЛ, КБМ 1, старше 2015'
            elif row['Марка ТС'] == 'Renault' and 'Logan' in row['Модель ТС'] and row['Тип страхователя'] == 2 and row['КБМ по последнему полису'] >= 1.17:
                return 'Логан ФЛ, КБМ 1,17 и более'
            else:
                return ''

        df_una.insert(38, "Сегмент",  df_una.apply(get_segment, axis=1))
        df_una.insert(50, "Отчисления РСА", df_una['Заработанная премия'] * 0.043)
        df_una.insert(51, "Кол-во дог. Страхования", 1)
        df_una.insert(62, 'Заработанное ДВОУ',
        np.where(df_una['ДВОУ'] is not None, (df_una['ДВОУ'] * df_una['Доля действия договора']), 0))

        def seg(row):
            segment_value = str(row['Сегмент заключения'])
            if "B_Base" in segment_value or "B_Increased" in segment_value or "B_Redused" in segment_value:
                return row['Сегмент заключения']
            else:
                return 'Прочие'

        df_una.insert(68, "Сегмент заключения для свода", df_una.apply(seg, axis=1))
        df_una.insert(73, "Административные расходы", "")
        df_una.insert(74, "Разница ПВУ", "")

        with pd.ExcelWriter('!Убыточность.xlsx', engine='openpyxl', mode='w') as excel_writer:
            df_una.to_excel(excel_writer, index=False, freeze_panes=(1, 0), sheet_name='Данные')


        import subprocess
        subprocess.call('!Убыточность.xlsx', shell=True)