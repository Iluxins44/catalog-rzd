import pandas as pd

# Загрузка файлов Excel
file_mtr = 'C:/Users/smirn/OneDrive/Рабочий стол/ржд 2/merged_output.xlsx'
file_okpd2 = 'C:/Users/smirn/OneDrive/Рабочий стол/ржд 2/ED_IZM.xlsx'
# Чтение данных из файлов
mtr_df = pd.read_excel(file_mtr, sheet_name='merged_output')
okpd2_df = pd.read_excel(file_okpd2, sheet_name='ED_IZM')

# Выполнение INNER JOIN по столбцам ОКПД2 и OKPD2
merged_df = pd.merge(mtr_df, okpd2_df, left_on='Базисная Единица измерения', right_on='Код ЕИ', how='inner')

# Сохранение результата в новый файл Excel
merged_df.to_excel('merged_ot.xlsx', index=False)

print("Данные успешно сопоставлены и сохранены в 'merged_ot.xlsx'")