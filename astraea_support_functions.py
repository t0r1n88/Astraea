"""
Вспомогательные функции
"""
import pandas as pd
from collections import Counter

def count_value_in_column(df:pd.DataFrame,name_column:str):
    """
    Функция для подсчета сколько раз то или иное значение встречается в колонке
    :param col: серия
    :return: датафрейм
    """
    lst_count = []  # список для хранения значений которые были разделены точкой с запятой
    tmp_lst = df[name_column].tolist()
    for value_str in tmp_lst:
        lst_count.extend(value_str.split(';'))

    # Делаем частотную таблицу и сохраняем в словарь
    counts_df = pd.DataFrame.from_dict(dict(Counter(lst_count)), orient='index')
    counts_df = counts_df.reset_index()
    counts_df.columns = [name_column, 'Количество']
    counts_df.sort_values(by='Количество', ascending=False, inplace=True)

    counts_df['Доля в % от общего'] = round(
        counts_df[f'Количество'] / counts_df['Количество'].sum(), 2) * 100

    counts_df.loc['Итого'] = counts_df.sum(numeric_only=True)



    return counts_df

