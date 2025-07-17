"""
Вспомогательные функции
"""
import pandas as pd


def count_value_in_column(col:pd.Series,sep_str:str):
    """
    Функция для подсчета сколько раз то или иное значение встречается в колонке
    :param col: серия
    :param sep_str: разделитель значений внутри ячейки
    :return: словарь
    """
    print(col)
    print(type(col))
