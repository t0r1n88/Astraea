"""
Функция для обработки данных из Яндекс и гугл форм
"""
from astraea_support_functions import count_value_in_column
import pandas as pd
import openpyxl
from collections import Counter




def extract_answer_several_option(row:pd.Series):
    """
    Функция для извлечения ответов из колонок
    :param row:строка датафрейма с определенными колонками
    :return: строка значений разделенных точкой с запятой
    """
    temp_lst = row.tolist() # делаем список
    temp_lst = [value for value in temp_lst if pd.notna(value)] # отбрасываем незаполненное
    return ';'.join(temp_lst)





def extract_data_from_form(path_to_file:str,path_end_folder:str):
    """

    :param path_to_file:
    :param path_end_folder:
    :return:
    """
    dct_df = dict() # словарь для хранения листовых датафреймов
    sev_df = pd.DataFrame() # выходной датафрейм для нескольких вопросов
    all_df = pd.DataFrame() # выходной датафрейм для нескольких вопросов и шкалы

    check_set_answer = set()  # множество для проверки есть ли такой вопрос или нет

    count_question = 1


    df = pd.read_excel(path_to_file,dtype=str)

    for idx, name_column in enumerate(df.columns):
        # получаем последующую колонку
        if idx +1 == len(df.columns):
            # проверяем достижение предела
            cont_name_column = idx
        else:
            cont_name_column = idx+1
        # отбрасываем обычные колонки
        if ' / ' not in name_column:
            # Проверяем наличие точки с запятой, что может свидетельствовать о гугловской таблице
            contains_at_symbol = list(df[name_column].str.contains(';'))
            if True not in contains_at_symbol:
                # Создаем частотную таблицу
                simple_df = df[name_column].value_counts(sort=True).to_frame().rename(columns={'count':'Количество'}).reset_index()
                simple_df['Доля в % от общего'] = round(
                    simple_df[f'Количество'] / simple_df['Количество'].sum(), 2) * 100
                simple_df.loc['Итого'] = simple_df.sum(numeric_only=True)

                dct_df[idx+1] = simple_df
                count_question += 1
                continue
            else:
                # создаем частотную таблицу
                counts_df = count_value_in_column(df.copy(), name_column)
                dct_df[idx + 1] = counts_df

                count_question += 1

        else:
            lst_union_name_column = [name_column] # список для хранения названий колонок
            lst_answer = name_column.split(' / ')
            # Получаем ответ
            question = lst_answer[0] # Вопрос
            if question not in check_set_answer:
                threshold = len(df.columns[cont_name_column:]) # сколько колонок осталось до конца датафрейма
                for temp_idx,temp_name_column in enumerate(df.columns[cont_name_column:]):
                    temp_lst_question = temp_name_column.split(' / ') # сплитим каждую последующую колонку пока вопрос равен предыдущему
                    temp_question = temp_lst_question[0] # вопрос в следующей колонке
                    if question == temp_question:
                        lst_union_name_column.append(temp_name_column) # добавляем название
                        # для случая если такой вопрос последний в таблице
                        if temp_idx+1 == threshold:
                            sev_df[question] = df[lst_union_name_column].apply(extract_answer_several_option, axis=1)
                            all_df[question] = df[lst_union_name_column].apply(extract_answer_several_option, axis=1)
                            # создаем частотную таблицу
                            counts_df = count_value_in_column(sev_df.copy(), question)
                            dct_df[idx+1] = counts_df
                            break


                    else:
                        sev_df[question] = df[lst_union_name_column].apply(extract_answer_several_option,axis=1)
                        all_df[question] = df[lst_union_name_column].apply(extract_answer_several_option,axis=1)
                        # создаем частотную таблицу
                        counts_df = count_value_in_column(sev_df.copy(),question)

                        dct_df[idx+1] = counts_df
                        check_set_answer.add(question)

                        count_question += 1

            else:
                continue

    with pd.ExcelWriter(f'{path_end_folder}/Общий результат.xlsx', engine='xlsxwriter') as writer:
        for name_sheet,out_df in dct_df.items():
            out_df.to_excel(writer,sheet_name=str(name_sheet),index=False)

































if __name__ == '__main__':
    main_file = 'data/Яндекс таблица.xlsx'
    # main_file = 'data/2025-07-17 Anketa 9 klass.xlsx'
    # main_file = 'data/Гугл таблица.xlsx'
    main_end_folder = 'data/Результат'


    extract_data_from_form(main_file,main_end_folder)
    print('Lindy Booth')