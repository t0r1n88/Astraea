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


    df = pd.read_excel(path_to_file,dtype=str)

    for idx, name_column in enumerate(df.columns):
        # получаем предыдущую и последующую колонку
        if idx +1 == len(df.columns):
            cont_name_column = idx
        else:
            cont_name_column = idx+1
        # отбрасываем обычные колонки
        if ' / ' not in name_column:
            # Сохраняем частотную таблицу
            dct_df[idx+1] = df[name_column].value_counts(sort=True).to_frame().rename(columns={'count':'Количество'}).reset_index()
            sev_df[name_column] = df[name_column]
            all_df[name_column] = df[name_column]
            continue
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

                            lst_count = [] # список для хранения значений которые были разделены точкой с запятой
                            tmp_lst = sev_df[question].tolist()
                            for value_str in tmp_lst:
                                lst_count.extend(value_str.split(';'))

                            # Делаем частотную таблицу и сохраняем в словарь
                            counts_df = pd.DataFrame.from_dict(dict(Counter(lst_count)),orient='index')
                            counts_df = counts_df.reset_index()
                            counts_df.columns = [question,'Количество']
                            counts_df.sort_values(by='Количество',ascending=False,inplace=True)
                            dct_df[idx + 1] = counts_df

                            break


                    else:
                        sev_df[question] = df[lst_union_name_column].apply(extract_answer_several_option,axis=1)
                        all_df[question] = df[lst_union_name_column].apply(extract_answer_several_option,axis=1)
                        lst_count = []  # список для хранения значений которые были разделены точкой с запятой
                        tmp_lst = sev_df[question].tolist()
                        for value_str in tmp_lst:
                            lst_count.extend(value_str.split(';'))

                        # Делаем частотную таблицу и сохраняем в словарь
                        counts_df = pd.DataFrame.from_dict(dict(Counter(lst_count)), orient='index')
                        counts_df = counts_df.reset_index()
                        counts_df.columns = [question, 'Количество']
                        counts_df.sort_values(by='Количество', ascending=False, inplace=True)
                        dct_df[idx + 1] = counts_df
                        check_set_answer.add(question)

            else:
                continue

    with pd.ExcelWriter(f'{path_end_folder}/Общий результат.xlsx', engine='xlsxwriter') as writer:
        for name_sheet,out_df in dct_df.items():
            out_df.to_excel(writer,sheet_name=str(name_sheet),index=False)

































if __name__ == '__main__':
    main_file = 'data/Яндекс таблица.xlsx'
    main_end_folder = 'data/Результат'


    extract_data_from_form(main_file,main_end_folder)
    print('Lindy Booth')