"""
Функция для обработки данных из Яндекс и гугл форм
"""
import pandas as pd
import openpyxl




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
        if idx == 0:
            prev_name_column = ''
            cont_name_column = df.columns[idx+1]
        else:
            if idx +1 == len(df.columns):
                prev_name_column = idx-1
                cont_name_column = idx
            else:
                prev_name_column = idx-1
                cont_name_column = idx+1
        # отбрасываем обычные колонки
        if ' / ' not in name_column:
            sev_df[name_column] = df[name_column]
            all_df[name_column] = df[name_column]
            continue
        else:
            lst_union_name_column = [name_column] # список для хранения названий колонок
            print(name_column)
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
                            break


                    else:
                        sev_df[question] = df[lst_union_name_column].apply(extract_answer_several_option,axis=1)
                        all_df[question] = df[lst_union_name_column].apply(extract_answer_several_option,axis=1)
                        check_set_answer.add(question)

            else:
                continue
    sev_df.to_excel('data/sev.xlsx',index=False)
    all_df.to_excel('data/all.xlsx',index=False)
































if __name__ == '__main__':
    main_file = 'data/Яндекс таблица.xlsx'
    main_end_folder = 'data/Результат'


    extract_data_from_form(main_file,main_end_folder)
    print('Lindy Booth')