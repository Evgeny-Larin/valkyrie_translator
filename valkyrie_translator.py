import PySimpleGUI as sg
from tempfile import mkdtemp
from zipfile import ZipFile
import pandas as pd
from shutil import move, rmtree
import os
import re


# функция для удаления файла из zip архива https://stackoverflow.com/questions/4653768/overwriting-file-in-ziparchive
def remove_from_zip(zipfname, *filenames):
    tempdir = mkdtemp()
    try:
        tempname = os.path.join(tempdir, 'new.zip')
        with ZipFile(zipfname, 'r') as zipread:
            with ZipFile(tempname, 'w') as zipwrite:
                for item in zipread.infolist():
                    if item.filename not in filenames:
                        data = zipread.read(item.filename)
                        zipwrite.writestr(item, data)
        move(tempname, zipfname)
    finally:
        rmtree(tempdir)


# функция для вызова нового окна с выбором исходного языка для перевода
def popup_select(the_list):
    layout = [[sg.Listbox(the_list, key='_LIST_', size=(45, len(the_list)), select_mode='single', bind_return_key=True),
               sg.OK(), sg.Button('Назад')]]
    window = sg.Window('Valkyrie Translator', layout=layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Назад':
            window.close()
            break
        if values['_LIST_'] == []:
            sg.Popup('Язык не выбран!')
            continue
        else:
            window.close()
            return values['_LIST_'][0]


# функция для замены подстроки по словарю
def replacer(text, dictionary):
    # в некоторых сценариях значение text может быть NaN , пропускаем такой текст
    if isinstance(text, float):
        return ""
    for word, replacement in dictionary.items():
        text = text.replace(word, replacement)
    return text


def transform_file(path, loc_select):
    path_dir = '\\'.join(path.split('/')[:-1])  # это путь до папки, содержащей файл
    # открываем файл, читаем с помощью pandas
    with archive.open(loc_select) as text:
        df = pd.read_csv(text,
                         sep='^([^,]+),',
                         engine='python',
                         header=None,
                         usecols=range(1, 3))
        df.rename(columns={2: 'text'}, inplace=True)
        df.fillna('', inplace=True)

    # ищем все системные {операторы} и <операторы>
    special_symb = df['text'].apply(lambda x: re.findall(r'{.*?}?[}]|<.*?>', x)).tolist()
    # разворачиваем двумерный список в одномерное множество (чтобы убрать дубликаты)
    special_symb = {a for b in special_symb for a in b}
    # переводим в список, сортируем по длине оператора
    # сначала заменятся большие операторы, потом маленькие, так не столкнемся с заменой маленького оператора внутри большого
    special_symb = sorted(list(special_symb), key=len, reverse=True)
    # присваиваем номер каждому системному {оператору}, сохраняем в словарь
    special_symb = {k: '<' + str(v + 1) + '>' for v, k in enumerate(special_symb)}
    special_symb[r'\n'] = '<' + str(0) + '>'
    # заменяем все вхождения системных операторов на их номер из словаря
    df['text'] = df['text'].apply(lambda x: replacer(x, special_symb))

    # сохраняем текст для перевода
    name = path.split('/')[-1].split('.')[-2] + ' TextForTranslate.xlsx'
    df['text'].to_excel(f'{path_dir}\\{name}', index=False)

    # возвращаем преобразованный df и словарь операторов
    return df, special_symb


def update_file(translated_df, df, special_symb):
    # иногда переводчик добавляет лишние пробелы или неправильно распознает специальные символы - исправляем
    translated_df = translated_df.replace(to_replace=' >', value='>', regex=True) \
        .replace(to_replace='< ', value='<', regex=True)
    translated_df.rename(columns={translated_df.columns[0]: 'text'}, inplace=True)

    # слепляем столбец триггеров и столбец с переведенным текстом
    translated_df = df.merge(translated_df,
                             left_index=True,
                             right_index=True)
    translated_df = translated_df[[1, 'text_y']]
    translated_df.rename(columns={'text_y': 'text'}, inplace=True)

    translated_df.loc[0, translated_df.columns[1]] = 'Russian'

    # создаем инвертированный словарь и заменяем порядковые номера на специальные значения
    special_symb = dict(zip(special_symb.values(), special_symb.keys()))
    translated_df['text'] = translated_df['text'].apply(lambda x: replacer(x, special_symb))
    translated_df['text'] = translated_df['text'].apply(lambda x: x.strip() if isinstance(x, str) else x)
    return translated_df


def load_to_zip(translated_df, archive):
    # добавляем файл в архив с сюжетом
    translated_df.to_csv('Localization.Russian.txt', sep=',', index=False, header=False)
    archive.write('Localization.Russian.txt')
    os.remove('Localization.Russian.txt')

    # извлекаем файл quest.ini и удаляем его в архиве сюжета
    archive.extract('quest.ini')
    archive.close()
    remove_from_zip(fr'{path}', 'quest.ini')

    # редактируем файл quest.ini и добавляем в архив с сюжетом
    file = open('quest.ini', 'r')
    pozition = file.readlines().index('[QuestText]\n') + 1
    file = open('quest.ini', 'r')
    new_file = file.readlines()
    elem = 'Localization.Russian.txt\n'
    new_file.insert(pozition, elem)

    with open('quest.ini', 'w') as file:
        for row in new_file:
            s = "".join(map(str, row))
            file.write(s)

    archive = ZipFile(fr'{path}', 'a')
    archive.write('quest.ini')
    archive.close()
    os.remove('quest.ini')


sg.theme('SystemDefault')  # тема окна
# элементы окна
layout = [[sg.Text('1. Скачайте сюжет с поддержкой английского языка через Valkyrie\n'
                   '2. Выберите файл сюжета Valkyrie (например, Possessed.valkyrie)\n'
                   'По умолчанию он находятся в папке \\AppData\\Roaming\\Valkyrie\\Download')],
          [sg.Text('Файл сюжета:'), sg.InputText(),
           sg.FileBrowse(initial_folder=f'{os.getenv("APPDATA")}\\Valkyrie\\Download',
                         file_types=(('.valkyrie', '*.valkyrie'),))],
          [sg.Button('OК'), sg.Button('Закрыть')]]

window = sg.Window('Valkyrie translator', layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Закрыть':  # закрывается программа
        break
    elif values[0] == '':  # если ничего не выбрано
        sg.Popup('Файл не выбран!')
        continue
    else:
        path = values[0]  # это путь до файла

        # открываем архив в режиме 'a' (для добавления файлов в архив)
        archive = ZipFile(fr'{path}', 'a')
        # получаем весь список файлов в архиве
        files_list = [text_file.filename for text_file in archive.infolist()]
        # проверяем, есть ли русская локализация в файле, иначе получаем список других локализаций
        if 'Localization.Russian.txt' in files_list:
            sg.Popup('Сюжет уже переведен на русский язык')
            continue
        else:
            # если файл содержит 'Localization', сохраняем в словарь пару 'название файла':'язык'
            loc_list = {file.split('.')[1]: file for file in files_list if file.find('Localization') >= 0}
            # выбор пользователем исходного языка
            loc_select = popup_select(list(loc_list.keys()))
            loc_select = loc_list[loc_select]

        df, special_symb = transform_file(path, loc_select)

    while True:
        file_translated = sg.PopupGetFile('Текст извлечён и сохранен в папку с сюжетами в формате .xlsx\n'
                                          'Переведите файл в Google Переводчике, скачайте и укажите\n'
                                          'путь к переведенному файлу', file_types=(('Excel XLSX', '*.xlsx'),),
                                          initial_folder=os.path.join(os.environ['USERPROFILE'], "Downloads"))
        if file_translated == '':
            sg.Popup('Файл не выбран!')
            continue
        elif file_translated is None:
            break
        else:
            # считываем переведенный файл
            translated_df = pd.read_excel(fr'{file_translated}')
            if translated_df.columns[-1] == 'text':
                sg.Popup('Файл не переведен!')
                continue

            translated_df = update_file(translated_df, df, special_symb)

            load_to_zip(translated_df, archive)

            sg.Popup('Готово!\nФайл сюжета обновлен!\nМожно играть!')
            break

window.close()
