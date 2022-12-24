#!/usr/bin/env python
# coding: utf-8

# In[1]:


#библиотека для перевода внутри python
#решил пока не использовать потому, что выдает перевод хуже, чем в ручном режиме

#!pip install translators
#import translators as ts
#ts.translate_text('mansions', translator = 'google', to_language = 'ru')


# In[2]:


import PySimpleGUI as sg
from tempfile import mkdtemp
from zipfile import ZipFile
import pandas as pd
from shutil import move, rmtree
import os


# In[3]:


#функция для удаления файла из zip архива https://stackoverflow.com/questions/4653768/overwriting-file-in-ziparchive
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


# In[4]:


def special_symb(): #запоминаем системные значения {}
    special = df.text[df.text.str.contains('{') == True].unique().tolist() #ищем все строчки, в которых есть {}
    sep = '}'
    list2 = []
    for i in special:  #отделяем {} от {} если они в одной строчке
        if i.find('{') >= 0 and i.find('}') >= 0:
            list2.extend([x+sep for x in i.split(sep)])

    special = []       #чистим мусор (наверное можно оптимизировать, пока хз как)
    for i in list2:
        start = i.find('{')
        if start >= 0:
            end = i.find('}')+1
            special.append(i[start:end])

    special = set(special) #удаляем дубликаты
    
    global special_dict
    special_dict = {}
    special_list = []
    n = 0
    for i in special: #собираем все специальные значения
        start = i.find('{')
        end = i.find('}', start)+1
        special_list.append(i[start:end])
    special_list = set(special_list) #сохраняем только уникальные специальные значения
    for i in special_list: #сохраняем уникальные значения в словарь и присваиваем порядковые номера
        n = n+1
        special_dict[i] = '<'+str(n)+'>'


# In[5]:


def special_symb2(): #запоминаем системные значения <>
    special2 = df.text[df.text.str.contains('<') == True].unique() #ищем все строчки, в которых есть <>
    sep = '>'
    list2 = []
    for i in special2:  #отделяем {} от {} если они в одной строчке
        if i.find('<') >= 0 and i.find('>') >= 0:
            list2.extend([x+sep for x in i.split(sep)])

    special2 = []       #чистим мусор (наверное можно оптимизировать, пока хз как)
    for i in list2:
        start = i.find('<')
        if start >= 0:
            end = i.find('>')+1
            special2.append(i[start:end])
    special2 = set(special2) #удаляем дубликаты

    global special_dict2
    special_dict2 = {}
    special_list2 = []
    n = 0
    for i in special2: #собираем все специальные значения
        start = i.find('<')
        end = i.find('>', start)+1
        special_list2.append(i[start:end])
    special_list2 = set(special_list2) #сохраняем только уникальные специальные значения
    for i in special_list2: #сохраняем уникальные значения в словарь и присваиваем порядковые номера
        n = n-1
        special_dict2[i] = '<'+str(n)+'>' 


# In[6]:


sg.theme('SystemDefault')   # Add a touch of color
# All the stuff inside your window.
layout = [  [sg.Text('1. Скачайте сюжет с поддержкой английского языка через Valkyrie\n2. Выберите файл сюжета Valkyrie (например, Possessed.valkyrie)\nПо умолчанию он находятся в папке \\AppData\\Roaming\\Valkyrie\\Download')],
            [sg.Text('Файл сюжета:'), sg.InputText(), sg.FileBrowse(file_types=(('.valkyrie', '*.valkyrie'),)), ],
            [sg.Button('Ok'), sg.Button('Cancel')]]

window = sg.Window('Valkyrie translator v0.2, dev: Evgeny Larin', layout)

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
    elif values[0] == '':
        sg.Popup('Файл не выбран!')
        continue
    path = values[0] #это путь до файла
    path_dir = '\\'.join(path.split('/')[:-1])  #это путь до папки, содержащей файл
      
    z = ZipFile(fr'{path}', 'a') #открываем архив
    #Проверяем, есть ли в архиве перевод
    files_list = [text_file.filename for text_file in z.infolist() ]
    if 'Localization.Russian.txt' in files_list:
        sg.Popup('Сюжет уже переведен на русский язык')
        continue
    
    #открываем файл, читаем его пандасом
    with z.open('Localization.English.txt') as f: 
        df = pd.read_csv(f,
                         sep='^([^,]+),',
                         engine='python',
                         header = None,
                         usecols = range(1,3))
        df.rename(columns={2:'text'}, inplace = True)
    
    #ручные замены
    df = df.replace(to_replace = '</i\|\|\|', value = '</i>|||', regex = True)
    
    special_symb()  #запоминаем системные значения {}
    special_symb2() #запоминаем системные значения <>
    
    #заменяем специальные значения на порядковые номера (чтобы переводчик их не переводил) 
    df["text"] = df.text.replace(special_dict2, regex=True).replace(special_dict, regex=True)
    
    #сохраняем текст для перевода
    name = path.split('/')[-1].split('.')[-2] + ' TextForTranslate.xlsx'
    df.text.to_excel(f'{path_dir}\\{name}')

    while True:
        file_translated = sg.PopupGetFile('Текст извлечён и сохранен в папку с сюжетами в формате .xlsx\nПереведите файл в Google Переводчике, скачайте и укажите\nпуть к переведенному файлу', file_types=(('Excel XLSX', '*.xlsx'), ))
        if file_translated == '':
            sg.Popup('Файл не выбран!')
            continue
        elif file_translated == None:
            break
        else:
            #считываем переведенный файл
            df2 = pd.read_excel(fr'{file_translated}', usecols = [1])
            
            if df2.columns[-1] == 'text':
                sg.Popup('Файл не переведен!')
                continue
            
            #иногда переводчик добавляет лишние пробелы или неправильно распознает специальные символы - исправляем
            df2 = df2.replace(to_replace = ' >', value = '>', regex = True)                     .replace(to_replace = '\\\п', value = '\\n', regex = True)
            
            #слепляем столбец триггеров и столбец с переведенным текстом
            df_done = df.merge(df2, left_index=True, right_index=True).drop(columns=['text'])
            df_done.rename(columns = {'текст':'text'}, inplace = True)
            
            #создаем инвертированный словарь и заменяем порядковые номера на специальные значения
            dict_special = dict(zip(special_dict.values(), special_dict.keys()))
            dict_special2 = dict(zip(special_dict2.values(), special_dict2.keys()))
            df_done['text'] = df_done['text'].replace(dict_special, regex=True).replace(dict_special2, regex=True).replace('Английский', 'Russian', regex = True)
            
            #удаляем пробелы, который добавляет переводчик 
            df_done['text'] =  df_done['text'].str.strip()
            
            #добавляем файл в архив с сюжетом
            df_done.to_csv('Localization.Russian.txt', sep = ',', index = False, header = False)
            z.write('Localization.Russian.txt')
            os.remove('Localization.Russian.txt')
            
            #извлекаем файл quest.ini и удаляем его в архиве сюжета
            z.extract('quest.ini')
            z.close()
            remove_from_zip(fr'{path}', 'quest.ini')   
            
            #редактируем файл quest.ini и добавляем в архив с сюжетом
            file = open('quest.ini', 'r')
            pozition = file.readlines().index('[QuestText]\n')+1
            file = open('quest.ini', 'r')
            new_file = file.readlines()
            elem = 'Localization.Russian.txt\n'
            new_file.insert(pozition, elem)

            with open('quest.ini', 'w') as file:
                for row in new_file:
                    s = "".join(map(str, row))
                    file.write(s)

            z = ZipFile(fr'{path}', 'a')
            z.write('quest.ini')
            z.close()
            os.remove('quest.ini')
            
            sg.Popup('Готово!\nФайл сюжета обновлен!\nМожно играть!')
            break

window.close()


# In[ ]:




