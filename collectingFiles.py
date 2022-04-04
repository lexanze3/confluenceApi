#!/usr/bin/env python
# coding: utf-8
from docxcompose.composer import Composer
from win32com.client import constants
from atlassian import Confluence
import win32com.client as win32
from docx import Document
import os
import re

#ваша ссылка
url_full='https://сonfluence.ru/'

confluence = Confluence(
    url='https://сonfluence.ru/',
    #логин
    username='********',
    #пароль
    password='******')

def get_id(url_full):
    return url_full.split('=')[-1]

def get_child(id):
    t = confluence.cql(cql='parent={0}'.format(id), start=0, limit=None, expand=None, include_archived_spaces=None, excerpt=None)['results']
    return sort_json(t)

def cursor(list_child,level=0):
    for i in list_child:
        id = i['content']['id']
        print(level, id)
        if len(get_child(id))>0:
            page = confluence.get_page_by_id(page_id=id)
            response = confluence.get_page_as_word(page['id'])
            contents.append(response)
            cursor(get_child(id),level=level+1)
        else:
            page = confluence.get_page_by_id(page_id=id)
            response = confluence.get_page_as_word(page['id'])
            contents.append(response)

def start_parser(url,url_full):
    #Получили ID родителя
    id_url = get_id(url_full)
    print(id_url)
    page = confluence.get_page_by_id(page_id=id_url)
    response = confluence.get_page_as_word(page['id'])
    contents.append(response)
    #Получаем вложенные объекты
    list_child = get_child(id_url)
    cursor(list_child)
    return page['title']

def save_files(contents):
    paths=list()
    for i,j in enumerate(contents):
        path = os.path.join(os.getcwd(),'files',f'{i}.doc')
        with open(path, mode='wb') as file_pdf:
            file_pdf.write(j)
        paths.append(path)
    return paths

def sort_json(json_list):
    titles = list()
    for i in json_list:
        titles.append(i['title'])
    titles.sort()
    sort_json = list()
    for i in titles:
        for j in json_list:
            if i==j['title']:
                sort_json.append(j)
    return sort_json

def convert_doc_to_docx(files):
    paths=list()
    for path in files:
        word = win32.gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(path)
        doc.Activate()
        # Rename path with .docx
        new_file_abs = os.path.abspath(path)
        new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
        # Save and Close
        word.ActiveDocument.SaveAs(
            new_file_abs, FileFormat=constants.wdFormatXMLDocument
        )
        paths.append(new_file_abs)
        doc.Close(False)
        os.remove(path)
    return paths

def composer(files):
    master = Document('pattern.docx')
    composer = Composer(master)
    for file in files:
        doc2 = Document(file)
        if file != files[-1]:
            doc2.add_page_break()
        composer.append(doc2)
    composer.save(f"{global_title}.docx")
    print('load "{0}" complite'.format(global_title))
    
contents = list()
#начало работы по сбору страниц
global_title = start_parser(url,url_full)
#загрузка страниц в doc файлы
all_paths = save_files(contents)
#преобразование их в .docx
new_paths=convert_doc_to_docx(all_paths)
#Объединение всех файлов в 1 заранее подготовленный шаблон
composer(new_paths)
#удаление временых файлов
for i in new_paths:
    os.remove(i)

