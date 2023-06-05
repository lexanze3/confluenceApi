#!/usr/bin/env python
# coding: utf-8
from docxcompose.composer import Composer
from win32com.client import constants
from atlassian import Confluence
import win32com.client as win32
from docx import Document
from urllib.parse import urlparse
import os
import re
import sys
import shutil

#создать директорию files, для временных файлов
path = os.path.join(os.getcwd(), "files")
shutil.rmtree(path,ignore_errors=True)
os.makedirs(path, exist_ok=True)
    
#параметры
username=sys.argv[1]
password=sys.argv[2]
space=sys.argv[3]
url_full=sys.argv[4]
parsed_url=urlparse(url_full)
url="{0}://{1}".format(parsed_url.scheme, parsed_url.netloc)

#глобальные переменные
confluence = Confluence(url=url,username=username,password=password)
id_position = dict();

def get_all_pages(confluence, space):
    start = 0
    limit = 100
    _all_pages = []
    while True:
        pages = confluence.get_all_pages_from_space(space, start, limit, status=None, expand=None, content_type='page')
        _all_pages = _all_pages + pages
        if len(pages) < limit:
            break
        start = start + limit
    return _all_pages

def get_id_pos():
    id_position = dict()
    pages = get_all_pages(confluence, space)
    for page in pages:
        id_position[page['id']] = page['extensions']['position'] if page['extensions']['position'] != 'none' else page['title'] 
    return id_position
    
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

def start_parser(url_full):
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
    for json in json_list:
        json['position'] = id_position[json['content']['id']]
    json_list.sort(key=lambda x: x['position'])
    return json_list

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
#загрузить позицию страниц в дереве переходов
id_position = get_id_pos()
#начало работы по сбору страниц
global_title = start_parser(url_full)
#загрузка страниц в doc файлы
all_paths = save_files(contents)
#преобразование их в .docx
new_paths=convert_doc_to_docx(all_paths)
#Объединение всех файлов в 1 заранее подготовленный шаблон
composer(new_paths)
#удаление временых файлов
for i in new_paths:
    os.remove(i)

