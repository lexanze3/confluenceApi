# confluenceApi

Для запуска требуется python 3.11 или выше

предварительно настроить окружение python:
pip install -U setuptools wheel
pip install -r requirements.txt

запуск скрипта:
python collectingFiles.py {username} {password} {space_key} {url}
где:
username - имя пользователя под которым будет запущен экспорт
password - пароль пользователя
space_key - ключ пространства confluence из страниц которого выгружаются страницы (находится в инструменты пространства -> обзор)
url - путь до родительской страницы

Страницы выгружаются в соответствии с иерархией указанной в confluence (меню переходов между страницами)