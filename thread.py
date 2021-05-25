"""Потоки для отображения прогресса операций"""

import os
from savepdf import ToPDF
from docxtpl import DocxTemplate
from PyQt5.QtCore import QThread, pyqtSignal


class ThreadDOCX(QThread):
    """Поток для формирования документов DOCX"""
    # Создаём собственный сигнал,
    # принимающий параметр типа int:
    signal = pyqtSignal(int)

    def __init__(self, *args, parent=None):
        """Инициализация потока"""
        QThread.__init__(self, parent)
        self.studs, self.tplfile, self.curr_tpls, self.context, self.docdir, self.docpaths = args

    def run(self):
        # Просматриваем все пути исходных DOCX-файлов по каждому студенту:
        i = 0  # Счётчик обработанных пакетов документов
        for curr_tpl in self.curr_tpls:
            # Задаём папку для группы с названием шаблона:
            folder = f"{self.context['group']} - {curr_tpl}"
            if not os.path.isdir(f'{self.docdir}/{folder}'):
                os.mkdir(f'{self.docdir}/{folder}')
            # Создаём документы для всех студентов группы:
            for s in self.studs:
                filedoc = f"{self.docdir}/{folder}/{s['student']} - {curr_tpl}.docx"

                if s['student'] in self.docpaths:
                    self.docpaths[s['student']] = self.docpaths[s['student']] + [filedoc]
                else:
                    self.docpaths[s['student']] = [filedoc]

                doc = DocxTemplate(f'tpl/{self.tplfile[curr_tpl]}')
                doc.render(s)
                doc.save(filedoc)
                i += 1  # Увеличиваем счётчик
                # Отправляем значение счётчика в основную программу:
                self.signal.emit(i)


class ThreadPDF(QThread):
    """Поток для формирования пакетов документов в формате PDF"""
    # Создаём собственный сигнал,
    # принимающий параметр типа int:
    signal = pyqtSignal(int)

    def __init__(self, *args, parent=None):
        """Инициализация потока"""
        QThread.__init__(self, parent)
        self.docpaths, self.curr_packdocs = args

    def run(self):
        # Просматриваем все пути исходных DOCX-файлов по каждому студенту:
        i = 0  # Счётчик обработанных пакетов документов
        for student, doc_files in self.docpaths:
            # Конвертируем DOCX-файлы
            # каждого студента в отдельную папку:
            for file in doc_files:
                # Папка для студента /Фамилия Имя Отчество - Пакет документов на практику:
                folder = f"{self.curr_packdocs}/{student} - Пакет документов на практику"
                # Конвертация DOCX -> PDF. Исходные файлы остаются без изменений.
                # file - путь к документу DOCX, который надо конвертировать:
                ToPDF(folder).doc2pdf(file)
            i += 1  # Увеличиваем счётчик
            # Отправляем значение счётчика в основную программу:
            self.signal.emit(i)
