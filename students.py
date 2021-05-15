"""Students 0.7.6"""

import os
import sys
import xlrd
import subprocess
from PyQt5 import uic
from docxtpl import DocxTemplate
from PyQt5.QtWidgets import (
    QMessageBox,
    QApplication,
    QMainWindow,
    QFileDialog,
    QTableWidgetItem,
    QProgressDialog
)


# Глобальный спиcок студентов. Данные по каждому студенту считываются из файла XLS в словарь.
# Одна строка таблицы - это один студент, данные по которому образуют один словарь.
# В итоге получается список словарей.
studs = list()


class Students(QMainWindow):
    def __init__(self):
        super().__init__(None)
        # Загрузка интерфейса:
        uic.loadUi('students.ui', self)

        # Файл ещё не открыт:
        self.fileopen = False

        # Ошибок открытия файла ещё не было:
        self.errorOpen = False

        # Кнопки pb_save_docx и pb_save_pdf дезактивированы,
        # поскольку на данный момент ещё нечего сохранять:
        self.pb_save_docx.setDisabled(True)
        self.pb_save_pdf.setDisabled(True)

        # Открыть файл XLS. Сигнал pb_open_xls --> слот open_xls:
        self.pb_open_xls.clicked.connect(self.open_xls)

        # Сохранить файл DOCX. Сигнал pb_save_docx --> слот savedocx:
        self.pb_save_docx.clicked.connect(self.savedocx)

        # Сохранить файл PDF. Сигнал pb_save_pdf --> слот savepdf:
        self.pb_save_pdf.clicked.connect(self.savepdf)
        self.statusBar().showMessage('Изучите инструкцию и приступайте к работе')

    def open_xls(self):
        """Чтение файла XLS"""
        if self.fileopen:
            QMessageBox.information(
                self,
                'Инфо',
                '<h4>Файл уже был открыт,<br>но можно выбрать другой.</h4>'
            )
            fname, _ = QFileDialog.getOpenFileName(
                self,
                'Выбрать файл',
                None,
                'Microsoft Excel (*.xls)'
            )
        else:
            fname, _ = QFileDialog.getOpenFileName(
                self,
                'Выбрать файл',
                None,
                'Microsoft Excel (*.xls)'
            )

        try:
            # Открываем книгу XLS
            workbook = xlrd.open_workbook(fname)

            # Читаем первый лист:
            sh = workbook.sheet_by_index(0)

            # Пробегаем по строкам таблицы:
            for i in range(sh.nrows):
                # Считываем строку из таблицы:
                row = sh.row_values(i)
                # Создаём временный словарь для студента
                # и помещаем его в глобальный список studs:
                tmp = {
                    'surname': row[0],
                    'name': row[1],
                    'middlename': row[2]
                }
                # Объединяем ячейки Фамилия, Имя, Отчество
                # в одну строку 'Фамилия Имя Отчество' для создания папки:
                curr_student = ' '.join(row)
                # Создаём папку для студента:
                try:
                    os.makedirs(f'dir/{curr_student}')
                except FileExistsError as error:
                    QMessageBox.information(
                        self,
                        'Инфо',
                        f'<h4>Студент:<br>{curr_student}<br>'
                        f'<br>Папка для этого студента уже имеется.</h4>'
                    )
                # Выводим фамилию, имя, отчество
                # в таблицу "Список студентов" графического интерфейса:
                for j in range(len(row)):
                    self.tw.setItem(i, j, QTableWidgetItem(row[j]))
            print(studs)
            # Флаг: файл открыт
            self.fileopen = True

            # Сообщение: файл открыт
            msg = QMessageBox.information(self, 'Инфо',
                                          '<h4>Файл со списком студентов открыт.'
                                          '<br>Все папки созданы. Можно продолжить работу.</h4>')
            self.statusBar().showMessage('Сохраните файл в формате DOCX')
            self.pb_save_docx.setDisabled(False)

        # Обработка исключения:
        except FileNotFoundError:
            if self.fileopen and not self.errorOpen:
                message = '<h4>Вы уже открыли файл</h4>'
            else:
                message = '<h4>Вы не открыли файл,<br>попробуйте ещё раз.</h4>'
            QMessageBox.information(self, 'Инфо', message)
            self.errorOpen = True

    def savedocx(self):
        """Сохранение DOCX"""
        # Словарь для рендеринга:
        context = {
            'a': 1,
            'b': 2,
            'c': 3
        }
        # Диалоговое окно сохранения файла docx:
        saveDialog = QFileDialog()
        saveDialog.setDefaultSuffix('docx')
        fname, _ = saveDialog.getSaveFileName(self, 'Сохранить документ', '',
                                              'Microsoft Word 2007–365 (*.docx)')

        if fname != str():
            self.statusBar().showMessage('Идёт процесс формирование документа...')
            # Делаем кнопки неактивными:
            self.pb_open_xls.setDisabled(True)
            self.pb_save_docx.setDisabled(True)
            doc = DocxTemplate('template.dotx')
            doc.render(self.context)
            doc.save(self.fname)

            # Выводим окно QProgressDialog на ожидание рендеринга.
            # HTML-сообщение с иконкой:
            self.save_error = False
            msg = '<table border = "0"> <tbody> <tr>' \
                  '<td> <img src = "pic/save-icon.png"> </td>' \
                  '<td> <h4>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Идёт сохранение документа,<br>' \
                  '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;подождите пожалуйста.</h4> </td>'
            self.dialog = QProgressDialog(msg, None, 0, 0, self)
            self.dialog.setModal(True)
            self.dialog.setWindowTitle('Инфо')
            self.dialog.setRange(0, 0)
            self.dialog.show()
        else:
            QMessageBox.warning(
                self,
                'Внимание!',
                '<h4>Вы не задали имя файла<br>для сохранения.</h4>'
            )

        """
        if s == 'error':
            self.dialog.close()
            self.save_error = True
            msg = QMessageBox.warning(self, 'Внимание!',
                                      '<h4>Не удалось сохранить файл.<br>'
                                      'Возможно, у вас нет доступа<br>к целевой папке.</h4>')
            self.statusBar().showMessage('Не удалось создать файл')
        """

        self.dialog.close()
        if not self.save_error:
            # Выводим информационное сообщение:
            msg = QMessageBox.information(self, 'Инфо',
                                          '<h4>Документ сохранён.</h4>')
            self.statusBar().showMessage('Документ сохранён')
        # Делаем кнопки "Открыть..." и "Сохранить..." активными:
        self.pb_open_xls.setDisabled(False)
        self.pb_save_docx.setDisabled(False)

    def savepdf(self, inputfile, outfolder):
        """Сохранение PDF"""
        cmd = 'libreoffice --convert-to pdf --outdir'.split() + [outfolder] + [inputfile]
        p = subprocess.Popen(cmd, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
        p.wait(timeout=10)
        """
        stdout, stderr = p.communicate()    
        if stderr:
            raise subprocess.SubprocessError(stderr)
        """


def except_hook(cls, exception, traceback):
    """Функция для отслеживания ошибок PyQt5"""
    sys.__excepthook__(cls, exception, traceback)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Students()
    ex.show()
    # Ловим и показываем ошибки PyQt5 в терминале:
    sys.excepthook = except_hook
    sys.exit(app.exec_())
