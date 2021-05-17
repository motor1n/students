"""Students 0.8.0"""

import os
import sys
import xlrd
from PyQt5 import uic
from datetime import datetime, timedelta
from sendmail import SendLetter
from savepdf import ToPDF
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

# Словарь соответствия названия шаблона и файла шаблона:
tpl_file = {
    'Заявление о направлении на практику': 'tpl_app_for_practice.dotx',
    'Рабочий график (план) проведения практики': 'tpl_work_shedule.dotx',
    'Индивидуальное задание на практику': 'tpl_individual_task.dotx',
    'Инструктаж по технике безопасности': 'tpl_check_list.dotx'
}


class Students(QMainWindow):
    def __init__(self):
        super().__init__(None)
        # Загрузка интерфейса:
        uic.loadUi('students.ui', self)

        # Файл ещё не открыт:
        self.fileopen = False

        # Ошибок открытия файла ещё не было:
        self.errorOpen = False

        # Ошибок сохранения файлов ещё не было:
        self.save_error = False

        # Изначально пустой словарь для рендеринга:
        self.context = dict()

        # Изначально текущий тип шаблона не выбран:
        self.curr_tpl = '---'

        # Изначально текущий файл шаблона не определён:
        self.curr_file = str()

        # Кнопки pb_save_docx и pb_save_pdf дезактивированы,
        # поскольку на данный момент ещё ничего не сделано:
        self.pb_open_xls.setDisabled(True)
        self.pb_save_docx.setDisabled(True)
        self.pb_save_pdf.setDisabled(True)
        self.pb_send_email.setDisabled(True)

        # Сигнал отслеживания изменения QComboBox при выборе типа шаблона:
        self.cb1.currentTextChanged.connect(self.tpl_select)

        # Открыть файл XLS. Сигнал pb_open_xls --> слот open_xls:
        self.pb_open_xls.clicked.connect(self.open_xls)

        # Сохранить файл DOCX. Сигнал pb_save_docx --> слот savedocx:
        self.pb_save_docx.clicked.connect(self.savedocx)
        # self.pb_save_docx.clicked.connect(lambda checked, data=self.context: self.savedocx(data))

        # Сохранить пакеты документов PDF. Сигнал pb_save_pdf --> слот savepacks:
        self.pb_save_pdf.clicked.connect(self.savepacks)

        # Отправить письмо на почту. Сигнал pb_send_email --> слот sendingmail:
        self.pb_send_email.clicked.connect(self.sendingmail)

        self.statusBar().showMessage('Изучите инструкцию и приступайте к работе')

    def tpl_select(self):
        """Выбор типа шаблона"""
        if self.cb1.currentText() != '---':
            # Если файл уже открыт,
            # активируем кнопку сохранения:
            if self.fileopen:
                self.pb_save_docx.setDisabled(False)
                msg = 'Сохраните файл в формате DOCX'
                self.statusBar().showMessage(msg)
            # Задаём текущий тип шаблона:
            self.curr_tpl = self.cb1.currentText()
            # Задаём текущий файл шаблона:
            self.curr_file = tpl_file[self.curr_tpl]
            # Если тип шаблона выбран, активируем кнопку "Открыть файл XLS..."
            self.pb_open_xls.setDisabled(False)
            if self.fileopen:
                msg = 'Сохраните файл в формате DOCX'
            else:
                msg = 'Откройте файл с исходными данными'
            self.statusBar().showMessage(msg)
        else:
            # Дезактивируем все кнопки:
            self.pb_open_xls.setDisabled(True)
            self.pb_save_docx.setDisabled(True)
            self.pb_save_pdf.setDisabled(True)
            self.pb_send_email.setDisabled(True)
            msg = 'Выберите шаблон документа.'
            self.statusBar().showMessage(msg)

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
            for i in range(2, sh.nrows):
                # Считываем строку из таблицы:
                row = sh.row_values(i)

                # Фамилия Имя Отчество текущего студента:
                curr_student = row[0]

                # Словарь для рендеринга, читаем из XLS-файла:
                self.context = {
                    'student': curr_student,
                    'course': row[1],
                    'group': row[2],
                    'forms': row[3],
                    'phone': row[4],
                    'mail': row[5],
                    'hh': row[6],
                    'date': date_conv(row[7], workbook),
                    'date1': date_conv(row[8], workbook),
                    'date2': date_conv(row[9], workbook),
                    'chief': row[10],
                    'place': row[11],
                    'location': row[12],
                    'address': row[13],
                    'teacher': row[14],
                    'yy': row[15],
                    'date3': date_conv(row[16], workbook),
                    'date4': date_conv(row[17], workbook),
                    'date5': date_conv(row[18], workbook),
                    'date6': date_conv(row[19], workbook),
                    'date7': date_conv(row[20], workbook),
                    'date8': date_conv(row[21], workbook),
                    'date9': date_conv(row[22], workbook),
                    'date10': date_conv(row[23], workbook)
                }

                # print(self.context)

                # Помещаем словарь (данные по студенту) в глобальный список studs:
                studs.append(self.context)

                # Выводим фамилию, имя, отчество
                # в таблицу "Список студентов" графического интерфейса:
                for j in range(5):
                    self.tw.setItem(i - 2, j, QTableWidgetItem(row[j]))

                # Автоподбор ширины столбцов:
                self.tw.resizeColumnsToContents()

                # Создаём папку для студента:
                """
                try:
                    os.makedirs(f'dir/{curr_student}')
                except FileExistsError as error:
                    QMessageBox.information(
                        self,
                        'Инфо',
                        f'<h4>Студент:<br>{curr_student}<br>'
                        f'<br>Папка для этого студента уже имеется.</h4>'
                    )
                """
            # Флаг: файл открыт
            self.fileopen = True

            # Сообщение: файл открыт
            msg = QMessageBox.information(self, 'Инфо',
                                          '<h4>Файл со списком студентов открыт.'
                                          '<br>Можно продолжить дальнейшую работу.</h4>')
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

        # Окно диалога выбора папки для сохранения файлов:
        self.docdir = QFileDialog.getExistingDirectory(self, 'Выбрать папку', '.')

        if self.docdir != str():
            self.statusBar().showMessage('Идёт процесс формирование документов...')
            # Делаем кнопки неактивными:
            self.pb_open_xls.setDisabled(True)
            self.pb_save_docx.setDisabled(True)

            # Задаём папку для группы с названием шаблона:
            folder = f"{self.context['group']} - {self.curr_tpl}"
            if not os.path.isdir(f'{self.docdir}/{folder}'):
                os.mkdir(f'{self.docdir}/{folder}')

            # Создаём документы для всех студентов из списка:
            for s in studs:
                # packdoc_dir = f"Группа {self.context['group']} - Пакеты документов на практику"
                # pdf_dir = f"{self.docdir}/{packdoc_dir}/{s['student']}"
                filedoc = f"{self.docdir}/{folder}/{s['student']} - {self.curr_tpl}.docx"
                doc = DocxTemplate(f'tpl/{self.curr_file}')
                doc.render(s)
                doc.save(filedoc)
                # Конвертируем файл DOCX в PDF:
                # ToPDF(pdf_dir).doc2pdf(filedoc)

            if not self.save_error:
                # Выводим информационное сообщение:
                msg = QMessageBox.information(self, 'Инфо',
                                              '<h4>Документы сохранены.</h4>')
                self.statusBar().showMessage('Документы сохранены')

            # Делаем кнопки активными:
            self.pb_open_xls.setDisabled(False)
            self.pb_save_docx.setDisabled(False)
            self.pb_save_pdf.setDisabled(False)
            self.pb_send_email.setDisabled(False)
        else:
            QMessageBox.warning(
                self,
                'Внимание!',
                '<h4>Вы не выбрали папку<br>для сохранения.</h4>'
            )

        """
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
        
        if s == 'error':
            self.dialog.close()
            self.save_error = True
            msg = QMessageBox.warning(self, 'Внимание!',
                                      '<h4>Не удалось сохранить файл.<br>'
                                      'Возможно, у вас нет доступа<br>к целевой папке.</h4>')
            self.statusBar().showMessage('Не удалось создать файл')
            
        self.dialog.close()       
        """

    def savepacks(self):
        """Сохранение пакетов документов"""
        for s in studs:
            packdoc_dir = f"Группа {self.context['group']} - Пакеты документов на практику"
            pdf_dir = f"{self.docdir}/{packdoc_dir}/{s['student']}"
            # filedoc = f"{self.docdir}/{folder}/{s['student']} - {self.curr_tpl}.docx"
            # Конвертируем файл DOCX в PDF:
            # ToPDF(pdf_dir).doc2pdf(filedoc)

    @staticmethod
    def sendingmail():
        """Отправка писем"""
        for i in studs:
            print(i['mail'])
            # SendLetter(i['mail'], i['student'], 'students.xls')
        # Передача параметров классу SendLetter, который генерирует и отправляет письма
        # SendLetter('chmferks@gmail.com', 'student fio', 'students.xls')


def date_conv(xldate, book):
    """Конвертация даты из ячейки таблицы в нормальный формат"""
    time = datetime(*xlrd.xldate_as_tuple(xldate, book.datemode))
    return time.strftime('%d.%m.%Y')


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
