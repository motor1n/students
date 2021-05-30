"""Student Practice 0.9.1"""

import os
import sys
import xlrd
import glob
from zipfile import ZipFile
from PyQt5 import uic
from datetime import datetime
from PyQt5.QtCore import Qt
from sendmail import SendLetter
from thread import ThreadPDF, ThreadDOCX
from PyQt5.QtCore import QThread
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

        # Изначально множество выбранных названий шаблонов пустое:
        self.curr_tpls = set()

        # Изначально множество выбранных файлов шаблонов пустое:
        self.curr_files = set()

        # Папка сохранения документов DOCX:
        self.docdir = str()

        # Текущая папка сохранения документов PDF для рассылки e-mail:
        self.curr_packdocs = str()

        # Словарь путей сохранённых документов пользователей:
        self.docpaths = dict()

        # Сообщение о результате операции:
        self.msg = str()

        # Максимальное значение QProgressDialog:
        self.max_value = 0

        # Кнопки pb_save_docx и pb_save_pdf дезактивированы,
        # поскольку на данный момент ещё ничего не сделано:
        self.pb_open_xls.setDisabled(True)
        self.pb_save_docx.setDisabled(True)
        self.pb_save_pdf.setDisabled(True)
        self.pb_send_email.setDisabled(True)

        # Группа кнопок для сигнала отслеживания изменения CheckBox при выборе типа шаблона:
        self.bg.buttonClicked.connect(self.tpl_select)

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

    def tpl_select(self, checkbox):
        """Выбор шаблонов"""
        # Если хотя бы один чекбокс активирован:
        if any([cb.isChecked() for cb in self.bg.buttons()]):
            # Если файл уже открыт,
            # активируем кнопку сохранения:
            if self.fileopen:
                self.pb_save_docx.setDisabled(False)
                msg = 'Сохраните файл в формате DOCX'
                self.statusBar().showMessage(msg)

            # Выбираем отмеченные шаблоны:
            self.curr_tpls = {cb.text() for cb in self.bg.buttons() if cb.isChecked()}

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
            msg = 'Выберите шаблоны документов.'
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
                    'hh': int(row[6]),
                    'date': date_conv(row[7], workbook),
                    'date1': date_conv(row[8], workbook),
                    'date2': date_conv(row[9], workbook),
                    'chief': row[10],
                    'place': row[11],
                    'location': row[12],
                    'address': row[13],
                    'teacher': row[14],
                    'yy': int(row[15]),
                    'date3': date_conv(row[16], workbook),
                    'date4': date_conv(row[17], workbook),
                    'date5': date_conv(row[18], workbook),
                    'date6': date_conv(row[19], workbook),
                    'date7': date_conv(row[20], workbook),
                    'date8': date_conv(row[21], workbook),
                    'date9': date_conv(row[22], workbook),
                    'date10': date_conv(row[23], workbook)
                }

                # Помещаем словарь (данные по студенту) в глобальный список studs:
                studs.append(self.context)

                # Выводим исходные данные
                # в таблицу "Список студентов" графического интерфейса:
                column = 0
                for value in self.context.values():
                    if column > 23:
                        break
                    self.tw.setItem(i - 2, column, QTableWidgetItem(str(value)))
                    column += 1

                # Автоподбор ширины столбцов:
                self.tw.resizeColumnsToContents()

            # Флаг: файл открыт
            self.fileopen = True

            # Сообщение: файл открыт
            msg = QMessageBox.information(self, 'Инфо',
                                          '<h4>Файл со списком студентов открыт.'
                                          '<br>Можно продолжить дальнейшую работу.</h4>')
            self.statusBar().showMessage('Сохраните документы в формате DOCX')
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

        self.msg = '<table border = "0"> <tbody> <tr>' \
                   '<td> <img src = "pic/save-icon.png"> </td>' \
                   '<td> <h4>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Идёт процесс формирование документов,' \
                   '<br>' \
                   '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;подождите пожалуйста.</h4> </td>'

        self.max_value = len(studs) * len(self.curr_tpls)

        if self.docdir != str():
            self.statusBar().showMessage('Идёт процесс формирование документов...')
            # Делаем кнопки неактивными:
            self.pb_open_xls.setDisabled(True)
            self.pb_save_docx.setDisabled(True)

            # Выбираем отмеченные шаблоны:
            self.curr_tpls = {cb.text() for cb in self.bg.buttons() if cb.isChecked()}

            self.thread1 = ThreadDOCX(
                studs,
                tpl_file,
                self.curr_tpls,
                self.context,
                self.docdir,
                self.docpaths
            )

            # Сигнал запуска потока thread отправляем на слот thread_start:
            self.thread1.started.connect(self.thread_start)

            # Qt.QueuedConnection - сигнал помещается в очередь обработки событий интерфейса Qt:
            self.thread1.signal.connect(self.thread_process, Qt.QueuedConnection)

            # Сигнал завершения потока thread отправляем на слот thread_stop:
            self.thread1.finished.connect(self.thread_stop)

            # Запускаем поток рендеринга:
            self.thread1.start(priority=QThread.IdlePriority)

            # Делаем кнопки активными:
            self.pb_open_xls.setDisabled(False)
            self.pb_save_docx.setDisabled(False)
            self.pb_save_pdf.setDisabled(False)
        else:
            QMessageBox.warning(
                self,
                'Внимание!',
                '<h4>Вы не выбрали папку<br>для сохранения.</h4>'
            )

    def savepacks(self):
        """Сохранение пакетов PDF-документов"""
        # Окно диалога выбора папки для сохранения пакетов документов:
        pdfdir = QFileDialog.getExistingDirectory(self, 'Выбрать папку', '.')

        # Папка для пакетов документов текущей группы:
        packdoc_dir = f"Группа {self.context['group']} - Пакеты документов на практику"

        # Текущая папка сохранения пакетов документов,
        # из которой будут рассылаться электронные письма:
        self.curr_packdocs = f'{pdfdir}/{packdoc_dir} - PDF'

        self.statusBar().showMessage('Идёт формирование пакетов документов...')

        self.msg = '<table border = "0"> <tbody> <tr>' \
                   '<td> <img src = "pic/save-icon.png"> </td>' \
                   '<td> <h4>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Идёт формирование пакетов документов,' \
                   '<br>' \
                   '&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;подождите пожалуйста.</h4> </td>'

        self.max_value = len(self.docpaths)

        self.thread2 = ThreadPDF(self.docpaths.items(), self.curr_packdocs)

        # Сигнал запуска потока thread отправляем на слот thread_start:
        self.thread2.started.connect(self.thread_start)

        # Qt.QueuedConnection - сигнал помещается в очередь обработки событий интерфейса Qt:
        self.thread2.signal.connect(self.thread_process, Qt.QueuedConnection)

        # Сигнал завершения потока thread отправляем на слот thread_stop:
        self.thread2.finished.connect(self.thread_stop)

        # Запускаем поток рендеринга:
        self.thread2.start(priority=QThread.IdlePriority)

        # Делаем кнопку отправки e-mail активной:
        # self.pb_send_email.setDisabled(False)

    def sendingmail(self):
        """Отправка писем"""
        self.statusBar().showMessage('Идёт рассылка электронных писем')
        for stud in studs:
            # Формируем список с путями к файлам по ФИО студентов и группе:
            tmp = glob.glob(f"{self.curr_packdocs}/{stud['student']}*")
            print(tmp)
            # Создаем zip-файл со всеми pdf файлами студента:
            with ZipFile(
                    stud['student'] + ' - Пакеты документов на практику' + '.zip',
                    'w'
            ) as zipobj:
                for foldername, subfolders, filenames in os.walk(''.join(tmp).replace('\\', '/')):
                    for filename in filenames:
                        filepath = os.path.join(foldername, filename)
                        zipobj.write(filepath, os.path.basename(filepath))
            # Передача параметров классу SendLetter, который генерирует и отправляет письма:
            SendLetter(
                stud['mail'],
                stud['student'],
                ''.join(glob.glob(f'{stud["student"]} - Пакеты документов на практику.zip'))
            )
            # Удаление отправленного файла:
            os.remove(stud['student'] + ' - Пакеты документов на практику' + '.zip')
        self.statusBar().showMessage('Письма отправлены')

    def thread_start(self):
        """Вызывается при событии запуска потока"""
        self.dialog = QProgressDialog(self.msg, None, 0, 0, self)
        self.dialog.setModal(True)
        self.dialog.setWindowTitle('Инфо')
        self.dialog.setRange(0, self.max_value)
        self.dialog.show()

    def thread_process(self, val):
        """Вызывается сигналами которые отправляет поток"""
        # Параметр val - это сигнал полученный из потока thread
        if val == 'error':
            self.dialog.close()
            self.save_error = True
            msg = QMessageBox.warning(self, 'Внимание!',
                                      '<h4>Ошибка выполнеия операции.</h4>')
            self.statusBar().showMessage('Ошибка выполнения операции')
        else:
            # Изменение процента выполнения процесса:
            self.dialog.setValue(val)

    def thread_stop(self):
        """Вызывается при событии завершения потока"""
        self.dialog.close()
        # Выводим информационное сообщение:
        msg = QMessageBox.information(
            self,
            'Инфо',
            '<h4>Операция успешно завершена.</h4>'
        )
        self.statusBar().showMessage('Операция успешно завершена.')


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
