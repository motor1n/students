import os
import smtplib
import mimetypes
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from email.mime.multipart import MIMEMultipart
from platform import python_version


class SendLetter:
    """
    Отправка письма
    e_mail - электронный адрес получателя
    student - ФИО студента
    files_path - путь к папке с документами
    """
    def __init__(self, e_mail, student, files_path):
        self.recipient = e_mail
        # Оставляем только 'Имя Отчество':
        self.student = student.partition(' ')[-1]
        self.files_path = files_path
        self.sender = 'studentpractic.sgugit@gmail.com'
        self.password = 'vkr-2021'
        self.server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        self.subject = 'СГУГиТ - Пакет документов к учебной практике'
        self.text = f'Здравствуйте, <b>{self.student}</b>!<br>' \
                    f'Данное письмо содержит пакет документов, ' \
                    f'который необходимо приложить к отчетной документации ' \
                    f'по прохождению учебной практики.'
        self.html = '<html><head></head><body><p>' + self.text + '</p></body></html>'
        self.part_html = MIMEText(self.html, 'html')

    def send_email(self):
        """
        Отправка письма
        attach_files - файлы для вложения в письмо
        """
        # Создаём экземпляр сообщения:
        msg = MIMEMultipart()
        # Заполняем поля сообщения:
        msg['From'] = 'СГУГиТ - автоматическая рассылка документов <' + self.sender + '>'
        msg['To'] = self.recipient
        msg['Subject'] = self.subject
        msg['Reply-To'] = self.sender
        msg['Return-Path'] = self.sender
        msg['X-Mailer'] = 'Python/ ' + (python_version())
        # Добавляем сообщение:
        msg.attach(self.part_html)
        # Добавляем в сообщение все файлы:
        process_attachement(msg, self.files_path)
        # Создаём объект для работы с SMTP-сервером:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        # Включаем режим отладки, если надо:
        # server.set_debuglevel(True)
        # Получаем доступ к SMTP:
        server.login(self.sender, self.password)
        # Отправляем сообщение:
        server.send_message(msg)
        # Выходим:
        server.quit()


def process_attachement(msg, attach_path):
    """
    Добавление файлов к сообщению
    msg - сообщение
    attach_path - путь к файлу
    """
    for f in attach_path:
        # Если файл существует:
        if os.path.isfile(f):
            # Добавляем файл к сообщению:
            attach_file(msg, f)
        # Если путь - не файл и существует, значит - папка:
        elif os.path.exists(f):
            # Получаем список файлов в папке:
            dir_list = os.listdir(f)
            # Перебираем все файлы:
            for file in dir_list:
                # Добавляем файл к сообщению:
                attach_file(msg, f + '/' + file)


def attach_file(msg, filepath):
    """Добавление файла в сообщение"""
    filename = os.path.basename(filepath)
    # Определяем тип файла на основе его расширения
    ctype, encoding = mimetypes.guess_type(filepath)
    # Если тип файла не определяется:
    if ctype is None or encoding is not None:
        # Будем использовать общий тип:
        ctype = 'application/octet-stream'
    # Получаем тип и подтип:
    maintype, subtype = ctype.split('/', 1)
    # Если текстовый файл:
    if maintype == 'text':
        # Открываем файл для чтения
        with open(filepath) as fp:
            # Используем тип MIMEText:
            file = MIMEText(fp.read(), _subtype=subtype)
            # После использования файл обязательно нужно закрыть:
            fp.close()
    # Если изображение:
    elif maintype == 'image':
        with open(filepath, 'rb') as fp:
            file = MIMEImage(fp.read(), _subtype=subtype)
            fp.close()
    # Если аудио:
    elif maintype == 'audio':
        with open(filepath, 'rb') as fp:
            file = MIMEAudio(fp.read(), _subtype=subtype)
            fp.close()
    # Неизвестный тип файла:
    else:
        with open(filepath, 'rb') as fp:
            # Используем общий MIME-тип:
            file = MIMEBase(maintype, subtype)
            # Добавляем содержимое общего типа (полезную нагрузку):
            file.set_payload(fp.read())
            fp.close()
            # Содержимое должно кодироваться как Base64:
            encoders.encode_base64(file)
    # Добавляем заголовки:
    file.add_header(
        'Content-Disposition',
        'attachment',
        filename=filename
    )
    # Присоединяем файл к сообщению:
    msg.attach(file)
