import os
import smtplib
from platform import python_version
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication


class SendLetter:
    def __init__(self, mail_name, student_name, filepath):
        self.recipient = mail_name
        self.sender = 'studentpractic.sgugit@gmail.com'
        self.password = 'vkr-2021'
        self.server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        self.subject = 'СГУГиТ - Пакет документов к производственной практике'
        self.text = f'Уважаемый(ая) {student_name}!\n' \
                    f'Вы получили пакет сформированных документов,' \
                    f'который необходимо приложить к отчетной документации' \
                    f'по прохождению учебной практики.'
        self.html = '<html><head></head><body><p>' + self.text + '</p></body></html>'
        self.filepath = filepath
        self.basename = os.path.basename(self.filepath)
        self.msg = MIMEMultipart('alternative')
        self.part_text = MIMEText(self.text, 'plain')
        self.part_html = MIMEText(self.html, 'html')
        self.attachment = MIMEApplication(open(self.filepath, 'rb').read())
        self.attachment.add_header(
            'Content-Disposition',
            'attachment',
            filename=self.basename
        )
        self.msg.attach(self.attachment)
        self.create_msg()

    def create_msg(self):
        self.msg['Subject'] = self.subject
        self.msg['From'] = 'СГУГиТ - автоматическая рассылка документов <' + self.sender + '>'
        self.msg['To'] = self.recipient
        self.msg['Reply-To'] = self.sender
        self.msg['Return-Path'] = self.sender
        self.msg['X-Mailer'] = 'Python/ ' + (python_version())
        self.msg.attach(self.part_text)
        self.msg.attach(self.part_html)
        self.send_msg()

    def send_msg(self):
        self.server.login(self.sender, self.password)
        self.server.send_message(self.msg)
        self.server.quit()
