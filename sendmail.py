import smtplib
import os
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from platform import python_version


class SendLetter:
    def __init__(self, mail_name, student_name, filepath):
        self.recipient = mail_name
        self.sender = 'pwotirn@gmail.com'
        self.password = 'qwertypassworddonotsteal'
        self.server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        self.subject = 'Тема сообщения'
        self.text = 'Текст сообщения ' + student_name
        self.html = '<html><head></head><body><p>' + self.text + '</p></body></html>'
        self.filepaths = filepath
        self.basenames = list()
        for i in filepath:
            self.basenames.append(os.path.basename(i))
        self.msg = MIMEMultipart('alternative')
        self.part_text = MIMEText(self.text, 'plain')
        self.part_html = MIMEText(self.html, 'html')
        for num, f in enumerate(self.basenames):
            self.filepath = self.filepaths[num]
            self.attachment = MIMEApplication(open(self.filepath, "rb").read())
            self.attachment.add_header('Content-Disposition', 'attachment', filename=f)
            self.msg.attach(self.attachment)
        self.create_msg()

    def create_msg(self):
        self.msg['Subject'] = self.subject
        self.msg['From'] = 'Python script <' + self.sender + '>'
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
