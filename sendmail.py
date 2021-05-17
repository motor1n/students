import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from platform import python_version


class SendLetter:
    def __init__(self, mail_name, student_name, filepath):
        self.recipients = mail_name
        self.sender = 'pwotirn@gmail.com'
        self.password = 'qwertypassworddonotsteal'
        self.server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        self.subject = 'Тема сообщения'
        self.text = 'Текст сообщения' + student_name
        self.html = '<html><head></head><body><p>' + self.text + '</p></body></html>'

        self.msg = MIMEMultipart('alternative')
        self.part_text = MIMEText(self.text, 'plain')
        self.part_html = MIMEText(self.html, 'html')
        self.part_file = MIMEBase('application', 'octet-stream; name="{}"'.format(self.basename))

        self.filepath = filepath
        self.basename = os.path.basename(filepath)
        self.filesize = os.path.getsize(filepath)
        self.create_msg()

    def create_msg(self):
        self.msg['Subject'] = self.subject
        self.msg['From'] = 'Python script <' + self.sender + '>'
        self.msg['To'] = self.recipients
        self.msg['Reply-To'] = self.sender
        self.msg['Return-Path'] = self.sender
        self.msg['X-Mailer'] = 'Python/ ' + (python_version())
        self.part_file.set_payload(open(self.filepath, "rb").read())
        self.part_file.add_header('Content-Description', self.basename)
        self.part_file.add_header(
            'Content-Disposition',
            'attachment; filename="{}"; size={}'.format(self.basename, self.filesize)
        )
        encoders.encode_base64(self.part_file)

        self.msg.attach(self.part_text)
        self.msg.attach(self.part_html)
        self.msg.attach(self.part_file)

        self.send_msg()

    def send_msg(self):
        self.server.login(self.sender, self.password)
        self.server.send_message(self.msg)
        self.server.quit()
