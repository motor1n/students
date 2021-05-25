"""Конвертация DOCX в PDF"""

import os
import subprocess


class ToPDF:
    def __init__(self, path):
        # Путь для сконвертированных файлов
        self.path = path

    def doc2pdf(self, doc):
        """
        Конвертация в OS Windows
        doc - путь к документу
        """
        try:
            from comtypes import client
        except ImportError:
            client = None

        doc = os.path.abspath(doc)
        if client is None:
            return self.doc2pdf_linux(doc)
        name, ext = os.path.splitext(doc)

        try:
            word = client.CreateObject('Word.Application')
            worddoc = word.Documents.Open(doc)
            worddoc.SaveAs(name + '.pdf', FileFormat=17)
        except Exception:
            raise
        finally:
            worddoc.Close()
            word.Quit()

    def doc2pdf_linux(self, doc):
        """
        Конвертация в OS Linux
        Требуется установеленный LibreOffice
        doc - путь к документу
        """
        # cmd = 'libreoffice --convert-to pdf'.split() + [doc]
        cmd = 'libreoffice --headless --convert-to pdf'.split() + [doc, '--outdir', self.path]
        p = subprocess.Popen(cmd, stderr=subprocess.PIPE, stdout=subprocess.PIPE)
        p.wait(timeout=10)
        """
        stdout, stderr = p.communicate()
        if stderr:
            raise subprocess.SubprocessError(stderr)
        """
