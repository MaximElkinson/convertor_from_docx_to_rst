from docx import Document as Document_for_reading
from datetime import datetime
import zipfile, pathlib

import queue
from spire.doc import *
from spire.doc.common import *
from bs4 import BeautifulSoup
import xml.etree.ElementTree as ET




def extract_xml(docx):
    document = zipfile.ZipFile(docx)
    soup = str(BeautifulSoup(document.read('word/document.xml'), 'html.parser')).split('<w:t>')
    rst = ''
    pn = 0
    for i in soup[1:]:
        text = i.split('</w:t>')
        rst += text[0] + '\n\n'
        if 'Рисунок ' in text[1]:
            for _ in text[1].split('Рисунок ')[1:]:
                rst += f'.. image:: image{pn}.png' + '\n\n'
                pn += 1
    with open('index.rst', 'w', encoding='utf-8') as f:
        f.write(rst)


extract_xml('admin.docx')

