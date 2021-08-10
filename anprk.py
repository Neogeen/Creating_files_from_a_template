from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, TimeoutException, UnexpectedAlertPresentException
from selenium.webdriver.firefox.firefox_binary import FirefoxBinary
from selenium.webdriver.firefox.options import Options as FirefoxOptions
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import os
import codecs
import html2text

spisok = os.listdir(r'C:\Users\new2r\Desktop\Anuar4k\mag')
from docxtpl import DocxTemplate



for i in spisok:
    file = codecs.open(f'C:\\Users\\new2r\\Desktop\\Anuar4k\\mag\\{i}', 'r', 'utf_8')
    code = file.read()
    code2 = html2text.html2text(code)
    code1 = code2.replace('*', '')
    ride = True
    main_content = ''
    for line in code1.split('\n'):
        try:
            line.split(' ')[1]
        except:
            print('')
        try:
            if line.split(' ')[1] == 'соответствии':
                ride = True
        except:
            continue
        try:
            if line.split(' ')[0] == 'Ректор':
                ride = False
        except:
            continue
        if ride:
            main_content += '\n' + line
    print(main_content)
    doc = DocxTemplate("1.docx")
    context = {'date': f"{i.split('Приказ')[0]}", 'main_content': main_content}
    doc.render(context)
    doc.save(f"{i.split('.html')[0]}.docx")
