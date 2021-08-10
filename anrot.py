import os
import codecs
import html2text

spisok = os.listdir(r'C:\Users\new2r\Desktop\Anuar4k\mag')
spisok.remove('desktop.ini')
from docxtpl import DocxTemplate



for i in spisok:
    file = codecs.open(f'C:\\Users\\new2r\\Desktop\\Anuar4k\\mag\\{i}', 'r', 'utf_8')
    try:
        code = file.read()
    except:
        print(i)
        break
    code2 = html2text.html2text(code)
    code1 = code2.replace('*', '')
    ride = True
    main_content = ''
    end = 0
    for line in code1.split('\n'):
        try:
            line.split(' ')[1]
        except:
            end = 1
        try:
            if line.split(' ')[1] == 'соответствии':
                ride = True
        except:
            continue
        try:
            if line.split(' ')[0] == 'Ректор' or line.split(' ')[0] == 'РЕКТОР' or line.split(' ')[0] == 'РЕКТОР|' or line.split(' ')[0] == 'Проректор':
                ride = False
        except:
            continue
        if ride:
            main_content += '\n' + line
    doc = DocxTemplate("1.docx")
    context = {'date': f"{i.split('Приказ')[0].split('.')[0]}", 'main_content': main_content}
    doc.render(context)
    doc.save(f"{i.split('.html')[0]}.docx")