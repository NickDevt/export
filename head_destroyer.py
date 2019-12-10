import lxml
import docx
import os
from docx import Document
from docx.shared import Inches

# https://github.com/python-openxml/python-docx/issues/33
# https://ru.stackoverflow.com/questions/1056841/%d0%a3%d0%b4%d0%b0%d0%bb%d0%b5%d0%bd%d0%b8%d0%b5-%d1%87%d0%b0%d1%81%d1%82%d1%8c-%d1%81%d0%be%d0%b4%d0%b5%d1%80%d0%b6%d0%b8%d0%bc%d0%be%d0%b3%d0%be-docx-%d1%81-%d0%bf%d0%be%d0%bc%d0%be%d1%89%d1%8c%d1%8e-python?noredirect=1#comment1807402_1056841


doc = docx.Document('библио/Педагогическое образование библиотечное дело в школьных учреждениях , 512 часов.docx')
para = doc.paragraphs[:4]
for i in para:
    i.text = None
# doc.save('библио/сука.docx')


three = os.listdir('C:/Users/ma-tarasov/Desktop/Задачи ---проекты конкретные/таблицы')
# for i in three:
#     print(i)



main_directory = ['C:/Users/ma-tarasov/Desktop/Задачи ---проекты конкретные/таблицы/ПП', 'C:/Users/ma-tarasov/Desktop/Задачи ---проекты конкретные/таблицы/ПК']


def get_folders_list():
    pp = main_directory[0]

    pk = main_directory[1]
    pp_folders = os.listdir(pp)
    pk_folders = os.listdir(pk)
    for i in pp_folders:
        if '.' not in i:
            list1 = ['C:/Users/ma-tarasov/Desktop/Задачи ---проекты конкретные/таблицы/ПП/' + i]
            for i in list1:
                xfile = os.listdir(i)
                for i in xfile:
                    final = os.path.abspath(i)
                    # for i in final:
                    print(pp)




get_folders_list()


def head_destroy():
    pass























# def delete_paragraph(paragraph):
#     p = paragraph._element
#     p.getparent().remove(p)
#     p._p = p._element = None
#
# delete_paragraph(para)


