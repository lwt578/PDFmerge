import os
import re
from pdf2docx import Converter
import PySimpleGUI as sg
import pdfplumber
from openpyxl import Workbook #保存表格，需要安装openpyxl


sg.theme('Default')
layout = [
    [
        sg.Text('请选择文件:', text_color="black"),
        sg.Text('',
                key='-file-',
                background_color="white",
                text_color="black",
                size=(20,None)
                ),
        sg.FileBrowse(
            '浏览...',
            target='-filename-',
            key='-filename-',  # target和key必须都有，且都一样，不知道为啥
            file_types=[("pdf文件","*.pdf")],
            enable_events=True)
    ],
    [sg.Text('', key='-result-', expand_x=True)],
    [
        sg.Button('转换为Word', disabled=True),
        sg.Button('转换为Excel', disabled=True),
        sg.Button('打开文件夹', disabled=True),
        sg.Button('退出')
    ]
]

window = sg.Window('PDF转换工具',
                   layout,
                   auto_size_buttons=True,
                   auto_size_text=True,
                   font='11')

while True:
    event, values = window.read()
    if event in (None, '退出'):
        break

    elif event == '-filename-':
        pdf_adress = values['-filename-']
        folderpath=os.getcwd()
        pdf_file=pdf_adress
        filename=re.split("[/.]",pdf_adress)[-2]
        window['-file-'].update(pdf_adress)
        window['转换为Word'].update(disabled=False)
        window['转换为Excel'].update(disabled=False)
        

    elif event == '转换为Word':
        
        window['-result-'].update('转换中……')      

        cv = Converter(pdf_file)
        cv.convert(folderpath + '\\' + filename + '.docx')
        cv.close()
        
        sg.popup('转换成功！文件保存在：' + folderpath + '\\' + filename + '.docx')
        window['-result-'].update('转换成功')
        window['打开文件夹'].update(disabled=False)
    
    elif event == '转换为Excel':

        window['-result-'].update('转换中……')

        workbook = Workbook()
        sheet = workbook.active
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                for row in table:
                    sheet.append(row)
        workbook.save(filename=folderpath + "\\" + filename + ".xlsx")
        
        sg.popup('转换成功！文件保存在：' + folderpath + '\\' + filename + ".xlsx")
        window['-result-'].update('转换成功')
        window['打开文件夹'].update(disabled=False)
    


    elif event == '打开文件夹':
        os.startfile(folderpath)

window.close()

