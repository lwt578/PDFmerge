import os
from PyPDF4 import PdfFileReader, PdfFileWriter
import PySimpleGUI as sg

sg.theme('TanBlue')
layout = [
    [
        sg.Text('请选择文件夹:', text_color="black"),
        sg.Text('                     ',
                key='-folder-',
                background_color="white",
                text_color="black",
                border_width=2),
        sg.FolderBrowse(
            '浏览...',
            target='-path-',
            key='-path-',  # target和key必须都有，且都一样，不知道为啥
            enable_events=True)
    ],
    [sg.Text('', key='-result-', expand_x=True)],
    [
        sg.Button('合并', disabled=True),
        sg.Button('打开文件夹', disabled=True),
        sg.Button('退出')
    ]
]

window = sg.Window('PDF合并工具',
                   layout,
                   auto_size_buttons=True,
                   auto_size_text=True,
                   font='11')

while True:
    event, values = window.read()
    if event in (None, '退出'):
        break

    elif event == '-path-':
        folderpath = values['-path-']
        window['-folder-'].update(folderpath)
        window['合并'].update(disabled=False)
        window['打开文件夹'].update(disabled=False)

    elif event == '合并':

        pdf_lst = [f for f in os.listdir(folderpath) if f.endswith('.pdf')]

        file_writer = PdfFileWriter()
        for i in pdf_lst:
            pdf = PdfFileReader(folderpath + '/' + i)
            for page in range(pdf.getNumPages()):
                file_writer.addPage(pdf.getPage(page))  # 合并pdf文件

        with open(folderpath + '/合并结果.pdf', 'wb') as f:
            file_writer.write(f)

        window['-result-'].update('合并成功！文件保存在：' + folderpath + "/合并结果.pdf")

    elif event == '打开文件夹':
        os.startfile(folderpath)

window.close()