'''
时间：3.15-3.21
Version:2.0
单词录入系统
'''
import csv
import openpyxl
import shutil
from translate import Translator
import PySimpleGUI as sg
from tkinter import *
import tkinter.messagebox
import os
import xlrd
import xlwt
from bs4 import BeautifulSoup
from xlutils.copy import copy
import webbrowser
import sys
import pandas as pd
import win32com.client as win32

try:
    deletenum=False
    numfile=open("numfile.log",'r')
    num=int(numfile.readline())
    numfile.close()
except FileNotFoundError:
    numfile = open("numfile.log", 'a+')
    numfile.write("0")
    numfile.close()

text1="自动导入excel"
try:
    os.mkdir("data")
except FileExistsError:
    pass

#合并表格
def hebing_excel(dir):
    filename_excel = []
    frames = []
    d = dir.replace('/','\\\\')
    if d.endswith('\\\\') == False:
        d = d + '\\\\'
    print("路径是：",d,"\n有以下文件：")
    for files in os.listdir(path=dir):
        print(files)
        if 'xlsx' in files or 'xls' in files :
            filename_excel.append(files)
            df = pd.read_excel(d+files)
            frames.append(df)
    if len(frames)!= 0:
        result = pd.concat(frames)
        result.to_excel(d+"合并后结果.xlsx")
        sg.popup_ok("success!")

#网络查词
def ch_to_en():
    global flang
    global tlang
    flang="chinese"
    tlang="english"
def en_to_ch():
    global flang
    global tlang
    flang="english"
    tlang="chinese"
#关于与GitHub
def about():
    tkinter.messagebox.showinfo("关于","岳阳市第十五中学423班 梁家诺制作\nVersion:2.0\n支持从英汉/汉英翻译词典或翻译中调入单词数据并自动输入到excel中")
def GitHub():#链接到GitHub
    webbrowser.open('https://github.com/lele20050226/word-entry')
#加载本地词条
def loading():
    try:
        local_dict = {}
        with open(path,'r',encoding="gbk") as file:
            for line in file:
                line = line.rstrip().split(',')
                local_dict[line[0]] = line[1]
        return local_dict
    except NameError:
        select_dictionary()
#在线查词
def net_search(word):
    global num
    global v
    try:
        translator = Translator(from_lang=flang, to_lang=tlang)
    except NameError:
        tkinter.messagebox.showerror("错误","还未设置在线翻译模式！")
    ans = translator.translate(word)
    string = '在线查询结果：\n'+word+':'+ans+'\n'
    if int(v.get()) == 1:
        rbook = xlrd.open_workbook(excel_path, formatting_info=True)  # 打开文件
        w_sheet = rbook.sheet_by_index(0)
        try:
            if word in w_sheet.col_values(1):
                # 如果录入过就提示
                tkinter.messagebox.showwarning("提示", "这个单词已经录入过了！")
            else:
                num = num + 1
                xh = num + 1
                wbook = copy(rbook)  # 复制文件并保留格式
                w_sheet = wbook.get_sheet(0)  # 索引sheet表
                w_sheet.write(num, 0, label=num)
                w_sheet.write(num, 1, label=word)
                w_sheet.write(num, 2, label=ans)
                wbook.save(excel_path)  # 保存文件
        except PermissionError:
            tkinter.messagebox.showerror("错误", "请关闭Excel软件后重试！")
        return string
    else:
        return string
#选择字典、表格和合并表格路径
def select_dictionary():
    global path
    path = sg.popup_get_file("请选择词典文件：",file_types=(("词典文件(.csv)",".csv"),))
def select_excel():
    global excel_path
    excel_path = sg.popup_get_file("请选择表格文件：",file_types=(("表格文件",".xls"),("表格文件（结果请在目录下words.xls搜寻）",".xlsx")))
def hebing():
    dir = sg.popup_get_folder("请选择表格目录：")
    hebing_excel(dir)
#查找单词
def search(word,local_dict):
    count=0
    global intocsv
    global num
    global deletenum
    global new_mean
    result='本地部分匹配结果：\n'
    for key in local_dict.keys():
        if word.upper()==key.upper():#大小写不敏感
            if int(v.get())==1:
                try:
                    if os.path.splitext(excel_path)[1]==".xls":
                        rbook = xlrd.open_workbook(excel_path,formatting_info=True)#打开文件
                        w_sheet = rbook.sheet_by_index(0)
                        try:
                            if key in w_sheet.col_values(1):
                                #如果录入过就提示
                                tkinter.messagebox.showwarning("提示","这个单词已经录入过了！")
                            else:
                                num=num+1
                                xh=num+1
                                wbook = copy(rbook)#复制文件并保留格式
                                w_sheet = wbook.get_sheet(0)#索引sheet表
                                w_sheet.write(num,0,label=num)
                                w_sheet.write(num,1,label=key)
                                w_sheet.write(num,2,label=local_dict.get(key))
                                wbook.save(excel_path)#保存文件
                        except PermissionError:
                            tkinter.messagebox.showerror("错误", "请关闭Excel软件后重试！")
                    elif os.path.splitext(excel_path)[1]==".xlsx":
                        rbook = xlrd.open_workbook("words.xls", formatting_info=True)  # 打开文件
                        w_sheet = rbook.sheet_by_index(0)
                        try:
                            if key in w_sheet.col_values(1):
                                # 如果录入过就提示
                                tkinter.messagebox.showwarning("提示", "这个单词已经录入过了！")
                            else:
                                num = num + 1
                                xh = num + 1
                                wbook = copy(rbook)  # 复制文件并保留格式
                                w_sheet = wbook.get_sheet(0)  # 索引sheet表
                                w_sheet.write(num, 0, label=num)
                                w_sheet.write(num, 1, label=key)
                                w_sheet.write(num, 2, label=local_dict.get(key))
                                wbook.save("words.xls")  # 保存文件
                        except PermissionError:
                            tkinter.messagebox.showerror("错误", "请关闭Excel软件后重试！")

                        num = 0
                        xh = 1
                        workbook = xlwt.Workbook(encoding='utf-8')
                        worksheet = workbook.add_sheet('words')
                        worksheet.write(num, 0, label='序号')
                        worksheet.write(num, 1, label='单词')
                        worksheet.write(num, 2, label='释义')
                        workbook.save('words.xls')

                except NameError:
                    select_excel()
            else:
                pass
            numfile1=open("numfile.log",'w')
            numfile1.write(str(num))
            numfile1.close()
            numfile2=open("numfile.log",'r')
            num=int(numfile2.readline())
            numfile2.close()
            return '本地查询结果：\n'+key+': '+local_dict.get(key)+'\n'
        elif word.upper() in key.upper():
            result=result+key+': '+local_dict.get(key)+'\n'
            count=count+1
    if count==0:
        return net_search(word)+'-'*45+'\n本地词库匹配无结果\n'
    else:
        return net_search(word)+'-'*45+'\n'+result
def delete_num():
    numfile=open("numfile.log",'w')
    numfile.write('0')
    numfile.close()
    num=0
    xh=1

def search_word():
    word = entry.get().strip()
    if len(word) != 0:
        local_dict=loading()
        result = search(word,local_dict)
        doc.delete(1.0, 'end')
        doc.insert('end', result)
    else:
        doc.delete(1.0, 'end')
        doc.insert('end', "请输入查询词条,按回车或点击查询...\n")
def search_word_enter(self):
   search_word()
def add_words_in():
    global path
    file=open(path,'a+',encoding="gbk")
    if word=="None" or word=="" or meaning=="" or meaning=="None":
        sg.popup_error("单词输入错误,请检查！",auto_close_duration=2,keep_on_top=True,no_titlebar=True,auto_close=True)
    else:
        word_meaning = str(word)+","+str(meaning)
        meaning_word = str(meaning)+","+str(word)
        file.write("\n"+word_meaning)
        file.write("\n"+meaning_word)
        file.close()
def add_words():
    global word
    global meaning
    word = sg.popup_get_text('请输入单词')
    meaning = sg.popup_get_text('请输入释义')
    add_words_in()
if __name__=='__main__':
    body = Tk()
    body.title("英汉汉英词典录入软件v2.0")
    body.resizable(0,0)
    body.geometry("400x230")
    body.iconbitmap("1.ico")

    word=StringVar()
    meaning=StringVar()
    v = IntVar()

    menubar=Menu(body)
    filemenu=Menu(menubar,tearoff=False)
    filemenu.add_command(label="重置计数器", command=delete_num)
    filemenu.add_command(label="合并单词表格", command=hebing)
    filemenu.add_separator()
    filemenu.add_command(label="选择词典路径...",command=select_dictionary)
    filemenu.add_command(label="选择表格路径...", command=select_excel)
    filemenu.add_command(label="帮助我们添加单词",command=add_words)
    filemenu.add_separator()
    filemenu.add_command(label="退出",command=body.quit)
    menubar.add_cascade(label="工具",menu=filemenu)

    editmenu=Menu(menubar,tearoff=False)
    editmenu.add_command(label="关于",command=about)
    editmenu.add_command(label="访问我们的GitHub", command=GitHub)
    editmenu.add_command(label="联系我们", command=GitHub)
    menubar.add_cascade(label="关于...",menu=editmenu)

    editmenu=Menu(menubar,tearoff=False)
    editmenu.add_command(label="中译英",command=ch_to_en)
    editmenu.add_command(label="英译中",command=en_to_ch)
    menubar.add_cascade(label="网络翻译设置",menu=editmenu)

    body.config(menu=menubar)
    
    frame_in = Frame(body, width=300, height=30)
    frame_in.place(x=50, y=10)
    entry = Entry(frame_in, width=30)
    entry.pack(side="left")
    
    
    btn = Button(frame_in, text="查找", width=10, command=search_word)
    btn.pack(side="right", padx=10)
    
    entry.bind("<Return>", search_word_enter)

    frame_doc = Frame(body, width=350, height=200)
    frame_doc.place(x=20, y=40)
    bar = Scrollbar(frame_doc)
    bar.pack(side="right", fill=Y)
    doc = Text(frame_doc, width=50, height=9.2)
    doc.pack(side="bottom", pady=15)
    doc.config(yscrollcommand=bar.set)
    bar.config(command=doc.yview)
    doc.insert('end', "请输入查询词条,按回车或点击查询...\n")
    #按钮部分
    Button(body,text="退出",width=10,command=body.quit).place(x=300,y=190)
    c = Checkbutton(body,text=text1,variable=v).place(x=30,y=190)

    body.update_idletasks()
    x = (body.winfo_screenwidth() - body.winfo_reqwidth()) / 2
    y = (body.winfo_screenheight() - body.winfo_reqheight()) / 2
    body.geometry("+%d+%d" % (x, y))
    
    body.mainloop()