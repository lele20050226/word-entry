#! python3

'''
时间：3.15-4.14
Version:2.0
单词录入系统
'''

import csv
import shutil
import win32com.client as wincl
from translate import Translator
import PySimpleGUI as sg
from tkinter import *
import tkinter.messagebox
import os
import xlrd
import xlwt
from xlutils.copy import copy
import webbrowser
import sys
import win32com.client as win32
import ctypes
from tkinter import *
from tkinter.ttk import *
import win32con
import win32api

try:
    deletenum=False
    numfile=open("numfile.log",'r')
    num=int(numfile.readline())
    numfile.close()
except FileNotFoundError:
    numfile = open("numfile.log", 'a+')
    numfile.write("0")
    numfile.close()
    win32api.SetFileAttributes('numfile.log', win32con.FILE_ATTRIBUTE_HIDDEN)
text1="自动导入excel"

change=False
failcount=0

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
    tkinter.messagebox.showinfo("关于","岳阳市岳阳楼区第十五中学 413班梁家诺制作\nVersion:2.0")
def GitHub():#链接到GitHub
    webbrowser.open('https://github.com/lele20050226/word-entry')
def help():
    webbrowser.open('https://github.com/lele20050226/word-entry/issues')
#加载本地词条
def loading():
    try:
        try:
            local_dict = {}
            with open(path,'r',encoding="gbk") as file:
                for line in file:
                    line = line.rstrip().split(',')
                    local_dict[line[0]] = line[1]
            return local_dict
        except FileNotFoundError:
            pass
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
    try:
        ans = translator.translate(word)
        string = '在线查询结果：\n'+word+':'+ans+'\n'
    except UnboundLocalError:
        pass

    if int(v.get()) == 1:
        try:
            rbook = xlrd.open_workbook(excel_path, formatting_info=True)  # 打开文件
        except NameError:
            select_excel()
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
    if path=='':
        tkinter.messagebox.showwarning("提醒","您将不使用本地词典，只选用在线词典。如果要再次选择，请单击工具-选择词典路径")
        f=open("temporary_dic.csv","a")
        f.write("abandonment,n.放弃")
        win32api.SetFileAttributes('temporary_dic.csv', win32con.FILE_ATTRIBUTE_HIDDEN)
        f.close()
def select_excel():
    global excel_path
    excel_path = sg.popup_get_file("请选择表格文件：",file_types=(("表格文件",".xls"),("表格文件",".xlsx")))
    if os.path.splitext(excel_path)[1] == ".xls":
        pass
    elif os.path.splitext(excel_path)[1] == ".xlsx":
        path2 = excel_path.strip('x')
        workbook = xlwt.Workbook(encoding='utf-8')
        worksheet = workbook.add_sheet('words')
        worksheet.write(0, 0, label='序号')
        worksheet.write(0, 1, label='单词')
        worksheet.write(0, 2, label='释义')
        workbook.save(path2)
        excel_path=path2
        tkinter.messagebox.showinfo("提示", "储存的单词将会保存到新的表格中")

#查找单词
def search(word,local_dict):
    count=0
    global key
    global intocsv
    global num
    global deletenum
    global new_mean
    global change
    global path2
    result='本地部分匹配结果：\n'
    try:
        for key in local_dict.keys():
            if word.upper() == key.upper():  # 大小写不敏感
                if int(v.get()) == 1:
                    try:
                        rbook = xlrd.open_workbook(excel_path, formatting_info=True)  # 打开文件
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
                                wbook.save(excel_path)  # 保存文件
                        except PermissionError:
                            tkinter.messagebox.showerror("错误", "请关闭Excel软件后重试！")
                    except NameError:
                        select_excel()
                else:
                    pass
                numfile1 = open("numfile.log", 'w')
                numfile1.write(str(num))
                numfile1.close()
                numfile2 = open("numfile.log", 'r')
                num = int(numfile2.readline())
                numfile2.close()
                return '本地查询结果：\n' + key + ': ' + local_dict.get(key) + '\n'
            elif word.upper() in key.upper():
                result = result + key + ': ' + local_dict.get(key) + '\n'
                count = count + 1
    except AttributeError:
        global failcount
        if failcount==0:
            failcount = failcount + 1
        else:
            count=0
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
        result = search(word, local_dict)
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

def read_word():
    try:
        speak = wincl.Dispatch("SAPI.SpVoice")
        speak.Speak(entry.get())
    except NameError:
        tkinter.messagebox.showinfo("提醒","您还没有选定或输入单词！")

if __name__=='__main__':
    en_to_ch()

    body = Tk()
    body.title("英汉汉英词典录入软件v2.0")
    body.resizable(0,0)
    body.geometry("400x230")
    body.iconbitmap("icon.ico")

    word=StringVar()
    meaning=StringVar()
    v = IntVar()

    menubar=Menu(body)
    filemenu=Menu(menubar,tearoff=False)
    filemenu.add_command(label="重置计数器", command=delete_num)
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
    editmenu.add_command(label="联系我们", command=help)
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
    doc = Text(frame_doc,bg="white", width=50, height=9.2)
    doc.pack(side="bottom", pady=15)
    doc.config(yscrollcommand=bar.set)
    bar.config(command=doc.yview)
    doc.insert('end', "请输入查询词条,按回车或点击查询...\n")
    #按钮部分
    Button(body,text="朗读",width=10,command=read_word).place(x=300,y=190)
    c = Checkbutton(body,text=text1,variable=v).place(x=30,y=190)

    body.update_idletasks()
    x = (body.winfo_screenwidth() - body.winfo_reqwidth()) / 2
    y = (body.winfo_screenheight() - body.winfo_reqheight()) / 2
    body.geometry("+%d+%d" % (x, y))
    
    body.mainloop()
