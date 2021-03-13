'''
时间：9.30-10.17
Version:1.0
单词录入系统
'''
import csv
from tkinter import *
import tkinter.messagebox
import os
import xlrd
import xlwt
from bs4 import BeautifulSoup
from xlutils.copy import copy
'''
import requests,re
'''
deletenum=False
numfile=open("numfile.txt",'r')
num=int(numfile.readline())
numfile.close()
intocsv=False

text1="自动导入excel"


def loading():
    # 加载本地词条
    local_dict = {}
    with open("英汉汉英词典.csv",'r',encoding="gbk") as file:
        for line in file:
            line = line.rstrip().split(',')
            local_dict[line[0]] = line[1]
    return local_dict
#2.0版本开发net search、analysis和download在线输入
def net_search(word):
	# 本地词库搜索失败后的备用在线搜索：有道词典
    try:
        '''
        if __name__=='__main__':
            while(1):
                analysis()
        string='本地词库尚未收录或拼写不全，下滑查看有道在线直接查询和本地部分匹配结果：\n'+word+': '
        for item in s:
            i=i.strip()
            if item.text:
                string+=str(item.text)+' '
        '''
        string+='\n'
    except Exception:
        string='请检查词的拼写是否正确\n'
    finally:
        return string
'''
在线查询（暂未启用）
def analysis():
    list1=re.findall("详细释义.+<p class=\"collapse-content\">",download(),re.S)   #这里对html字符串进行第一步加工，截取大概的信息
    list2=re.findall("                [a-zA-Z ]+",str(list1))      #将上面加工后的字符串进一步加工，直接提取到所有翻译后的单词信息
    for i in list2:
        i=i.strip()   #因为第二步加工后的信息并不干净，得到的单词前面会有空格，这里将空格删去
        return '在线查询结果：\n'+key+': '+i+'\n'
def download():
    word=key
    url="http://dict.youdao.com/w/eng/"+word+"/#keyfrom=dict2.index"   #合并URL地址
    html=requests.get(url).content.decode('utf-8')    #得到服务器的相应信息后将其转码为UTF-8
    return html
'''
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
                excel_path='words.xls'#文件路径
                rbook = xlrd.open_workbook(excel_path,formatting_info=True)#打开文件
                w_sheet = rbook.sheet_by_index(0)
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
            else:
                pass
            numfile1=open("numfile.txt",'w')
            numfile1.write(str(num))
            numfile1.close()
            numfile2=open("numfile.txt",'r')
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
    numfile=open("numfile.txt",'w')
    numfile.write('0')
    numfile.close()
    if os.path.exists("words.xls"):
        os.remove("words.xls")
    else:
        pass
    num=0
    xh=1
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('words')
    worksheet.write(num,0, label = '序号')
    worksheet.write(num,1, label = '单词')
    worksheet.write(num,2, label = '释义')
    workbook.save('words.xls')

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

def about():
    tkinter.messagebox.showinfo("关于","岳阳市第十五中学423班 梁家诺制作\nVersion:1.0\n支持从英汉/汉英翻译词典中调入单词数据并自动输入到excel中")

if __name__=='__main__':
    body = Tk()
    body.title("本地英汉汉英词典录入软件")
    body.resizable(0,0)
    body.geometry("400x270")
    body.iconbitmap("1.ico")

    v = IntVar()

    menubar=Menu(body)
    filemenu=Menu(menubar,tearoff=False)
    filemenu.add_command(label="删除计数器（将会删除excel内所有的数据！）",command=delete_num)
    filemenu.add_separator()
    filemenu.add_command(label="退出",command=body.quit)
    menubar.add_cascade(label="工具",menu=filemenu)

    editmenu=Menu(menubar,tearoff=False)
    editmenu.add_command(label="关于",command=about)
    editmenu.add_separator()
    editmenu.add_command(label="退出",command=body.quit)
    menubar.add_cascade(label="关于...",menu=editmenu)
    body.config(menu=menubar)

    
    frame_in = Frame(body, width=300, height=30)
    frame_in.place(x=50, y=10)
    entry = Entry(frame_in, width=30)
    entry.pack(side="left")
    
    
    btn = Button(frame_in, text="Search", width=10, command=search_word)
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
    
    Button(body,text="退出",width=10,command=body.quit).place(x=300,y=190)
    c = Checkbutton(body,text=text1,variable=v).place(x=30,y=190)
    Label(body,text="就绪",justify=LEFT,padx=10).place(x=0,y=220)
    
    body.mainloop()
