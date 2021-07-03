import csv
import random
import xlrd
import openpyxl
from xlutils.copy import copy
import requests
import os
import urllib.request
import urllib.parse
import xlwt
import sys

#调用函数
def search_data(a):#调用有道查找释意
    data = {
        'doctype': 'json',
        'type': 'AUTO',
        'i':a
    }
    url = 'http://fanyi.youdao.com/translate?smartresult=dict&smartresult=rule&sessionFrom=null'
    r = requests.get(url,params=data)
    result = r.json()
    result1=result['translateResult']
    result2=result1[0]
    result3=result2[0]
    result4=result3['tgt']
    print(result4)
    return result4

def write_excel_xlsx(path,value3):#数据写入excel
    index = len(value3)
    workbook = xlrd.open_workbook(path)
    sheets = workbook.sheet_names()
    worksheet = workbook.sheet_by_name(sheets[0])
    num=worksheet.cell_value(0, 2)
    num=num+index-1
    rows_old = worksheet.nrows-1
    new_workbook = copy(workbook)
    new_worksheet = new_workbook.get_sheet(0)
    new_worksheet.write(0,2,num)
    for i in range(0, index):
        new_worksheet.write(i+rows_old,0, value3[i][0])
        new_worksheet.write(i+rows_old,1, value3[i][1])
        new_worksheet.write(i+rows_old,2, 1)
        new_worksheet.write(i+rows_old,3, 0)
        new_worksheet.write(i+rows_old,4, 0)
        new_worksheet.write(i+rows_old,5, 0)
        new_worksheet.write(i+rows_old,6, 0)
    new_workbook.save(path)
def insert_data():#插入功能
    i=0
    value3=[]
    insert1=[]
    book_name_xlsx = 'data.xls'
    char = input()
    insert1.append(char)
    insert1.append(search_data(char))
    value3.append(insert1)
    i=i+1
    while char!=' ':
        char = input()
        insert1=[]
        insert1.append(char)
        insert1.append(search_data(char))
        value3.append(insert1)
        i=i+1
    write_excel_xlsx(book_name_xlsx,value3)
def right_answer(num,path):
    workbook = xlrd.open_workbook(path)
    sheets = workbook.sheet_names()  
    worksheet = workbook.sheet_by_name(sheets[0])
    chan1=worksheet.cell_value(0,2)
    chan2=worksheet.cell_value(num,2)
    chan3=worksheet.cell_value(num,3)
    chan4=worksheet.cell_value(num,4)
    chan5=worksheet.cell_value(num,5)
    rows_old = worksheet.nrows-1
    new_workbook = copy(workbook)
    new_worksheet = new_workbook.get_sheet(0)
    if chan3==0:
        chan2=chan2-0.1
        chan1=chan1-0.1
        new_worksheet.write(0,2,chan1)
        new_worksheet.write(num,2,chan2)
        new_worksheet.write(num,3,1)
        new_worksheet.write(num,4,0)
        new_worksheet.write(num,5,1)
    elif chan3==1 and chan4 == 0 and chan5 != 4:
        chan2=chan2-0.1
        chan1=chan1-0.1
        chan5=chan5+1
        new_worksheet.write(0,2,chan1)
        new_worksheet.write(num,2,chan2)
        new_worksheet.write(num,5,chan5)
    elif chan3 == 1 and chan4 == 1 and chan5!=5:
        chan0=(chan2-1)/5
        chan1=chan1-chan0
        chan2=chan2-chan0
        chan5=chan5+1
        new_worksheet.write(0,2,chan1)
        new_worksheet.write(num,2,chan2)
        new_worksheet.write(num,5,chan5)
    elif chan3 == 1 and chan4 ==1 and chan5 == 5:
        chan0=(chan2-1)/5
        chan1=chan1-chan0
        chan2=chan2-(chan2-1)/5
        chan4=0
        chan5=0
        new_worksheet.write(0,2,chan1)
        new_worksheet.write(num,2,chan2)
        new_worksheet.write(num,4,chan4)
        new_worksheet.write(num,5,chan5)
        new_worksheet.write(num,6,0)
    new_workbook.save(path)

def wrong_answer(num,path):
    workbook = xlrd.open_workbook(path)
    sheets = workbook.sheet_names()  
    worksheet = workbook.sheet_by_name(sheets[0])
    chan1=worksheet.cell_value(0,2)
    chan2=worksheet.cell_value(num,2)
    chan3=worksheet.cell_value(num,3)
    chan4=worksheet.cell_value(num,4)
    chan5=worksheet.cell_value(num,6)
    rows_old = worksheet.nrows-1
    new_workbook = copy(workbook)
    new_worksheet = new_workbook.get_sheet(0)
    if chan3 ==0:
        chan2=chan2+1/3
        chan1=chan1+1/3
        chan3=1
        chan4=1
        chan5=1
        new_worksheet.write(0,2,chan1)
        new_worksheet.write(num,2,chan2)
        new_worksheet.write(num,3,chan3)
        new_worksheet.write(num,4,chan4)
        new_worksheet.write(num,6,chan5)
    elif chan3==1 and chan4==0:
        chan1=chan1+1-chan2
        chan2=1
        chan4=1
        chan5=1
        new_worksheet.write(0,2,chan1)
        new_worksheet.write(num,2,chan2)
        new_worksheet.write(num,4,chan4)
        new_worksheet.write(num,5,0)
        new_worksheet.write(num,6,chan5)
        
    elif chan3==1 and chan4==1:
        chan1=chan1+(1/3)**chan5
        chan2=chan2+(1/3)**chan5
        chan5=chan5+1
        new_worksheet.write(0,2,chan1)
        new_worksheet.write(num,2,chan2)
        new_worksheet.write(num,6,chan5)
    new_workbook.save(path)
        
    
def review_data(path):#复习功能
    workbook = xlrd.open_workbook(path)
    sheets = workbook.sheet_names()  
    worksheet = workbook.sheet_by_name(sheets[0])
    rows_old = worksheet.nrows-1
    print('请输入是否开始进行复习（Y/N):')
    num=worksheet.cell_value(0,2)
    a=input()
    c=[]
    while a!='N'and a!='n':
        m=random.randint(1,rows_old-1)
        print((worksheet.cell_value(m,1)))
        a=input()
        if a == worksheet.cell_value(m, 0):
            right_answer(m,'data.xls')
            print("答案正确")
        else :
            print("答案错误")
            wrong_answer(m,'data.xls')
            print(worksheet.cell_value(m, 0))
            c.append(m)
    print("是否进行复习此次错误词汇？(Y/N)")
    choicen=input()
    if choicen == "Y" or choicen=='y':
        for k in c:
            print((worksheet.cell_value(k,1)))
            a=input()
            if a == worksheet.cell_value(k, 0):
                print("答案正确")
            else :
                print("答案错误")
    else:print("结束")
    
def maini():
    while input!='N'or'n':
        print("请选择功能")
        print("A.查询\tB.保存单词\tC.复习\n")
        choice2=input()
        if choice2=='A'or choice2=='a' :
            print('请输入要查询词汇')
            data= input()
            search_data(data)
        elif choice2 == 'B' or choice2== 'b':
            print("单独输入空格结束程序")
            insert_data()
        elif choice2 == 'C' or choice2== 'c':
            print("再次提醒未输入单词不能使用此功能")
            print("输入N结束复习")
            review_data('data.xls')
        print("是否继续？Y/N")
        ax=input()
        if ax=='Y'or  ax=='y':
            return maini();
        else:
            sys.exit()
#主函数
print("//注意：首次使用请不要直接选择B.C功能且谨慎选择第一次使用选项")
dir_path =os.path.dirname(os.path.realpath(__file__))
print("请输入是否第一次使用:Y/N")
choice=input()
if choice == 'Y'or choice=='y':
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet1 = workbook.add_sheet("测试表格")
    sheet1.write(0,0,'英文')
    sheet1.write(0,1,'中文')
    sheet1.write(0,2,0)
    sheet1.write(1,0,' ')
    workbook.save('data.xls');
    maini()
else:
    maini()
 
