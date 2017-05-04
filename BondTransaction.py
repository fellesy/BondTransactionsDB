# -*- coding: utf-8 -*-

import xlwt
import xlrd
import wx
import wx.grid as gridlib
import time
import sqlite3
import re
import sys
import os
import wx.lib.agw.customtreectrl as CT
from datetime import datetime
import MySQLdb
from threading import Thread
from Queue import Queue

reload(sys)
sys.setdefaultencoding('utf-8')
excel_title = [u"成交时间",u"期限",u"债券代码",u"债券简称",u"利率",u"信用评级",u"类型",u"中介机构", u"数据库编号"]
database_title = [u"成交时间",u"期限",u"债券代码",u"债券简称",u"利率",u"信用评级",u"类型",u"中介机构", u"筛选条件-天数", u"筛选条件-价格",u"筛选条件-评级1",u"筛选条件-评级2"]
select_columns = "date, term_text, bond_id, name, price_text, rating_text, type, agency, id"

def import_text(txtpath,xlpath,date):
    print "-------import from txt----------"
    punc = [" ", "\t"]
    agency = ""
    bond_type = ""
    agency_list = [u"平安信用", u"平安利率", u"BGC信用", u"国际信用", u"国际利率", u"国利信用", u"国利利率", u"信唐"]
    bond_list = [u"短融", u"企业债", u"公司债", u"其他", u"存单", u"国债", u"金融债", u"中票", u"金融债（固息）"]
    row_list = []
    export_data = []

    f = open(txtpath, 'r+')
    for row in f:
        temp =''.join([cell if cell not in punc else ' ' for cell in row]).split()
        row_list.append(temp)

    get_agency = lambda x: x if x in agency_list else agency
    get_bond_type  = lambda x: x if x in bond_list else bond_type

    for row in row_list:

        if len(row)==1:
            value = row[0].decode('gb2312').strip(u"： 成交")
            agency = get_agency(value)
            bond_type = get_bond_type(value)
        elif (len(row)>1) and (bond_type not in [u"国债", u"金融债", u"金融债（固息）"]):
            if agency ==u"平安信用":
                temp = row[2]
                row[2]= row[1]
                row[1] = temp
            if agency in [u"平安信用", u"国际信用",u"信唐"]:
                temp = row[3]
                row[3]= row[4]
                row[4]= temp
            temp_row = []
            temp_row.append(date.replace("-",""))
            for item in row:
                value = item.strip()
                temp_row.append(value.decode('gb2312'))
            temp_row.append(bond_type)
            temp_row.append(agency)
            export_data.append(temp_row)
    success_rows, fail_rows = test_insert(export_data)
    success_rows.insert(0,database_title)
    export_excel(data =success_rows, wrong_data=fail_rows, xlpath=xlpath)


def import_excel(xlpath, conn):
    print "-------import from excel----------"
    book = xlrd.open_workbook(xlpath)
    sheet = book.sheet_by_index(0)
    nrow = sheet.nrows
    data = []
    fail_rows1 = []

    for i in range(1,nrow):
        temp_row = sheet.row_values(i)
        row = []
        #判断格式是否正确，如果满足条件，对数据进行调整然后导入云端数据库

        #不能小于8格
        if (len(temp_row) <8 or ("" in temp_row[:8]) or (" " in temp_row[:8]) ):
            fail_rows1.append(temp_row)

        else:
            #去除行尾的空白格
            for j in range(len(temp_row)):
                if j < 8:
                        row.append(temp_row[j])
                elif (temp_row[j] not in [""," "]):
                    row.append(temp_row[j])

            #不可以大于12格
            if (len(row)>12):
                fail_rows1.append(row)
            else:
                #调整导入数据格式
                row = adjust_row(row[0:8])
                if row[-1]=="error":
                    fail_rows1.append(tuple(row))
                else:
                    data.append(tuple(row))
    # success_rows, fail_rows2 = insert_table(data=row_list,conn=conn)
    # return success_rows,fail_rows1+fail_rows2
    return data,fail_rows1


def export_excel(data,xlpath, wrong_data =None):
    print "-------export to excel----------"
    if len(data) ==0 :
        return False

    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('Sheet1')
    badBG = xlwt.Pattern()
    badBG.pattern = badBG.SOLID_PATTERN
    badBG.pattern_fore_colour = 3

    badFontStyle = xlwt.XFStyle()
    badFontStyle.pattern = badBG


    for i in range(0,len(data)):
            for j in range(len(data[i])):

                worksheet.write(i, j, str(data[i][j]).decode('utf-8'))
    if wrong_data !=None:
        for i in range(0,len(wrong_data)):
                for j in range(len(wrong_data[i])):
                    worksheet.write(i+len(data), j, str(wrong_data[i][j]).decode('utf-8'), badFontStyle)
    print "-----------xlpath-------------"
    print xlpath
    workbook.save(xlpath)

def create_table(conn, name ):
    cursor = conn.cursor()
    print "table name " + name
    cursor.execute('''CREATE TABLE %s (
                        id INTEGER PRIMARY KEY ,
                        date long,
                        term_text text,
                        bond_id text,
                        name text,
                        price_text char(50),
                        rating_text char(50),
                        type char(50),
                        agency char(50),
                        term int,
                        price real,
                        company_rating char(50),
                        bond_rating char(50));'''%(name))
    conn.commit()
    print "-------create table successfully--------"
    cursor.close()
    return True

def create_local_table(dbpath):
    #在本地创建sqlite3的数据库，主要用以测试数据是否能够导入数据库
    conn = sqlite3.connect(dbpath)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE TR (
                        id INTEGER PRIMARY KEY ,
                        date long,
                        term_text text,
                        bond_id text,
                        name text,
                        price_text char(50),
                        rating_text char(50),
                        type char(50),
                        agency char(50),
                        term int,
                        price real,
                        company_rating char(50),
                        bond_rating char(50));''')
    conn.commit()
    print "-------create local table successfully--------"
    cursor.close()
    return True

def insert_table(conn,data):
    print "------insert_table------"

    fail_collection = []
    success_collection = []

    temp =""
    for item in data:
        date = item[0][0:6]
        if temp!=date:
            temp = date
            if not IsTableExist(conn,"tr%s"%temp):
                if not create_table(conn,"tr%s"%temp):
                    continue

    cursor = conn.cursor()
    for item in data:
        date = item[0][0:6]
        table = "tr" +date
        cursor.execute("SELECT MAX(id) FROM %s "%table)


        max_id = cursor.fetchone()[0]
        if max_id ==None:
            i = 1
        else:
            i = max_id + 1

        temp = (i,) + tuple(item)
        try:
            cursor.execute("INSERT INTO "+table+" VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);", temp)
            i += 1
            success_collection.append(item)
        except Exception as err:
            print("Something went wrong: {}".format(err))
            print "fail to insert"
            fail_collection.append(item)

    conn.commit()
    cursor.close()
    return success_collection, fail_collection


def insert_local_table(dbpath,data):
    print "------insert_local_table------"
    conn = sqlite3.connect(dbpath)
    cursor = conn.cursor()

    fail_collection = []
    success_collection = []
    i=1

    for item in data:
        temp = (i,) + tuple(item)
        try:
            cursor.execute("INSERT INTO TR VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?);", temp)
            i += 1
            success_collection.append(item)
        except:
            print "fail to insert"
            fail_collection.append(item)

    conn.commit()
    cursor.close()
    return success_collection, fail_collection

def IsTableExist(conn,table):
    cursor = conn.cursor()
    tables = []
    try:
        cursor.execute("SHOW TABLES")
        result = cursor.fetchall()
        tables.extend(x[0] for x in result)
    except Exception as err:
        print("Something went wrong: {}".format(err))

    cursor.close()

    if not tables:
        print table + " not exist "
        return False

    if table in tables:
        return True
    else:
        print table + " not exist "
        return False

def drop_table(conn,table):
    cursor = conn.cursor()
    cursor.execute("DROP TABLE " + table )
    conn.commit()
    cursor.close()

def select_table(conn,filter_clause):
    cursor = conn.cursor()
    try:
        cursor.execute(filter_clause)
        result = cursor.fetchall()
        cursor.close()
        return result
    except:
        cursor.close()
        return []

def search_table(conn,item, keyword, table, filter=""):
    cursor = conn.cursor()
    if keyword:
        if filter:
            clause = "SELECT "+select_columns+" FROM " \
                + table+ " WHERE (" + item + " like '%"+keyword+"%') AND ("+filter +")"
        else:
            clause = "SELECT "+select_columns+" FROM " \
                + table+ " WHERE (" + item + " like '%"+keyword+"%')"
        print clause

        cursor.execute(clause)
        result = cursor.fetchall()
        print "first shot of search "+ str(result)

        if (len(result) == 0) and (item =="name"):
            regex = "REGEXP '%s'"%('.*'.join(keyword)+".*")

            if filter:
                clause = "SELECT %s FROM %s WHERE (name %s) AND (%s) order by date desc"%(select_columns,table,regex,filter)
            else:
                clause = "SELECT %s FROM %s WHERE name %s order by date desc"%(select_columns,table,regex)
            print clause
            cursor.execute(clause)
            result = cursor.fetchall()
            print "second shot of search: "+ str(result)
        cursor.close()
    else:
        if filter:
            clause = "SELECT %s from %s Where %s order by date desc"%(select_columns,table, filter)
        else:
            clause = "SELECT %s from %s order by date desc"%(select_columns,table)
        print clause

        cursor.execute(clause)
        result = cursor.fetchall()
    return result


def get_tables(conn):
    cursor = conn.cursor()
    tables = []
    cursor.execute("SHOW TABLES")
    result = cursor.fetchall()
    tables.extend(x[0] for x in result)
    cursor.close()
    tables.sort(reverse=True)
    return tables

def test_insert(data):
    test_db = 'test.db'
    try:
        create_local_table(test_db)
    except Exception as err:
        print err
    row_list = []
    for row in data[1:]:
        temp = adjust_row(row)
        row_list.append(tuple(temp))
    test_result = insert_local_table(dbpath=test_db, data =row_list)
    os.remove(test_db)
    return test_result

def del_row_table(conn, id,table):
    print "----------delete data in table----------"
    cursor = conn.cursor()
    print table
    cursor.execute("DELETE FROM %s WHERE ID = %s"%(table, str(id)))
    conn.commit()
    cursor.close()
    return True


# def fuzzyFinder(keyword,collection):
#     suggestions =[]
#     pattern = '.*'.join(keyword)
#     regex = re.compile(pattern)
#     for item in collection:
#         match = regex.search(item)
#         if match:
#             suggestions.append(item)
#     if suggestions:
#         return suggestions
#     else:
#         print " no perfect match. Find similar results"
#         pattern = '[%s]*'%keyword
#         regex = re.compile(pattern)
#         for item in collection:
#             match = regex.search(item)
#             if match:
#                 suggestions.append(item)
#         return suggestions



def get_time(type = 0):
    if type==0:
        return time.strftime('%Y%m%d',time.localtime(time.time()))
    elif type ==1:
        return time.strftime('%Y-%m-%d',time.localtime(time.time()))
    elif type ==2:
        return time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))

def adjust_row(data):
    adjusted_data = []

    date = str(data[0]).replace("-","")
    adjusted_data.append(date)
    for item in data[1:]:
        adjusted_data.append(item)


    re_term = u"[0-9.]+[DMYdmy]{0,1}"
    pattern = re.compile(re_term)
    term_result = re.match(pattern,str(data[1]))
    term = ""
    term2 = ""
    if term_result != None:
        term = term_result.group(0)
        re_term2 = u"[+][0-9.]+[DMYdmy]{1}"
        pattern2 = re.compile(re_term2)
        term_result2 = re.search(pattern2,str(data[1]))

        if term_result2!=None:
            term2 = term_result2.group(0)
        if term[-1] not in "DMYdmy":
            term += "Y"
        adjusted_data.append(StrToDays(term)+StrToDays(term2))
    else:
        if data[1] ==u"永续":
            adjusted_data.append("100000000")
            term_result = 'perpetuity'

    re_price = u"[0-9]{1,2}[.]{0,1}[0-9]{0,4}"
    pattern = re.compile(re_price)
    price_result = re.match(pattern,str(data[4]))
    if price_result!= None:
        price = float(price_result.group(0))
        adjusted_data.append(price)
    else:
        re_price = u"[1-9]{1,4}[.]{0,1}[1-9]{0,4}"
        pattern = re.compile(re_price)
        price_result = re.search(pattern, str(data[4]))
        if price_result != None:
            price = float(price_result.group(0))
            adjusted_data.append(price)

    rating1 = "0"
    rating2 = "0"
    re_rating = u"[ABC+]+[/]{0,1}[ABC01-]{0,4}"
    pattern = re.compile(re_rating)
    rating_result = re.match(pattern,str(data[5]))
    if rating_result!=None:
        re_rating1 = u"[ABC]{1,3}[+-]{0,1}"
        re_rating2 = u"[/][ABC]{1,3}[+-]{0,1}[1,3]{0,1}"

        result1 = re.match(re.compile(re_rating1),str(data[5]))
        result2 = re.search(re.compile(re_rating2),str(data[5]))
        if result1 !=None:
            rating1 = result1.group(0)
        else:
            rating1 = "0"
        if result2 != None:
            rating2 = result2.group(0).replace("/","")

        adjusted_data.append(rating1)
        adjusted_data.append(rating2)

    elif data[5] in (' ', '', '0.0','0',0):
        adjusted_data.append("0")
        adjusted_data.append("0")
        rating_result = '0'
    else:
        try:
           rating1 = int(data[5])
           if rating1 ==0:
                adjusted_data.append("0")
                adjusted_data.append("0")
        except:
            pass

    if not (price_result) or (not term_result) or (len(adjusted_data)>12) or (not rating_result):
        adjusted_data.append("error")

    return adjusted_data


def StrToDays(term):
    term_in_days = 0
    if ("Y" in term) or ("y" in term):
        term_in_days = int(float(term[:-1]) * 365)
    elif ("D" in term) or ("d" in term):
        term_in_days = int(float(term[:-1]))
    elif ("M" in term) or ("m" in term):
        term_in_days = int(float(term[:-1]) * 30)
    return term_in_days

def rating_index(rating):
    rating_list = ["AAA+","AAA","AAA-","AA+","AA","AA-","BBB+","BBB","BBB-","BB+", "BB","BB-","B+","B","B-"]
    cp_rating_list = ["A-1", "A-2", "A-3"]

    if rating in rating_list:
        return float(rating_list.index(rating)+1)
    elif rating in cp_rating_list:
        return float(cp_rating_list.index(rating) +101)
    else:
        return 0.0

def IsDate(*date):
    re_date = u"^(?:(?!0000)[0-9]{4}-(?:(?:0[1-9]|1[0-2])-(?:0[1-9]|1[0-9]|2[0-8])|(?:0[13-9]|1[0-2])-(?:29|30)|(?:0[13578]|1[02])-31)|(?:[0-9]{2}(?:0[48]|[2468][048]|[13579][26])|(?:0[48]|[2468][048]|[13579][26])00)-02-29)$"
    p = re.compile(re_date)
    results =[]
    for item in date:
        result = re.match(p, item)
        if result!= None:
            results.append(result.group(0))
    if len(results) == len(date):
        return True
    else:
        dlg = wx.MessageDialog(None, u"日期格式错误", u"错误提示", wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_YES:
            dlg.Destroy()
        return False

def IsNumber(*number):
    try:
        for item in number:
            float(item)
        return True
    except:
        return False

def monthdelta(date1,date2):
    years = date2.year - date1.year
    months = []
    month_format = lambda x:"0"+str(x) if x<10 else str(x)

    if years ==0:
        months.extend(str(date1.year) + month_format(x) for x in range(date1.month, date2.month+1))
        return months
    else:
        months.extend(str(date1.year)+month_format(x) for x in range(date1.month,13))
        for year in range(date1.year+1,date2.year):
            months.extend(str(year)+month_format(x) for x in range(1,13))

        months.extend(str(date2.year)+month_format(x) for x in range(1,date2.month+1))
        return months


class MainWindow(wx.Frame):
    def __init__(self, parent, title):
        self.connection = self.Connect_MySQL()
        if not self.connection:
            sys.exit(0)

        wx.Frame.__init__(self, parent, title = title,size = (700,300))

        ANCHOR = 20
        SPACE = 10
        WIDTH = 80
        HEIGHT = 25

        self.bond_types = [u"短融",u"企业债", u"公司债",u"存单",u"中票",u"金融债",u"其他"]
        self.company_ratings = ["AAA+", "AAA", "AAA-", "AA+", "AA", "AA-","A+", "A", "A-","BBB+", "BBB", "BBB-", "BB+", "BB", "BB-", "B+", "B","B-","0"]
        self.bond_ratings = ["AAA+", "AAA", "AAA-", "AA+", "AA", "AA-","A-1","A+", "A", "A-","BBB+", "BBB", "BBB-", "BB+", "BB", "BB-", "B+", "B","B-","A-2","0"]
        self.agencies =[u"平安",u"BGC",u"国际",u"国利",u"信唐",u"空缺"]
        self.term_units=[[u"年",u"月",u"日"],[365,30,1]]
        self.search_column = [[u"简称", u"代码"],["name","bond_id"]]
        bkg = wx.Panel(self,style=wx.TAB_TRAVERSAL | wx.CLIP_CHILDREN | wx.FULL_REPAINT_ON_RESIZE)

        getdata_button = wx.Button(bkg, label=u"提取数据", size = (WIDTH,HEIGHT), pos = (ANCHOR+(WIDTH+SPACE)*6.2,ANCHOR))
        getdata_button.Bind(wx.EVT_BUTTON, self.OnGetData)

        txt_button = wx.Button(bkg, label = u"txt转excel", size = (WIDTH,HEIGHT), pos=(ANCHOR+(WIDTH+SPACE)*6.2,ANCHOR+(HEIGHT+SPACE)*1))
        txt_button.Bind(wx.EVT_BUTTON, self.OnImportTxt)

        xl_button  = wx.Button(bkg, label = u"导入excel",    size = (WIDTH,HEIGHT), pos=(ANCHOR+(WIDTH+SPACE)*6.2,ANCHOR+(HEIGHT+SPACE)*2))
        xl_button.Bind(wx.EVT_BUTTON, self.OnImportExcel)

        self.search_text = wx.TextCtrl(bkg,size = (WIDTH,HEIGHT), style= wx.TE_PROCESS_ENTER, pos = (ANCHOR+WIDTH+SPACE,ANCHOR+SPACE))
        self.search_text.Bind(wx.EVT_TEXT_ENTER, self.OnGetData)
        self.search_column_cb = wx.ComboBox(bkg,choices=self.search_column[0], size=(WIDTH/1.5,HEIGHT),pos =(ANCHOR+(WIDTH+SPACE/2)*2,ANCHOR+SPACE),value= self.search_column[0][0])


        self.BondTypeTree = TreeCtrl(parent =bkg,id = wx.NewId(), pos=(ANCHOR,ANCHOR+(HEIGHT+SPACE)*4),
                                     size =(WIDTH*2,HEIGHT*4.5),root=u"全部类型",items=self.bond_types)
        self.CRCompanyTree = TreeCtrl(parent =bkg,id = wx.NewId(),pos=(ANCHOR+WIDTH*2+SPACE,ANCHOR+(HEIGHT+SPACE)*4),
                                     size =(WIDTH*2,HEIGHT*4.5),root=u"全部主体评级",items=self.company_ratings)
        self.CRBondTree = TreeCtrl(parent =bkg,id = wx.NewId(),pos=(ANCHOR+WIDTH*2*2+SPACE*2,ANCHOR+(HEIGHT+SPACE)*4),
                                     size =(WIDTH*2,HEIGHT*4.5),root=u"全部债券评级",items=self.bond_ratings)
        self.AgencyTree = TreeCtrl(parent =bkg,id = wx.NewId(), pos=(ANCHOR+WIDTH*2*3+SPACE*3,ANCHOR+(HEIGHT+SPACE)*4),
                                     size =(WIDTH*2,HEIGHT*4.5),root=u"全部中介",items=self.agencies)

        self.BondTypeTree.ExpandAll()
        self.CRCompanyTree.ExpandAll()
        self.CRBondTree.ExpandAll()
        self.AgencyTree.ExpandAll()


        self.label0 = wx.StaticText(bkg, -1, u"关键字", pos=(ANCHOR, ANCHOR+SPACE))
        self.label1 = wx.StaticText(bkg, -1, u"开始时间", pos=(ANCHOR, 60))
        self.label2 = wx.StaticText(bkg, -1, u"结束时间", pos=(260, 60))
        self.label3 = wx.StaticText(bkg, -1, u"最低利率", pos=(ANCHOR, 90))
        self.label4 = wx.StaticText(bkg, -1, u"最高利率", pos=(260, 90))
        self.label5 = wx.StaticText(bkg, -1, u"最小期限", pos=(ANCHOR, 120))
        self.label6 = wx.StaticText(bkg, -1, u"最大期限", pos=(260, 120))
        self.StartDateText = wx.TextCtrl(bkg,-1,size=(WIDTH*1.7,HEIGHT),pos = (ANCHOR+WIDTH+SPACE,60),value=get_time(1))
        self.EndDateText   = wx.TextCtrl(bkg,-1,size=(WIDTH*1.7,HEIGHT), pos = (350,60),value=get_time(1))
        self.MaxPriceText = wx.TextCtrl(bkg,-1,size=(WIDTH*1.7,HEIGHT),pos = (350,90),value = "5.80")
        self.MinPriceText   = wx.TextCtrl(bkg,-1,size=(WIDTH*1.7,HEIGHT), pos = (ANCHOR+WIDTH+SPACE,90),value="2.51")
        self.MaxTermText = wx.TextCtrl(bkg,-1,size=(WIDTH,HEIGHT),pos = (350,120),value = "10")
        self.MinTermText   = wx.TextCtrl(bkg,-1,size=(WIDTH,HEIGHT), pos = (ANCHOR+WIDTH+SPACE,120),value="1")
        self.term_unit_cb1 = wx.ComboBox(bkg,id= wx.NewId(),choices=self.term_units[0],
                                         size=(WIDTH*0.7,HEIGHT),pos =(430,120),value=u"年")
        self.term_unit_cb2 = wx.ComboBox(bkg,id= wx.NewId(),choices=self.term_units[0],
                                         size=(WIDTH*0.7,HEIGHT),pos =(ANCHOR+(WIDTH)*2+SPACE,120),value=u"日")

        self.Bind(wx.EVT_CLOSE,self.OnClose)

        self.txtpath   = ""
        self.xlpath    = ""
        self.xlpath_ex = ""
        self.checked_items =[]
        self.Show(True)
        self.data = []
        self.export_data= []
        self.date      = ""
        self.queue = Queue()
        self.success_collection=[]
        self.fail_collection =[]

    def MultiThread(self,data,func,args):
        startTime = get_time(2)
        num_thread = 10
        print func.__doc__
        for item in data:
            self.queue.put(item)

        for i in range(num_thread):
            worker = Thread(target=func,args=args)
            worker.setDaemon(True)
            worker.start()

        worker = Thread(target=self.initGauge(self.queue.qsize()))
        worker.setDaemon(True)
        worker.start()

        self.queue.join()
        print "--------Start multi Thread at", startTime
        print "--------End multi Thread working at,", get_time(2)


    def InsertLineToDB(self,host, user, passwd, db):
        '''
        -----To insert a line in to data base online-----
        '''
        while not self.queue.empty():
            content = self.queue.get()
            conn = MySQLdb.connect(host=host, user=user, passwd=passwd, db=db)
            cursor = conn.cursor()

            try:

                cursor.execute("INSERT INTO " + content[0] + " VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s);", content[1:])
                self.success_collection.append(content[1:])

            except Exception as err:
                print "InsertLineToDB : Somthing went wrong"
                print err
                self.fail_collection.append(content[1:])

            cursor.close()
            conn.commit()
            self.queue.task_done()

    def initGauge(self,range):
        progressMax = range
        dialog = wx.ProgressDialog(u"数据导入进度",u"导入进度",maximum=progressMax,parent=None,style=wx.PD_AUTO_HIDE|wx.PD_ELAPSED_TIME | wx.PD_REMAINING_TIME)
        count = 0
        while count < progressMax:
            count = range - self.GetCount()
            wx.Sleep(0.5)
            dialog.Update(count)

        dialog.Destroy()

    def GetCount(self):
        if self.queue.qsize():
            return self.queue.qsize()
        else:
            return 0

    def OnClose(self,e):
        try:
            if self.connection:
                self.connection.close()
        except Exception as err:
            self.ErrDialog(err,"On Close ")

        self.Destroy()
        e.Skip()
        sys.exit(0)

    def OnDelData(self,e):
        dialog = wx.MessageDialog(None, u"确定要从数据库删除这条记录？删除之后数据将无法恢复", u"提醒", wx.YES_NO | wx.ICON_QUESTION)
        if dialog.ShowModal() == wx.ID_YES:
            print "get current selected Range"
            select_info = self.xlsFrame.GetOffset((0,8))
            select_row = self.xlsFrame.GetCurrentlySelectedRows()
            tables = [ "tr"+str(item[0]).replace("-","")[0:6] for item in select_info ]
            database_ids = [ str(item[1]) for item in select_info ]
            data_info = ((tables[i],database_ids[i]) for i in range(len(tables)))

            try:
                if len(tables) <500:
                    for i in range(len(tables)):
                        if del_row_table(conn=self.connection, id=int(database_ids[i]),table=tables[i]):
                                print "delete id= " + database_ids[i] + " from database successfully"
                else:
                    self.MultiThread(data=data_info,func=self.DeleteDataFromDB,
                                         args=(self.host,self.username,self.password,self.db))
            except Exception as err:
                self.connection.rollback()
                print("Something went wrong: {}".format(err))
                errdialog = wx.MessageDialog(None, u"数据删除失败: \n {}".format(err), u"错误提示", wx.ICON_QUESTION)
                if errdialog.ShowModal() == wx.ID_OK:
                    if self.fail_collection:
                        self.fail_collection.insert(0, database_title)
                        fail_xlsFrame = XLFrame(self.fail_collection, title=u"删除失败的数据", export_func=self.OnExport)
                        fail_xlsFrame.Show()
                    errdialog.Destroy()

            self.fail_collection = []
            self.success_collection = []

            temp_data = self.xlsFrame.GetData()
            self.xlsFrame.Destroy()
            export_data = []
            for i in range(len(temp_data)):
                if i not in select_row:
                    export_data.append(temp_data[i])

            self.xlsFrame = XLFrame(export_data, self.OnExport, self.OnDelData)
            self.xlsFrame.Show()



    def DeleteDataFromDB(self,host, user, passwd, db):
        '''
        ------To delete data line by line from data base online------
        '''
        while not self.queue.empty():
            content = self.queue.get()
            conn = MySQLdb.connect(host=host, user=user, passwd=passwd, db=db)
            cursor = conn.cursor()

            try:
                cursor.execute("DELETE FROM %s WHERE ID = %s" % content)
                print "Successfully DELETE FROM %s WHERE ID = %s " % content

            except Exception as err:
                print "DeleteDataFromDB : Somthing went wrong"
                print err
                self.fail_collection.append(content)

            cursor.close()
            conn.commit()
            self.queue.task_done()


    def OnGetData(self,e):
        print self.search_text.GetValue()

        search_result = []
        search_column =self.search_column[1][self.search_column[0].index(self.search_column_cb.GetValue())]

        tables,filter = self.GetFilter(type=1)

        for table in tables:
            try:
                search_result += search_table(self.connection, search_column, self.search_text.GetValue(), table=table,
                                              filter=filter)
            except Exception as err :
                self.ErrDialog(err,"On Get Data ")
        self.data = search_result
        self.GetData()


    def OnExport(self,e):
        wildcard = u"Excel 文件(*.xls)|.xls"
        dialog = wx.FileDialog(None, "Save an Excel file...",wildcard=wildcard, style=wx.SAVE)
        if dialog.ShowModal() == wx.ID_OK:
            self.xlpath_ex = dialog.GetPath()#.encode('utf-8')
            export_excel(self.export_data,self.xlpath_ex)

    def OnImportTxt(self, e):
        datedlg = wx.TextEntryDialog(None,  u"请输入成交日期","", get_time(1))
        if datedlg.ShowModal() == wx.ID_OK:
            date = datedlg.GetValue()
            if IsDate(date):
                self.date = date
                dialog = wx.FileDialog(None, "Choose a txt file...", style=wx.OPEN)
                if dialog.ShowModal() == wx.ID_OK:
                    self.txtpath = dialog.GetPath()#.encode('utf-8')
                    dialog.Destroy()
                    datedlg.Destroy()

                    wildcard = u"Excel 文件(*.xls)|.xls"
                    dialog2 = wx.FileDialog(None, "Save an Excel file...", wildcard=wildcard, style=wx.SAVE)
                    if dialog2.ShowModal() == wx.ID_OK:
                        self.xlpath = dialog2.GetPath()#.encode('utf-8')
                        import_text(self.txtpath,self.xlpath,date=date)
                        dialog2.Destroy()


    def OnImportExcel(self, e):
        dialog = wx.FileDialog(None, "Choose an excel file...", style=wx.OPEN)
        fail_collection = []
        err_exist = False
        cursor = self.connection.cursor()
        table = ""
        data = []

        if dialog.ShowModal() == wx.ID_OK:
            self.xlpath = dialog.GetPath()#.encode('utf-8')
            dialog.Destroy()

            try:
                #从excel表中导入成交数据
                temp_data,fail_collection = import_excel(self.xlpath,self.connection)
                for item in temp_data:
                    temp = "tr%s" % item[0][0:6]
                    # 如果table已存在，最大id为table里的最大id，否则创建新的table，最大id为0
                    if table != temp:
                        table = temp
                        if not IsTableExist(self.connection, table):
                            create_table(self.connection, table)
                            max_id = 0
                        else:
                            cursor.execute("SELECT MAX(id) FROM %s " % table)
                            max_id = cursor.fetchone()[0]
                            if max_id == None: max_id = 0

                        i = max_id + 1

                    data.append((table,i)+tuple(item))
                    i += 1

                #多线程插入数据至aliyun数据库
                self.MultiThread(data,self.InsertLineToDB,(self.host,self.username,self.password,self.db))
            except Exception as err:
                self.ErrDialog(err,"On Import Excel")
                err_exist = True

            # try:
            #     success_collection,fail_collection = import_excel(xlpath= self.xlpath,conn=self.connection)
            # except Exception as err:
            #     print("OnImport Excel | Something went wrong: {}".format(err))
            #     errdlg = wx.MessageDialog(None, u"错误发生: \n {}".format(err), u"错误提示", wx.ICON_QUESTION)
            #     err_exist = True
            #     if errdlg.ShowModal() == wx.ID_OK:
            #         errdlg.Destroy()
            #
            #

            self.fail_collection += fail_collection

            if self.fail_collection:
                self.fail_collection.insert(0,database_title)
                fail_xlsFrame = XLFrame(self.fail_collection, title=u"导入失败的数据", export_func=self.OnExport)
                fail_xlsFrame.Show()
                successdlg = wx.MessageDialog(None, u"已经导入 %s 条数据"%len(self.success_collection), u"提示")
                if successdlg.ShowModal() == wx.ID_OK:
                    successdlg.Destroy()
            else:
                if not err_exist:
                    dialog = wx.MessageDialog(None, u"数据已经全部成功导入数据库 \n 一共导入 %s 条数据"%len(self.success_collection),  u"提醒", wx.YES_NO)
                    dialog.ShowModal()

            if self.success_collection and err_exist:
                self.success_collection.insert(0,excel_title)
                suc_xlsFrame = XLFrame(self.success_collection, title=u"导入成功的数据", export_func=self.OnExport,menu_func=self.OnDelData)
                suc_xlsFrame.Show()

            self.success_collection = []
            self.fail_collection = []
            cursor.close()


    def ErrDialog(self,err,context):
        print(context+" Something went wrong: {}".format(err))
        if err.args[0] == 2006:
            errinfo = u" 数据库连接失败"
        else:
            errinfo = u" 出错: \n {}".format(err)

        errdialog = wx.MessageDialog(None, errinfo, u"错误提示", wx.ICON_QUESTION)
        if errdialog.ShowModal() == wx.ID_OK:
            errdialog.Destroy()

    def GetFilter(self,type=0):
        #如果type等于默认值0，返回完整的执行语句，如果type等于1，返回筛选条件语句
        self.filter = ""

        max_price = self.MaxPriceText.GetValue()
        min_price = self.MinPriceText.GetValue()

        start_date = self.StartDateText.GetValue()
        end_date = self.EndDateText.GetValue()

        max_term = self.MaxTermText.GetValue()
        max_unit =self.term_units[1][self.term_units[0].index(self.term_unit_cb1.GetValue())]

        min_term = self.MinTermText.GetValue()
        min_unit = self.term_units[1][self.term_units[0].index(self.term_unit_cb2.GetValue())]

        bond_types = self.BondTypeTree.get_checked_item()
        company_ratings = self.CRCompanyTree.get_checked_item()
        bond_ratings = self.CRBondTree.get_checked_item()
        agencies = self.AgencyTree.get_checked_item()

        selected_tables = []

        if IsDate(start_date,end_date) and IsNumber(min_term,max_term):
            s_datetime = datetime.strptime(start_date,"%Y-%m-%d")
            e_datetime = datetime.strptime(end_date, "%Y-%m-%d")

            selected_tables.extend("tr"+ x for x in monthdelta(s_datetime,e_datetime ))
            try:
                exist_tables = get_tables(self.connection)
            except Exception as err:
                self.ErrDialog(err,"Get Filter ")
                exist_tables = []

            tables =list(set(selected_tables) & set(exist_tables))
            tables.sort(reverse=True)

            self.filter =  " (%s >= price) AND (price >= %s)"%(max_price,min_price)
            self.filter += " AND ( %s >= date) AND (date >=%s)"%(end_date.replace("-", ""),start_date.replace("-", ""))
            self.filter += " AND ( %d >= term) AND (term>= %d)"%(float(max_term) * float(max_unit),float(min_term)*float(min_unit))

            selected_items = [[bond_types,company_ratings,bond_ratings,agencies],
                              [self.bond_types,self.company_ratings,self.bond_ratings,self.agencies],
                              ["type","company_rating","bond_rating","agency"]]

            for i in range(len(selected_items[0])):
                if selected_items[0][i] != selected_items[1][i]:
                    self.filter += " AND ("
                    for item in selected_items[0][i][:-1]:
                        self.filter += "(%s = '%s') OR "%(selected_items[2][i],item)
                    self.filter +="(%s = '%s'))"%(selected_items[2][i],selected_items[0][i][-1])

            #print self.filter
            filter_col =[]

            if type ==0:
                for table in tables:
                    filter_col.append("SELECT %s FROM %s Where %s order by date desc"%(select_columns,table,self.filter))
                return filter_col
            elif type ==1:
                return (tables,self.filter)


    def GetData(self):
        export_data = []
        export_data.append(excel_title)

        for i in range(len(self.data)):
            temp = str(self.data[i][0])
            if (len(temp) == 8):
                temp = temp[0:4] + "-" + temp[4:6] + "-" + temp[6:8]
                export_data.append([])
                export_data[i + 1] = []
                export_data[i + 1].append(temp)
            for j in range(1, len(self.data[i])):
                export_data[i + 1].append(self.data[i][j])
        self.export_data = export_data

        if export_data:
            self.xlsFrame = XLFrame(export_data, self.OnExport, self.OnDelData)
            self.xlsFrame.Show()


    def Connect_MySQL(self):
        username = ''
        password = ''
        db =''

        db_dlg = wx.TextEntryDialog(None, u"请输入数据库", "", '')
        if db_dlg.ShowModal()== wx.ID_OK:
            db = db_dlg.GetValue()
            user_dlg = wx.TextEntryDialog(None, u"请输入用户名", "", '')
            if user_dlg.ShowModal() == wx.ID_OK:
                username = user_dlg.GetValue()
                pwd_dlg = wx.TextEntryDialog(None, u"请输入密码", "", '')
                if pwd_dlg.ShowModal() == wx.ID_OK:
                    password = pwd_dlg.GetValue()

                    if username and password and db:
                        try:
                            conn = MySQLdb.connect(
                                host=".mysql.rds.aliyuncs.com",
                                port=3306,
                                user=username,
                                passwd=password,
                                db=db,
                                charset='utf8')
                            print "Successfully connect to MySQL on Aliyun"
                            self.host = "htzqbonds.mysql.rds.aliyuncs.com"
                            self.username = username
                            self.password = password
                            self.db =db
                            return conn
                        except Exception as err:
                            print("Something went wrong: {}".format(err))
                            if err.args[0]==1045:
                                errinfo = u"连接远程数据库失败: \n用户名或密码错误"
                            elif err.args[0]==1044:
                                errinfo = u"连接远程数据库失败: \n没有权限访问该数据库"
                            else:
                                errinfo = u"连接远程数据库失败: \n {}".format(err)

                            errdlg = wx.MessageDialog(None, errinfo+"\n是否重新登录？", u"错误提示", wx.YES_NO|wx.ICON_QUESTION)
                            if errdlg.ShowModal() == wx.ID_YES:
                                errdlg.Destroy()
                                return self.Connect_MySQL()



class TreeCtrl(CT.CustomTreeCtrl):
    def __init__(self,parent,id,root,items,pos=wx.DefaultPosition,size=wx.DefaultSize,style=wx.TR_DEFAULT_STYLE):
        CT.CustomTreeCtrl.__init__(self,parent,id,pos,size,style)
        self.root = self.AddRoot(root,ct_type=1)
        for item in items:
            self.AppendItem(self.root,item,ct_type=1)
        self.Bind(CT.EVT_TREE_ITEM_CHECKED, self.OnChecked)
        self.checked_items=[]
        self.CheckItem(self.root,True)

    def OnChecked(self,e):
        checked_item = e.GetItem()
        if (checked_item==self.root):
            if self.IsItemChecked(checked_item):
                self.CheckChilds(self.root)
                for item in self.get_tree_children(self.root):
                    self.checked_items.append(self.GetItemText(item))
            else:
                for item in self.get_tree_children(self.root):
                    self.CheckItem(item,False)
                self.checked_items =[]
        else:
            if self.IsItemChecked(checked_item):
                self.checked_items.append(self.GetItemText(checked_item))
                # print "add"
            else:
                if self.GetItemText(checked_item) in self.checked_items:
                    self.checked_items.remove(self.GetItemText(checked_item))
                if self.IsItemChecked(self.root):
                    self.CheckItem(self.root,False)
                    for item in self.get_tree_children(self.root):
                        self.CheckItem(item, True)
                    self.CheckItem(checked_item,False)


    def get_tree_children(self,item_obj):
        item_list = []
        (item,cookie) = self.GetFirstChild(item_obj)
        while item:
            item_list.append(item)
            # print "OK "
            (item,cookie) = self.GetNextChild(item_obj,cookie)
        return item_list

    def get_tree_child(self,item_obj,index):
        item_list = []
        (item,cookie) = self.GetFirstChild(item_obj)
        while item:
            item_list.append(item)
            (item,cookie) = self.GetNextChild(item_obj,cookie)
        return item_list[index]

    def get_checked_item(self):
        return self.checked_items



class XLFrame(wx.Frame):
    def __init__(self,data,export_func=None, menu_func = None,title =u"提取数据结果"):
        wx.Frame.__init__(self, parent=None, title=title, size=(800,600))
        panel = wx.Panel(self)
        self.data = data
        nrow = len(data)
        ncol = len(data[0])+5
        self.myGrid = gridlib.Grid(panel)
        self.myGrid.CreateGrid(nrow, ncol)
        self.myGrid.Bind(gridlib.EVT_GRID_SELECT_CELL, self.onSingleSelect)
        self.myGrid.Bind(gridlib.EVT_GRID_CELL_RIGHT_CLICK,self.onSingleSelect)
        self.myGrid.Bind(gridlib.EVT_GRID_RANGE_SELECT, self.onDragSelection)

        if menu_func:
            self.myGrid.Bind(gridlib.EVT_GRID_CELL_RIGHT_CLICK, self.showPopupMenu)
            self.menu_func = menu_func

        for i in range(len(data)):
            for j in range(len(data[i])):
                temp = ""
                try:
                    temp = data[i][j].decode('utf-8')
                    self.myGrid.SetCellValue(i, j, temp)
                except:
                    temp = data[i][j]
                    try:
                        temp = str(temp)
                        self.myGrid.SetCellValue(i, j, temp)
                    except:
                        pass

        export_button = wx.Button(panel,label=u"导出")
        export_button.Bind(wx.EVT_BUTTON,export_func)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(self.myGrid, 1, wx.EXPAND)
        sizer.Add(export_button,0)
        panel.SetSizer(sizer)

        self.currentlySelectedRange = []
        self.currentlySelectedCell = ()

    def showPopupMenu(self, e):
        menu = wx.Menu()
        item = wx.MenuItem(menu, wx.NewId(), u"从数据库内移除")
        menu.AppendItem(item)
        menu.Bind(wx.EVT_MENU,self.menu_func,item)
        self.PopupMenu(menu)
        menu.Destroy()


    def onDragSelection(self, e):
        if self.myGrid.GetSelectionBlockTopLeft():
            top_left = self.myGrid.GetSelectionBlockTopLeft()[0]
            bottom_right = self.myGrid.GetSelectionBlockBottomRight()[0]
            self.currentlySelectedRange = self.GetSelectedCells(top_left, bottom_right)


    def onSingleSelect(self, e):
        self.currentlySelectedCell = (e.GetRow(),e.GetCol())
        # print "current selected cell " + str(self.currentlySelectedCell)
        e.Skip()


    def GetSelectedCells(self, top_left, bottom_right):
        cells = []

        rows_start = top_left[0]
        rows_end = bottom_right[0]

        cols_start = top_left[1]
        cols_end = bottom_right[1]

        rows = range(rows_start, rows_end + 1)
        cols = range(cols_start, cols_end + 1)

        cells.extend([(row, col)
                      for row in rows
                      for col in cols])
        return cells


    def GetCurrentlySelectedCell(self):
        return self.currentlySelectedCell


    def GetCurrentlySelectedRange(self):
        if self.currentlySelectedRange:
            return self.currentlySelectedRange
        else:
            return [self.currentlySelectedCell]


    def GetOffset(self,offset):
        offsetRangeValue = []
        if self.GetCurrentlySelectedRange():
            for cell in self.GetCurrentlySelectedRange():
                temp=[]
                for i in offset:
                    temp.append(self.myGrid.GetCellValue(cell[0],cell[1]+i))
                offsetRangeValue.append(temp)
        return offsetRangeValue


    def GetCurrentlySelectedRows(self):
        return [x[0] for x in self.GetCurrentlySelectedRange()]


    def GetCellValue(self,row,col):
        return self.myGrid.GetCellValue(row,col)


    def GetData(self):
        return self.data


if __name__ == "__main__":
    print get_time()
    app = wx.App(False)
    frame = MainWindow(None, u'债券成交信息数据库')
    app.MainLoop()
