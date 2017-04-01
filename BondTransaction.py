# -*- coding: utf-8 -*-
# mypath should be the complete path for the directory containing the input text files

import xlwt
import xlrd
import wx
import wx.grid as gridlib
import time
import sqlite3
import re
import sys
import os
import pickle
import wx.lib.agw.customtreectrl as CT

reload(sys)
sys.setdefaultencoding('utf-8')
excel_title = [u"成交时间",u"期限",u"债券代码",u"债券简称",u"利率",u"信用评级",u"类型",u"中介机构"]

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
            temp_row.append(date)
            for item in row:
                value = item.strip()
                temp_row.append(value.decode('gb2312'))
            temp_row.append(bond_type)
            temp_row.append(agency)
            export_data.append(temp_row)
    export_excel(export_data,xlpath)

def import_excel(xlpath, dbpath):
    print "-------import from excel----------"
    book = xlrd.open_workbook(xlpath)
    sheet = book.sheet_by_index(0)
    nrow = sheet.nrows
    row_list = []
    for i in range(1,nrow):
        #print sheet.row_values(i)
        temp =adjust_row(sheet.row_values(i))
        try:
            row_list.append(tuple(temp))
        except:
            pass
    print row_list
    insert_table(row_list,dbpath)


def export_excel(data,xlpath):
    print "-------export to excel----------"
    workbook = xlwt.Workbook()
    worksheet = workbook.add_sheet('Sheet1')
    for j in range(len(excel_title)):
        worksheet.write(0,j,excel_title[j])
    for i in range(1,len(data)):
        for j in range(len(data[i])):
            worksheet.write(i,j,data[i][j])
    workbook.save(xlpath)

def create_table(open_path):
    conn = sqlite3.connect(open_path)
    cursor = conn.cursor()
    try:
        cursor.execute('''CREATE TABLE TR (
                        date long,
                        term_text text,
                        bond_id text,
                        name text,
                        price_text char(50),
                        rating char(50),
                        type char(50),
                        agency char(50),
                        term int,
                        price real);''')
        conn.commit()
        print "-------create table successfully--------"
    except:
        print "fail to create table"
        cursor.execute("DROP TABLE TR")
        create_table(open_path)
    conn.close()

def insert_table(data, open_path):
    print "------insert_table------"
    conn = sqlite3.connect(open_path)
    cursor = conn.cursor()
    for item in data:
        try:
            cursor.execute("INSERT INTO TR VALUES(?,?,?,?,?,?,?,?,?,?);",item)
        except:
            print "fail to insert"
            print len(item)
    conn.commit()
    conn.close()

def select_table(open_path,filter_clause= "SELECT * FROM TR "):
    conn = sqlite3.connect(open_path)
    cursor = conn.cursor()
    try:
        cursor.execute(filter_clause)
        return cursor.fetchall()
    except:
        return []

def get_time(type = 0):
    if type==0:
        return time.strftime('%Y%m%d',time.localtime(time.time()))
    elif type ==1:
        return time.strftime('%Y-%m-%d',time.localtime(time.time()))

def adjust_row(data):
    print "-------adjust row data----------"
    date = data[0]
    date = int(date.replace("-",""))#insert(0,int(date.replace("-","")))
    adjusted_data = []
    adjusted_data.append(date)

    for item in data[1:]:
        if item!="":
            adjusted_data.append(item)
    re_term = u"[0-9.]+[DMYdmy]"
    pattern = re.compile(re_term)
    result = re.match(pattern,data[1])

    term = ""
    term2 = ""
    if result != None:
        term = result.group(0)
        print "term " + term
        print "data[1]" + data[1]
        re_term_plus = u"[+][0-9.]+[DMYdmy]"
        pattern_plus = re.compile(re_term_plus)
        result_plus = re.search(pattern_plus,data[1])

        if result_plus!=None:
            term2 = result_plus.group(0)
            print "term  "  + term + " term2 " + term2
        adjusted_data.append(StrToDays(term)+StrToDays(term2))

    re_price = u"[0-9]{1,2}[.]+[0-9]+"
    pattern = re.compile(re_price)
    result = re.match(pattern,data[4])
    if result!= None:
        price = float(result.group(0))
        adjusted_data.append(price)

    return adjusted_data

def StrToDays(term):
    term_in_days = 0
    if ("Y" in term) or ("y" in term):
        term_in_days = int(float(term[:len(term) - 1]) * 360)
    elif ("D" in term) or ("d" in term):
        term_in_days = int(term[:len(term) - 1])
    elif ("M" in term) or ("m" in term):
        term_in_days = int(float(term[:len(term) - 1]) * 30)
    return term_in_days

def rating_index(rating):
    rating_list = ["AAA+","AAA","AAA-","AA+","AA","AA-","BBB+","BBB","BBB-","BB+", "BB","BB-","B+","B","B-"]
    if rating in rating_list:
        return rating_list.index(rating)

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
        dlg = wx.MessageDialog(None, u"日期格式错误", u"错误提示", wx.YES_NO | wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_YES:
            dlg.Destroy()
        return False

def IsNumber(*number):
    try:
        for item in number:
            float(item)
        return True
    except:
        dlg = wx.MessageDialog(None, u"数字错误", u"错误提示", wx.YES_NO | wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_YES:
            dlg.Destroy()
        return False


class MainWindow(wx.Frame):
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title = title,size = (550,300))
        self.gaugeFrame = GaugeFrame()
        #self.control = wx.TextCtrl(self, style=wx.TE_MULTILINE)
        bkg = wx.Panel(self,style=wx.TAB_TRAVERSAL | wx.CLIP_CHILDREN | wx.FULL_REPAINT_ON_RESIZE)
        txt_button = wx.Button(bkg, label = u"导入 txt", pos=(20,20))
        txt_button.Bind(wx.EVT_BUTTON, self.onImportTxt)

        xl_button = wx.Button(bkg, label=u"导入excel", pos=(120,20))
        xl_button.Bind(wx.EVT_BUTTON, self.onImportExcel)

        ex_button = wx.Button(bkg, label=u"导出excel", pos=(220,20))
        ex_button.Bind(wx.EVT_BUTTON, self.onExport)

        getdata_button = wx.Button(bkg, label=u"提取数据", pos = (320,20))
        getdata_button.Bind(wx.EVT_BUTTON, self.onGetData)

        db_button = wx.Button(bkg,label = u"数据库操作", pos =(420,20))
        db_button.Bind(wx.EVT_BUTTON,self.OnDB)


        #wx.StaticText(bkg, -1, u"债券类型", pos=(20, 120))
        #wx.StaticText(bkg, -1, u"信用评级", pos=(260, 120))
        self.bond_types = [u"短融",u"企业债", u"公司债",u"存单",u"中票",u"其他"]
        self.ratings = ["AAA+", "AAA", "AAA-", "AA+", "AA", "AA-", "BBB+", "BBB", "BBB-", "BB+", "BB", "BB-", "B+", "B","B-"]
        self.agencies =[u"平安信用",u"平安利率",u"BGC信用",u"国际信用",u"国际利率",u"国利信用",u"国利利率",u"信唐"]
        self.term_units=[[u"年",u"月",u"日"],[360,30,1]]
        self.BondTypeTree = TreeCtrl(parent =bkg,id = wx.NewId(), pos=(20,160),
                                     size =(150,100),root=u"全部类型",items=self.bond_types)
        self.CRTree = TreeCtrl(parent =bkg,id = wx.NewId(),pos=(200,160),
                                     size =(150,100),root=u"全部评级",items=self.ratings)
        self.AgencyTree = TreeCtrl(parent =bkg,id = wx.NewId(), pos=(380,160),
                                     size =(150,100),root=u"全部中介",items=self.agencies)
        self.label1 = wx.StaticText(bkg, -1, u"开始时间", pos=(20, 60))
        self.label2 = wx.StaticText(bkg, -1, u"结束时间", pos=(260, 60))
        self.label3 = wx.StaticText(bkg, -1, u"最低利率", pos=(20, 90))
        self.label4 = wx.StaticText(bkg, -1, u"最高利率", pos=(260, 90))
        self.label5 = wx.StaticText(bkg, -1, u"最小期限", pos=(20, 120))
        self.label6 = wx.StaticText(bkg, -1, u"最大期限", pos=(260, 120))
        self.StartDateText = wx.TextCtrl(bkg,-1,size=(150,25),pos = (80,60),value=get_time(1))
        self.EndDateText   = wx.TextCtrl(bkg,-1,size=(150,25), pos = (350,60),value=get_time(1))
        self.MaxPriceText = wx.TextCtrl(bkg,-1,size=(150,25),pos = (350,90),value = "4.80")
        self.MinPriceText   = wx.TextCtrl(bkg,-1,size=(150,25), pos = (80,90),value="4.51")
        self.MaxTermText = wx.TextCtrl(bkg,-1,size=(70,25),pos = (350,120),value = "10")
        self.MinTermText   = wx.TextCtrl(bkg,-1,size=(70,25), pos = (80,120),value="1")
        self.term_unit_cb1 = wx.ComboBox(bkg,id= wx.NewId(),choices=self.term_units[0],
                                         size=(70,25),pos =(430,120))
        self.term_unit_cb2 = wx.ComboBox(bkg,id= wx.NewId(),choices=self.term_units[0],
                                         size=(70,25),pos =(160,120),value=u"日")

        self.txtpath   = ""
        self.xlpath    = ""
        self.xlpath_ex = ""
        self.filter    ={}
        self.checked_items =[]
        self.Show(True)
        self.data = []
        self.export_data= []
        self.date      = ""

        try:
            self.dbpath = self.GetDBs()[0]
        except:
            dialog = wx.MessageDialog(None, u"暂无数据库，请新建至少一个数据库", u"提醒", wx.YES_NO | wx.ICON_QUESTION)
            if dialog.ShowModal() == wx.ID_YES:
                self.CreateDB()



    def OnDB(self,e):
        self.DBFrame = wx.Frame(None,title=u"数据库操作", size = (300,300))
        self.DBFrame.Show()
        p = wx.Panel(self.DBFrame,size =(500,300))
        create_db_btn = wx.Button(p,label=u"新建数据库",pos=(20,20),size=(140,30))
        choose_db_btn = wx.Button(p,label = u"选择默认数据库",pos=(20,60),size=(140,30))
        del_db_btn = wx.Button(p,label = u"删除数据库",pos=(20,100),size=(140,30))
        create_db_btn.Bind(wx.EVT_BUTTON,self.OnCreateDB)
        choose_db_btn.Bind(wx.EVT_BUTTON, self.OnChooseDefaultDB)
        del_db_btn.Bind(wx.EVT_BUTTON, self.OnDelDB)

    def OnCreateDB(self,e):
        self.CreateDB()

    def OnChooseDefaultDB(self,e):
        self.ChooseDefaultDB()

    def OnDelDB(self,e):
        choose_dlg = wx.SingleChoiceDialog(None,message=u"请选择要删除的数据库",caption= u"数据库操作", choices=list(self.GetDBs()))
        if choose_dlg.ShowModal() == wx.ID_OK:
            del_db= choose_dlg.GetStringSelection()
            dialog = wx.MessageDialog(None, u"确定删除数据库？所有数据将被删除，不可修复。", u"警告", wx.YES_NO | wx.ICON_QUESTION)
            if dialog.ShowModal() == wx.ID_YES:
                self.DelDB(del_db)
                dialog.Destroy()
                choose_dlg.Destroy()
        else:
            choose_dlg.Destroy()

    def ChooseDefaultDB(self):
        if self.GetDBs() ==():
            dialog = wx.MessageDialog(None, u"暂无数据库，请新建至少一个数据库", u"提醒", wx.YES_NO | wx.ICON_QUESTION)
            if dialog.ShowModal() == wx.ID_YES:
                if(self.CreateDB()):
                    self.ChooseDefaultDB()
        else:
            choose_dlg = wx.SingleChoiceDialog(None,message=u"请选择默认使用的数据库",caption= u"数据库操作", choices=list(self.GetDBs()))
            if choose_dlg.ShowModal() == wx.ID_OK:
                chosen_db = choose_dlg.GetStringSelection()
                self.SetDefaultDB(chosen_db)
                self.dbpath = chosen_db
                choose_dlg.Destroy()

    def CreateDB(self):
        dialog = wx.TextEntryDialog(None, u"请输入数据库名称(英文)..", "","tr" )
        if dialog.ShowModal() == wx.ID_OK:
            dbpath= dialog.GetValue() +".db"
            if self.AddDB(dbpath):
                create_table(dbpath)


    def DelDB(self, del_db):
        temp = self.GetDBs()
        dbs = ()
        for db in temp:
            if db != del_db:
                dbs += (db,)
        print "Del db " + del_db
        self.SetDBs(dbs)

    def AddDB(self,add_db):
        temp = self.GetDBs()
        if add_db in temp:
            dlg = wx.MessageDialog(None, u"数据库已存在", u"错误提示", wx.YES_NO | wx.ICON_QUESTION)
            if dlg.ShowModal() == wx.ID_YES:
                dlg.Destroy()
            return False
        else:
            dbs = temp + (add_db,)
            self.SetDBs(dbs)
            print "Add db " + add_db
            return True


    def GetDBs(self):
        try:
            dbs = pickle.load(open('dbs.pkl', 'rb'))
            print "Get dbs " + str(dbs)
            return dbs
        except:
            return ()

    def SetDefaultDB(self, chosen_db):
        temp = self.GetDBs()
        dbs = (chosen_db,)
        for db in temp:
            if db!= chosen_db:
                dbs += (db,)
        self.SetDBs(dbs)
        print "Set default db " + chosen_db

    def SetDBs(self,dbs):
        try:
            pickle.dump(dbs, open('dbs.pkl', 'wb'))
            print "Set dbs " + str(dbs)
        except:
            print "fail to set dbs"

    def onGetData(self,e):
        self.data = select_table(self.dbpath,self.GetFilter())
        export_data = []
        export_data.append(excel_title)
        for i in range(len(self.data)):
            temp = str(self.data[i][0])
            if (len(temp) == 8):
                temp = temp[0:4] + "-" + temp[4:6] + "-" + temp[6:8]
                export_data.append([])
                export_data[i+1] =[]
                export_data[i+1].append(temp)
            for j in range(1,len(self.data[0])-2):
                export_data[i+1].append(self.data[i][j])
        self.export_data = export_data
        try:
            self.xlsFrame = XLFrame(export_data)
            self.xlsFrame.Show()
        except:
            pass

    def onExport(self,e):
        wildcard = u"Excel 文件(*.xls)|.xls|"
        dialog = wx.FileDialog(None, "Save an Excel file...",wildcard=wildcard, style=wx.SAVE)
        if dialog.ShowModal() == wx.ID_OK:
            self.xlpath_ex = dialog.GetPath()#.encode('utf-8')
            export_excel(self.export_data,self.xlpath_ex)


    def onImportTxt(self, e):
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

                    wildcard = u"Excel 文件(*.xls)|.xls|"
                    dialog2 = wx.FileDialog(None, "Save an Excel file...", wildcard=wildcard, style=wx.SAVE)
                    if dialog2.ShowModal() == wx.ID_OK:
                        self.xlpath = dialog2.GetPath()#.encode('utf-8')
                        import_text(self.txtpath,self.xlpath,date=date)
                        dialog2.Destroy()


    def onImportExcel(self, e):
        dialog = wx.FileDialog(None, "Choose an excel file...", style=wx.OPEN)
        if dialog.ShowModal() == wx.ID_OK:
            self.xlpath = dialog.GetPath()#.encode('utf-8')
            import_excel(xlpath= self.xlpath,dbpath=self.dbpath)
            dialog.Destroy()

    def GetFilter(self):
        max_price = self.MaxPriceText.GetValue()
        min_price = self.MinPriceText.GetValue()

        start_date = self.StartDateText.GetValue()
        end_date = self.EndDateText.GetValue()

        max_term = self.MaxTermText.GetValue()
        max_unit =self.term_units[1][self.term_units[0].index(self.term_unit_cb1.GetValue())]

        min_term = self.MinTermText.GetValue()
        min_unit = self.term_units[1][self.term_units[0].index(self.term_unit_cb2.GetValue())]

        bond_types = self.BondTypeTree.get_checked_item()
        ratings = self.CRTree.get_checked_item()
        agencies = self.AgencyTree.get_checked_item()


        if IsDate(start_date,end_date) and IsNumber(min_term,max_term):
            self.filter = " SELECT * FROM TR WHERE (" + max_price + " >= price) AND (price >= " + min_price + ")"
            self.filter += " AND (" + end_date.replace("-", "") + " >= date) AND( date >= " + start_date.replace("-", "") + ")"
            self.filter += "AND (" + str(int(float(max_term) * float(max_unit))) + ">= term) AND ( term>= " + str(int(float(min_term)*float(min_unit))) +")"

            if (bond_types != self.bond_types) and(len(bond_types)!=0):
                self.filter += " AND ( "
                for bond in bond_types:
                    self.filter += " ( type = '" + str(bond) + "' ) "
                    if bond!= bond_types[-1]:
                        self.filter +=  " OR "
                    else:
                        self.filter +=" ) "

            if (ratings != self.ratings) and (len(ratings)!=0):
                self.filter += " AND ( "
                for rating in ratings:
                    self.filter += " ( rating = '" + str(rating) + "' ) "
                    if rating!= ratings[-1]:
                        self.filter +=  " OR "
                    else:
                        self.filter += " ) "

            if (agencies != self.agencies) and (len(ratings)!=0):
                self.filter += " AND ( "
                for agency in agencies:
                    self.filter += " ( agency = '" + str(agency) + "' ) "
                    if agency!= agencies[-1]:
                        self.filter +=  " OR "
                    else:
                        self.filter += " ) "

            print self.filter
        return self.filter

class TreeCtrl(CT.CustomTreeCtrl):
    def __init__(self,parent,id,root,items,pos=wx.DefaultPosition,size=wx.DefaultSize,style=wx.TR_DEFAULT_STYLE):
        #CT.CustomTreeCtrl.__init__(panel=panel,pos=pos,agwStyle=wx.TR_DEFAULT_STYLE)
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
                for item in self.checked_items:
                    self.checked_items.remove(item)
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
                # print "remove"
        print self.checked_items

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
            # print "OK "
            (item,cookie) = self.GetNextChild(item_obj,cookie)
        return item_list[index]

    def get_checked_item(self):
        return self.checked_items

class XLFrame(wx.Frame):
    def __init__(self,data):
        """Constructor"""
        nrow = len(data)
        ncol = len(data[0])
        wx.Frame.__init__(self, parent=None, title="data from data base", size=(600,400))
        panel = wx.Panel(self)

        myGrid = gridlib.Grid(panel)
        myGrid.CreateGrid(nrow, ncol)

        for i in range(len(data)):
            for j in range(len(data[i])):
                temp = ""
                try:
                    temp = data[i][j]
                    myGrid.SetCellValue(i, j, temp)
                except:
                    if (type(data[i][j]) == type(1)) or (type(data[i][j] == type(1.11))):
                        temp = str(temp)
                        myGrid.SetCellValue(i, j, temp)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(myGrid, 1, wx.EXPAND)
        panel.SetSizer(sizer)

class GaugeFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, -1, 'Gauge Example',
                          size=(350, 150))
        panel = wx.Panel(self, -1)
        self.count = 0
        self.gauge = wx.Gauge(panel, -1, 50, (20, 50), (250, 25))
        self.gauge.SetBezelFace(3)
        self.gauge.SetShadowWidth(3)
        self.Bind(wx.EVT_IDLE, self.OnIdle)

    def OnIdle(self, event):
        self.count = self.count + 1
        if self.count == 50:
            self.Hide()
        self.gauge.SetValue(self.count)

if __name__ == "__main__":
    print get_time()
    app = wx.App(False)
    frame = MainWindow(None, u'信息导入系统')
    app.MainLoop()
