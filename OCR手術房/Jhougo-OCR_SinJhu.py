# -*- coding: utf-8 -*-
#                       _oo0oo_
#                      o8888888o
#                      88" . "88
#                      (| -_- |)
#                      0\  =  /0
#                    ___/`---'\___
#                  .' \\|     |// '.
#                 / \\|||  :  |||// \
#                / _||||| -:- |||||- \
#               |   | \\\  -  /// |   |
#               | \_|  ''\---/''  |_/ |
#               \  .-\__  '-'  ___/-. /
#             ___'. .'  /--.--\  `. .'___
#          ."" '<  `.___\_<|>_/___.' >' "".
#         | | :  `- \`.;`\ _ /`;.`/ - ` : | |
#         \  \ `_.   \_ __\ /__ _/   .-` /  /
#     =====`-.____`.___ \_____/___.-`___.-'=====
#                       `=---='
#
#
#     ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
#
#               佛祖保佑         永無BUG
import os
import wx
import wx.lib.filebrowsebutton as filebrowse
import xlsxwriter
import Image
import ImageEnhance
import pytesseract
import time
import string
Dir = ""
da_v = False
ti_v = False
li_v = False
do_v = False
up_v = False
####存list違規路段
ocrfilepath = os.getcwd().decode('big5')+'\\path_OCR_SinJhu\\path_ocr'
file1 = file(ocrfilepath,'r')
content1 = file1.read()
file1.close()
content1=content1.replace(";"," ")
piece1 = string.split(content1)
y = len(piece1)
##########ico圖片路徑
icopath = os.getcwd()+'\\path_OCR_SinJhu\\te4.ico'
# -----------------------------------------YES or NO---------------------------------------------------------------
class TestPanelocr(wx.Panel):
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent)
        self.log = log
        
        a = wx.Button(self, 40, u'忽略辨識過的!', (270,40), style=wx.NO_BORDER)     
        a.SetToolTipString("YES.\n")
        b = wx.Button(self, 50, u'全部重新辨識!', (370,40), style=wx.NO_BORDER)
        b.SetToolTipString("NO.\n")
        self.Bind(wx.EVT_BUTTON, self.OnClick, a)
        self.Bind(wx.EVT_BUTTON, self.OnClick, b) 
        # self.txtx = wx.StaticText(self,label=u"是否忽略之前已辨識過的圖片?", pos=(10, 25))  #一樣方式產生文字
        str1 = u"是否忽略之前已辨識過的圖片?"
        text = wx.StaticText(self, -1, str1, (10, 45))
        font = wx.Font(13,wx.ROMAN, wx.NORMAL, wx.NORMAL)
        text.SetFont(font)
        # str2 = "You can also change the font."
        # text1 = wx.StaticText(self, -1, str2, (10, 15))
        # font1 = wx.Font(18,wx.DECORATIVE, wx.ITALIC, wx.NORMAL)
        # text1.SetFont(font1)
        # sizer = wx.BoxSizer(wx.VERTICAL)  
        # sizer.Add(self.txtx, 0, wx.ALL, 20)
        # box = wx.BoxSizer()
        # box.Add(sizer, 0, wx.ALL, 20)
        # self.SetSizer(box)

    def OnClick(self, event):
        eid = event.GetId()
        if eid == 40:
            apprun = Apprun()
            apprun.MainLoop()
        else:
            app = App()
            app.MainLoop()
class Frameocr( wx.Frame ):
    def __init__( self, parent ):
        wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = u"OCR相片辨識", pos = wx.DefaultPosition, size = wx.Size(500,150))
        panel = TestPanelocr(self, -1)

class Appocr(wx.App):
    def OnInit(self):
        self.frame = Frameocr(parent=None)
        icon = wx.EmptyIcon()
        icon.CopyFromBitmap(wx.Bitmap(icopath, wx.BITMAP_TYPE_ANY))
        self.frame.SetIcon(icon)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True
# -----------------------------------------------------------------------------------------------------------------
# -----------------------------------------xlsx or ocr or 檔案修正路徑---------------------------------------------------------------
class TestPanelfirst(wx.Panel):
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent)
        self.log = log

        c = wx.Button(self, 60, u'檔案路徑修正!', (35,25), style=wx.NO_BORDER,size=(120,60))
        c.SetToolTipString(u"檔案路徑修正.\n")
        a = wx.Button(self, 40, u'xlsx建檔!', (185,25), style=wx.NO_BORDER,size=(120,60))     
        a.SetToolTipString(u"xlsx建檔.\n")
        b = wx.Button(self, 50, u'OCR辨識!', (335,25), style=wx.NO_BORDER,size=(120,60))
        b.SetToolTipString(u"OCR辨識.\n")
        a.SetFont(wx.Font(13, wx.ROMAN, wx.NORMAL, wx.NORMAL))
        b.SetFont(wx.Font(13, wx.ROMAN, wx.NORMAL, wx.NORMAL))
        c.SetFont(wx.Font(13, wx.ROMAN, wx.NORMAL, wx.NORMAL))
        self.Bind(wx.EVT_BUTTON, self.OnClick, a)
        self.Bind(wx.EVT_BUTTON, self.OnClick, b) 
        self.Bind(wx.EVT_BUTTON, self.OnClick, c) 

    def OnClick(self, event):
        eid = event.GetId()
        if eid == 50:
            appocr = Appocr()
            appocr.MainLoop()
        elif eid == 40:
            appxlsx = Appxlsx()
            appxlsx.MainLoop()
        elif eid == 60:
            appdis = Appdis()
            appdis.MainLoop()
class Framefirst( wx.Frame ):
    def __init__( self, parent ):
        wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = u"選擇", pos = wx.DefaultPosition, size = wx.Size(500,150))
        panel = TestPanelfirst(self, -1)

class Appfirst(wx.App):
    def OnInit(self):
        self.frame = Framefirst(parent=None)
        icon = wx.EmptyIcon()
        icon.CopyFromBitmap(wx.Bitmap(icopath, wx.BITMAP_TYPE_ANY))
        self.frame.SetIcon(icon)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True
# -----------------------------------------------------------------------------------------------------------------
# ------------------------------------Generate TXT----------------------------------------------- 
class TestPanel(wx.Panel):
    def listpath(e,dirpath):
        tStart = time.time()
        count = 0
        for root, dirs, files in os.walk(dirpath):
            for f in files:
                if ".jpg" in os.path.join(root , f):
                    strsu =  os.path.join(f).replace('.jpg','')
                    strf =  os.path.join(root)
                    pathr = os.path.join(root , f).replace('.jpg','')
                    isexists = os.path.exists(pathr)
                    if not isexists :
                        if u'紅燈' in strf or u'不依標誌標線' in strf or u'未依標線之指示' in strf :
                            if strsu[len(strsu)-2:len(strsu)] == "F1" :
                                os.mkdir(pathr)
                                count = count+1
                                mainomli(strsu,strf,count)
                        if u'快速道路' in strf : 
                            if strsu[len(strsu)-2:len(strsu)] == "F1" :
                                os.mkdir(pathr)
                                count = count+1
                                mainomli2(strsu,strf,count)
                        if  u'紅燈' not in strf and  u'標線' not in strf and u'號誌' not in strf : 
                            count = count+1
                            os.mkdir(pathr)
                            mainom(strsu,strf,count) 
                    else:
                        if u'紅燈' in strf or u'不依標誌標線' in strf or u'未依標線之指示' in strf : 
                            if strsu[len(strsu)-2:len(strsu)] == "F1" :
                                count = count+1
                                mainomli(strsu,strf,count)
                        if u'快速道路' in strf : 
                            if  strsu[len(strsu)-2:len(strsu)] == "F1" :
                                count = count+1
                                mainomli2(strsu,strf,count)
                        if  u'紅燈' not in strf and  u'標線' not in strf and u'號誌' not in strf : 
                            count = count+1
                            mainom(strsu,strf,count)
        tEnd = time.time()
        print "Produce  "+str(count)+"  files  spend  "+str((tEnd - tStart)//1)+"  second" 
                        
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent)
        self.log = log
        self.dbb = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'圖片根目錄:', changeCallback = self.dbbCallback)
        b = wx.Button(self, 40, u'RUN', (500,28), style=wx.NO_BORDER)
        b.SetToolTipString("Run OCR\n")
        self.Bind(wx.EVT_BUTTON, self.OnClick, b)
        sizer = wx.BoxSizer(wx.VERTICAL)        
        sizer.Add(self.dbb, 0, wx.ALL, 5)
        box = wx.BoxSizer()
        box.Add(sizer, 0, wx.ALL, 20)
        b.SetInitialSize()  
        self.SetSizer(box)

        x1 = wx.CheckBox(self, 1, u"日期", (35, 100), (150, 20))  
        x2 = wx.CheckBox(self, 2, u"時間", (35, 120), (150, 20))  
        x3 = wx.CheckBox(self, 3, u"車牌", (35, 140), (150, 20))
        x4 = wx.CheckBox(self, 4, u"速限", (35, 160), (150, 20))
        x5 = wx.CheckBox(self, 5, u"速度", (35, 180), (150, 20))
        self.Bind(wx.EVT_CHECKBOX, self.OnClick2, x1)
        self.Bind(wx.EVT_CHECKBOX, self.OnClick2, x2)
        self.Bind(wx.EVT_CHECKBOX, self.OnClick2, x3)
        self.Bind(wx.EVT_CHECKBOX, self.OnClick2, x4)
        self.Bind(wx.EVT_CHECKBOX, self.OnClick2, x5)

    def dbbCallback(self, event):
        Dir = self.dbb.GetValue()
        self.path = Dir
        event.Skip()
    def OnClick2(self, event):
        global da_v,ti_v,li_v,up_v,do_v   
        eid = event.GetId()
        if eid == 1 and da_v == False:
            da_v = True
        elif eid == 2 and ti_v == False:
            ti_v = True
        elif eid == 3 and li_v == False:
            li_v = True
        elif eid == 4 and up_v == False:
            up_v = True
        elif eid == 5 and do_v == False:
            do_v = True
        elif eid == 1 and da_v == True:
            da_v = False
        elif eid == 2 and ti_v == True:
            ti_v = False
        elif eid == 3 and li_v == True:
            li_v = False
        elif eid == 4 and up_v == True:
            up_v = False
        elif eid == 5 and do_v == True:
            do_v = False

    def OnClick(self, event):         
        self.listpath(self.path)
    
        
class Frame ( wx.Frame ):
    def __init__( self, parent ):
        wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = u"重頭到尾跑一遍", pos = wx.DefaultPosition, size = wx.Size(700,250))
        panel = TestPanel(self, -1)

class App(wx.App):
    def OnInit(self):
        self.frame = Frame(parent=None)
        icon = wx.EmptyIcon()
        icon.CopyFromBitmap(wx.Bitmap(icopath, wx.BITMAP_TYPE_ANY))
        self.frame.SetIcon(icon)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True
# -------------------------------------------------------------------------------------------------
# ---------------------------------------------Generate TXT 忽略已有TXT的-------------------------------------------
class TestPanelrun(wx.Panel):
    def listpath(e,dirpath):
        tStart = time.time()
        count = 0
        for root, dirs, files in os.walk(dirpath):
            for f in files:
                if ".jpg" in os.path.join(root , f):
                    strsu =  os.path.join(f).replace('.jpg','')
                    strf =  os.path.join(root)
                    pathr = os.path.join(root , f).replace('.jpg','')
                    pathrtxt = os.path.join(root , f).replace('.jpg','')+u'\\'+strsu+u".txt"
                    isexists = os.path.exists(pathr)
                    isexiststxt = os.path.exists(pathrtxt)
                    if not isexiststxt :
                        if not isexists :
                            if u'紅燈' in strf or u'不依標誌標線' in strf or u'未依標線之指示' in strf : 
                                if strsu[len(strsu)-2:len(strsu)] == "F1" :
                                    os.mkdir(pathr)
                                    count = count+1
                                    mainomli(strsu,strf,count)
                            if u'快速道路' in strf : 
                                if strsu[len(strsu)-2:len(strsu)] == "F1" :
                                    os.mkdir(pathr)
                                    count = count+1
                                    mainomli2(strsu,strf,count)
                            if  u'紅燈' not in strf and  u'標線' not in strf and u'號誌' not in strf : 
                                count = count+1
                                os.mkdir(pathr)
                                mainom(strsu,strf,count) 
                        else:
                            count = count+1
                            if u'紅燈' in strf or u'不依標誌標線' in strf or u'未依標線之指示' in strf : 
                                if strsu[len(strsu)-2:len(strsu)] == "F1" :
                                    mainomli(strsu,strf,count)
                            if u'快速道路' in strf : 
                                if  strsu[len(strsu)-2:len(strsu)] == "F1" :
                                    mainomli2(strsu,strf,count)
                            if  u'紅燈' not in strf and  u'標線' not in strf and u'號誌' not in strf : 
                                mainom(strsu,strf,count)

        tEnd = time.time()
        print "Produce  "+str(count)+"  files  spend  "+str((tEnd - tStart)//1)+"  second" 
                        
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent)
        self.log = log
        self.dbb = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'圖片根目錄:', changeCallback = self.dbbCallback)
        b = wx.Button(self, 40, u'RUN', (500,28), style=wx.NO_BORDER)
        b.SetToolTipString("Run OCR\n") 
        self.Bind(wx.EVT_BUTTON, self.OnClick, b)
        sizer = wx.BoxSizer(wx.VERTICAL)        
        sizer.Add(self.dbb, 0, wx.ALL, 5)
        box = wx.BoxSizer()
        box.Add(sizer, 0, wx.ALL, 20)
        b.SetInitialSize()  
        self.SetSizer(box)

        x1 = wx.CheckBox(self, 1, u"日期", (35, 100), (150, 20))  
        x2 = wx.CheckBox(self, 2, u"時間", (35, 120), (150, 20))  
        x3 = wx.CheckBox(self, 3, u"車牌", (35, 140), (150, 20))
        x4 = wx.CheckBox(self, 4, u"速限", (35, 160), (150, 20))
        x5 = wx.CheckBox(self, 5, u"速度", (35, 180), (150, 20))
        self.Bind(wx.EVT_CHECKBOX, self.OnClick2, x1)
        self.Bind(wx.EVT_CHECKBOX, self.OnClick2, x2)
        self.Bind(wx.EVT_CHECKBOX, self.OnClick2, x3)
        self.Bind(wx.EVT_CHECKBOX, self.OnClick2, x4)
        self.Bind(wx.EVT_CHECKBOX, self.OnClick2, x5)
        # wx.Panel.__init__(self, parent)
        # self.log = log
        # self.dbb = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'TXT建檔:', changeCallback = self.dbbCallback)
        # self.txtx = wx.StaticText(self,label=u"建檔人: ")  
        # self.text = wx.TextCtrl(self, size=(200, 20))   
        # self.txtx1 = wx.StaticText(self,label=u"冊號: ")  
        # self.text1 = wx.TextCtrl(self, size=(200, 20)) 
        # self.txtx2 = wx.StaticText(self,label=u"違規地點: ")  
        # self.text2 = wx.TextCtrl(self, size=(200, 20)) 
        # b = wx.Button(self, 40, u'RUN', (500,28), style=wx.NO_BORDER)
        # b.SetToolTipString("This button has a style flag of wx.NO_BORDER.\n"
        #                    "On some platforms that will give it a flattened look.")
        # self.Bind(wx.EVT_BUTTON, self.OnClick, b)
        # sizer = wx.BoxSizer(wx.VERTICAL)        
        # sizer.Add(self.dbb, 0, wx.ALL, 5)
        # sizer.Add(self.txtx, 0, wx.ALL, 5)
        # sizer.Add(self.text, 0, wx.ALL, 5)
        # sizer.Add(self.txtx1, 0, wx.ALL, 5)
        # sizer.Add(self.text1, 0, wx.ALL, 5)
        # sizer.Add(self.txtx2, 0, wx.ALL, 5)
        # sizer.Add(self.text2, 0, wx.ALL, 5)
        # sizer.Add(b, 0, wx.ALL, 30)
        # box = wx.BoxSizer()
        # box.Add(sizer, 0, wx.ALL, 20)
        # b.SetInitialSize()  
        # self.SetSizer(box)

    def dbbCallback(self, event):
        Dir = self.dbb.GetValue()
        self.path = Dir
        event.Skip()

    def OnClick2(self, event):
        global da_v,ti_v,li_v,up_v,do_v   
        eid = event.GetId()
        if eid == 1 and da_v == False:
            da_v = True
        elif eid == 2 and ti_v == False:
            ti_v = True
        elif eid == 3 and li_v == False:
            li_v = True
        elif eid == 4 and up_v == False:
            up_v = True
        elif eid == 5 and do_v == False:
            do_v = True
        elif eid == 1 and da_v == True:
            da_v = False
        elif eid == 2 and ti_v == True:
            ti_v = False
        elif eid == 3 and li_v == True:
            li_v = False
        elif eid == 4 and up_v == True:
            up_v = False
        elif eid == 5 and do_v == True:
            do_v = False


    def OnClick(self, event): 
        # f = file(self.path+'.txt', 'a+') 
        # f.write(self.text.GetValue()+";"+self.text1.GetValue()+";"+self.text2.GetValue()+ "\n") # write text to file
        # f.close()       
        self.listpath(self.path)
        

class Framerun ( wx.Frame ):
    def __init__( self, parent ):
        wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = u"只跑沒有跑過的相片", pos = wx.DefaultPosition, size = wx.Size(700,150))
        panel = TestPanelrun(self, -1)

class Apprun(wx.App):
    def OnInit(self):
        self.frame = Framerun(parent=None)
        icon = wx.EmptyIcon()
        icon.CopyFromBitmap(wx.Bitmap(icopath, wx.BITMAP_TYPE_ANY))
        self.frame.SetIcon(icon)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True
# -----------------------------------------------------------------------------------------------------------------
# ---------------------------------------------格式修正-------------------------------------------
class TestPaneldis(wx.Panel):
    def listpath(e,dirpath): 
        num1 = True   
        tStart = time.time() 
        c = 0                     
        for root, dirs, files in os.walk(dirpath):
            for f in files:
                strf =  os.path.join(root)
                if ".jpg" in os.path.join(root , f):
                    if u'紅燈' in strf or u'標線' in strf or u'標誌' in strf or u'快速道路' in strf : 
                        if os.path.join(f).find("F1")<0 and os.path.join(f).find("F2")<0:
                            if c%2==0 : 
                                r1 = os.path.join(root , f)
                                r2 = os.path.join(root , f).replace('.jpg','_F1.jpg')
                                os.rename(r1,r2)
                            if c%2==1 : 
                                r1 = os.path.join(root , f)
                                r2 = os.path.join(root , f).replace('.jpg','_F2.jpg')
                                os.rename(r1,r2)
                            c=c+1
        for root, dirs, files in os.walk(dirpath):        
            for d in dirs:
                strf1 =  os.path.join(root,d)
                if strf1.find("(")>-1 or strf1.find(")")>-1 :
                    mainomli3(strf1)

        tEnd = time.time()
        print "Success"
        print " spend  "+str((tEnd - tStart)//1)+"  second" 
                        
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent)
        self.log = log
        self.dbb = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'圖片根目錄:', changeCallback = self.dbbCallback)
        b = wx.Button(self, 40, u'RUN', (500,28), style=wx.NO_BORDER)
        b.SetToolTipString("Run OCR\n") 
        self.Bind(wx.EVT_BUTTON, self.OnClick, b)
        sizer = wx.BoxSizer(wx.VERTICAL)        
        sizer.Add(self.dbb, 0, wx.ALL, 5)
        box = wx.BoxSizer()
        box.Add(sizer, 0, wx.ALL, 20)
        b.SetInitialSize()  
        self.SetSizer(box)
    def dbbCallback(self, event):
        Dir = self.dbb.GetValue()
        self.path = Dir
        event.Skip()
    def OnClick(self, event):    
        self.listpath(self.path)
        

class Framedis ( wx.Frame ):
    def __init__( self, parent ):
        wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = u"修正路徑 資料夾", pos = wx.DefaultPosition, size = wx.Size(700,150))
        panel = TestPaneldis(self, -1)

class Appdis(wx.App):
    def OnInit(self):
        self.frame = Framedis(parent=None)
        icon = wx.EmptyIcon()
        icon.CopyFromBitmap(wx.Bitmap(icopath, wx.BITMAP_TYPE_ANY))
        self.frame.SetIcon(icon)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True
# -----------------------------------------------------------------------------------------------------------------
# ----------------------------------------------------xlsx建檔-------------------------------------------------------------
 
class TestPanelxlsx(wx.Panel):
    def listpath(e,dirpath):
        addr = ""
        camID = ""
        datatime = "" 
        path = ""
        couse = ""
        count = 0      
        IDcount = 0
        nowpath = ""
        case = 0
        mutilphoto = False
        booknum = ['0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','G','H','K','L','M','N','P','Q','R','S','T','U','V','W','X','Y','Z']
        xlsxname = dirpath[dirpath.rfind('\\')+1:]+'.xlsx'
        #xlsxname = u'固定桿資料.xlsx'
        row = 1;
        IDcount =0;

        #print dirpath
        print dirpath[dirpath.rfind(u"\\")+1:]+".xlsx"
        workbook = xlsxwriter.Workbook(xlsxname)
        bold = workbook.add_format({'bold': True})
        worksheet = workbook.add_worksheet()
        worksheet.write('A1',u'文件夾', bold)
        worksheet.write('B1',u'檔名', bold)
        worksheet.write('C1',u'測照地點', bold)
        worksheet.write('D1',u'舉發樣態', bold)
        worksheet.write('E1',u'圖檔數', bold)
        worksheet.write('F1',u'筆數', bold)
        worksheet.write('G1',u'冊號數', bold)
        worksheet.write('H1',u'備註', bold)
        for root, dirs, files in os.walk(dirpath):
            count = 0
            if root == dirpath:
                path = root.replace(dirpath,'')
            else:               
                path = root.replace(dirpath+"\\",'')               
            if "\\" in path:

                camID = path[path.find("\\")+1:]               
                for f in files:
                        if ".jpg" in os.path.join(root, f):
                            count = count+1
                
                if count!=0:
                    if "\\" in camID:                                                           
                        mutilphoto = True                        
                        worksheet.write(row,3,camID[camID.find("\\")+1:])
                        camID = camID[:camID.rfind("\\")]               
                    
                    else :
                        mutilphoto = False
                        worksheet.write(row,3,u'超速')
                    addr = path[:path.find("\\")]
                    if nowpath!=addr:
                        worksheet.write(row,0,addr) 
                    addr2 = addr[:len(addr)-11]
                    for x in range(y):
                        str1 = piece1[x]
                        str2 = addr2.encode('utf-8')
                        if str1 == str2 :
                            content = piece1[x+1]
                    worksheet.write(row,2,content.decode('UTF-8'))
                    worksheet.write(row,1,camID)
                    IDcount = IDcount+1          
                    worksheet.write(row,4,count)
                    if(mutilphoto):
                        case = count/2
                    else:
                        case = count
                    
                    worksheet.write(row,5,case)
                    if u'慧珍' in root:
                        who = u'4'                       
                    else:
                        who = u'6'
                    date = dirpath[dirpath.rfind("\\")+3:dirpath.rfind("\\")+8]                    
                    worksheet.write(row,6,"9P"+date+who+booknum[IDcount])                    
                    
                    if u'-C' in camID:
                        worksheet.write(row,7,u'文字檔')
                else:
                    row=row-1

                row = row+1               
                nowpath=addr
        workbook.close()
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent)
        self.log = log

        self.dbb = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'建檔根目錄:', changeCallback = self.dbbCallback)
        b = wx.Button(self, 40, u'開始建立xlsx檔', (355,70), style=wx.NO_BORDER)
        b.SetToolTipString("This button has a style flag of wx.NO_BORDER.\n"
                           "On some platforms that will give it a flattened look.")
        self.Bind(wx.EVT_BUTTON, self.OnClick, b)
        
        sizer = wx.BoxSizer(wx.VERTICAL)        
        sizer.Add(self.dbb, 0, wx.ALL, 5)
        box = wx.BoxSizer()
        box.Add(sizer, 0, wx.ALL, 20)
        b.SetInitialSize()  
        self.SetSizer(box)

    def dbbCallback(self, event):
        Dir = self.dbb.GetValue()
        self.path = Dir
        os.chdir(Dir[:Dir.rfind(u"\\")+1])
        event.Skip()
    def OnClick(self, event):         
        self.listpath(self.path)
    
        
class Framexlsx ( wx.Frame ):
    def __init__( self, parent ):
        wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = u"xlsx建檔", pos = wx.DefaultPosition, size = wx.Size(500, 150))
        panel = TestPanelxlsx(self, -1)


class Appxlsx(wx.App):
    def OnInit(self):
        self.frame = Framexlsx(parent=None)
        icon = wx.EmptyIcon()
        icon.CopyFromBitmap(wx.Bitmap(icopath, wx.BITMAP_TYPE_ANY))
        self.frame.SetIcon(icon)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True

def runTest(frame, nb, log):
    win = TestPanelxlsx(nb, -1, log)
    return win
    
# -------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------OCR image-------------------------------------------------------
def omocr(img_name,path1,co):#ex:20150512_124104_906_1794_
   # print "A"
    content = ""
    patnum = path1[path1.rfind(u'\\'):len(path1)] #rfind find from right
    patnum1 = path1.replace(patnum, u'&')
    patnum2 = patnum1[patnum1.rfind(u'\\')+1:patnum1.rfind(u'&')-11]
    ispath = os.path.exists(u'path_OCR\\path_ocr')
   # print ispath
    print u"第"+str(co)+u"筆 "

    for x in range(y):
        str1 = piece1[x]
        str2 = patnum2.encode('utf-8')
        if str1 == str2 :
            content = piece1[x+1]
            print content.decode('UTF-8')  

    strsp = path1+"\\"+img_name
    im = Image.open(strsp+u'.jpg').convert('L')
    isExists = os.path.exists(strsp)
    if isExists :
        day2 = u' '
        time2 = u' '
        spdown2 = u' '
        spup2 = u' '
        if da_v == True:
            im.crop((235, 20, 640, 112)).save(strsp+u'\\day.png')
            day = pytesseract.image_to_string(Image.open(strsp+u'\\day.png')).replace('/','').replace(" ", "").replace("O", "0")
            day=filter(str.isdigit, day)
            day2 = day[0:8]
            if day[0] == "2":
                day2 = int(day2)-19110000
        if ti_v == True:
            im.crop((237, 118, 446, 198)).save(strsp+u'\\time.png')
            time = pytesseract.image_to_string(Image.open(strsp+u'\\time.png')).replace(':','').replace(" ", "").replace("O", "0")
            time=filter(str.isdigit, time)
            time2 = time[0:4]
        if do_v == True:
            im.crop((854, 13, 1148, 109)).save(strsp+u'\\spdown.png')
            spdown = pytesseract.image_to_string(Image.open(strsp+u'\\spdown.png')).replace(" ", "").replace("O", "0")
            spdown=filter(str.isdigit, spdown)
            spdown2 = spdown[0:3]
        if up_v == True:
            im.crop((859, 112, 1147, 196)).save(strsp+u'\\spup.png')
            spup = pytesseract.image_to_string(Image.open(strsp+u'\\spup.png')).replace(" ", "").replace("O", "0")
            spup=filter(str.isdigit, spup)
            spup2 = spup[0:3]
        if li_v == True:
            im.crop((1, 1400, 610, 1710)).save(strsp+u'\\li.png')

        f = file(strsp+'\\'+img_name+'.txt', 'w+')
        f.write(img_name+u';'+str(day2)+u';'+str(time2)+u';'+str(spup2)+u';'+str(spdown2)+u';') # write text to file
        if ispath :
            f.write(content)#寫入抓到的照相地點
        else:
            f.write(' ; ')
        f.close()
        print u'檔 名:'+img_name 
        print u'日 期:'+str(day2)
        print u'時 間:'+time2
        print u'速 限:'+spup2
        print u'車 速:'+spdown2
        print "OK"


###########################################################################
def omocrwr(path1):#ex:20150512_124104_906_1794_
    sssss = 0
    running = True
    running2 = True
    if path1.find("(")>-1 or path1.find(")")>-1 :
       # print "YA"
        while(running):
           # print "YA2"
            while(running2):
                if  path1.rfind("\\")<path1.find("(") or path1.rfind("\\")<path1.find(")")  : 
                    running2 = False 
                    sssss = 1
                   # print "YA3"
                if  path1.find("\\")<path1.find("(") and path1.rfind("\\")>path1.find("(")  or path1.find("\\")<path1.find(")") and path1.rfind("\\")>path1.find(")"): 
                    path1 = path1.replace("\\",";",1)
                   # print "YA4"
                if  path1.find("\\")>path1.find("(") or path1.find("\\")>path1.find(")") :   
                    running2 = False 
                    #print "YA5"
            running = False 
        if sssss == 0:
            path1 = path1[0:path1.find("\\")]
            path2 = path1.replace("(", '').replace(")", '').replace(";","\\")
           # print "YA6"
            os.rename(path1,path2)
        if sssss == 1:
            path2 = path1.replace("(", '').replace(")", '')
           # print "YA7"
            os.rename(path1,path2)


###########################################################################
def mainom(img,pa,co):
    if img.find("_")>12 :
        print " "
        omocr(img,pa,co)

def mainomli3(pa):   
    print " "
    omocrwr(pa)
# ------------------------------------------------------------------------------------------------------------------------
if __name__ == '__main__':
    appfirst = Appfirst()
    appfirst.MainLoop()
# if __name__ == '__main__':
#     app = App()
#     app.MainLoop()
# ------------------------------------------------------------------------------------------------------------------------



