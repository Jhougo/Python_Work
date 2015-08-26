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
#####################################################################
#########################   天祐新竹   ##############################
#####################################################################
import os
import wx
import wx.lib.filebrowsebutton as filebrowse
import xlsxwriter
import Image
import sys
import ImageEnhance
import pytesseract
import time
import string
import subprocess
import copy
from dbfpy import dbf 
import reportlab
from reportlab.lib import *
from reportlab.pdfbase import *
from reportlab.pdfgen import *
from reportlab.platypus import *
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib import fonts
from reportlab.lib.styles import getSampleStyleSheet
APP_OCR=1  #定义一个控件ID 
APP_PATH=1  #定义一个控件ID 
APP_xlsx=1  #定义一个控件ID 
APP_EXIT=1  #定义一个控件ID 
contPATH = 0
PATH = [] 
PATH1 = [] 
da_v = False
ti_v = False
li_v = False
do_v = False
up_v = False
a = 0
Dir = ""
Dir1 = ""
piece1  = ['start']
piece2  = ['start']
photo_path = []
story = []
story2 = []
st2 = []
Color = ""

####存list違規路段
ocrfilepath = os.getcwd().decode('big5')+'..\\path_OCR\\path_ocr'
file1 = file(ocrfilepath,'r')
content1 = file1.read()
file1.close()
content1=content1.replace(";"," ")
piece1 = string.split(content1)
y = len(piece1)
##########ico圖片路徑
icopath = os.getcwd()+'\\path_OCR\\te4.ico'  
# -----------------------------------------------------------------------------------------------------------------
class Example(wx.Frame): 
  def __init__(self, parent, id, title): 
    super(Example,self).__init__(parent, id, title)    
  
    self.InitUI()      
  
  def InitUI(self):  
  
    menubar = wx.MenuBar()    #生成選單 
    filemenu = wx.Menu()     
    filemenu2 = wx.Menu()    
    filemenu4 = wx.Menu()    
    filemenu5 = wx.Menu()    
    filemenu3 = wx.Menu()

    qmi2 = wx.MenuItem(filemenu, 1, u'檔案路徑修正!')   #   
    qmi2.SetBitmap(wx.Bitmap('path_OCR\\i2.ico' ))    
    qmi3 = wx.MenuItem(filemenu, 2, u'xlsx建檔!')   #   
    qmi3.SetBitmap(wx.Bitmap('path_OCR\\i3.ico' ))  

    qmi = wx.MenuItem(filemenu3, 4, "Quit")   #
    qmi.SetBitmap(wx.Bitmap('path_OCR\\i1.ico' ))  

    q1 = wx.MenuItem(filemenu2, 5, u'跳過辨識過的!')   #   
    q1.SetBitmap(wx.Bitmap('path_OCR\\i4.ico' ))    
    q2 = wx.MenuItem(filemenu2, 6, u'全部重新辨識!')   #   
    q2.SetBitmap(wx.Bitmap('path_OCR\\i5.ico' ))    

    qq = wx.MenuItem(filemenu4, 10, u'啟動固定桿建檔!')   #   
    qq.SetBitmap(wx.Bitmap('path_OCR\\i7.jpg' ))

    qpd = wx.MenuItem(filemenu5, 11, 'DBF-->PDF!')   #   
    qpd.SetBitmap(wx.Bitmap('path_OCR\\i8.ico' )) 

    filemenu.AppendItem(qmi2)       
    filemenu.AppendItem(qmi3)
    filemenu3.AppendItem(qmi)
    filemenu4.AppendItem(qq)
    filemenu5.AppendItem(qpd)

    filemenu2.AppendItem(q1) 
    filemenu2.AppendItem(q2)      
  
    menubar.Append(filemenu, "&File")    
    menubar.Append(filemenu2, "&OCR")    
    menubar.Append(filemenu4, u"&固定桿建檔")    
    menubar.Append(filemenu5, u"&製作清冊")
    menubar.Append(filemenu3, "&Quit")     
    self.SetMenuBar(menubar)      

    self.Bind(wx.EVT_MENU, self.PATH, id=1)   
    self.Bind(wx.EVT_MENU, self.PATH, id=2)   
    self.Bind(wx.EVT_MENU, self.PATH, id=4)   
    self.Bind(wx.EVT_MENU, self.PATH, id=5)   
    self.Bind(wx.EVT_MENU, self.PATH, id=6)   
    self.Bind(wx.EVT_MENU, self.PATH, id=10)  
    self.Bind(wx.EVT_MENU, self.PATH, id=11)  

    self.SetSize((800, 600))      
    self.SetTitle(u"固定桿") 
    self.Centre() 
    self.SetBackgroundColour('white') 

  
    self.Show(True)    #顯示框架 
  
###########################################################################
  def PATH(self, e): 
    eid = e.GetId()
    global a

    if eid == 1:
      if a == 1:
        self.dbb.Destroy()
        self.b.Destroy()
        a = 0  
      if a == 2:
        self.dbb1.Destroy()
        self.b1.Destroy()
        a = 0 
      if a == 5:
        self.dbbo.Destroy()
        self.bo.Destroy()
        self.x11.Destroy()
        self.x22.Destroy()
        self.x33.Destroy()
        self.x44.Destroy()
        self.x55.Destroy()
        a = 0 
      if a == 6:
        self.dbboo.Destroy()
        self.boo.Destroy()
        self.x1.Destroy()
        self.x2.Destroy()
        self.x3.Destroy()
        self.x4.Destroy()
        self.x5.Destroy()
        a = 0 
      if a == 11:
        self.dbbp1.Destroy()
        self.dbbp2.Destroy()
        self.bp.Destroy()
        a = 0 
      self.dbb = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'圖片根目錄:', changeCallback = self.dbbCallback)
      self.b = wx.Button(self, 40, 'RUN', (480,3), style=wx.NO_BORDER)
      self.b.SetToolTipString(u"格式更正\n") 
      self.Bind(wx.EVT_BUTTON, self.OnClick, self.b)
      sizer = wx.BoxSizer(wx.VERTICAL)        
      sizer.Add(self.dbb, 0, wx.ALL, 5)
      box = wx.BoxSizer()
      box.Add(sizer, 0, wx.ALL, 20)
      self.b.SetInitialSize()  
      self.SetSizer(box)
      a = 1

    if eid == 2 :
      if a == 1:
        self.dbb.Destroy()
        self.b.Destroy()
        a = 0  
      if a == 2:
        self.dbb1.Destroy()
        self.b1.Destroy()
        a = 0 
      if a == 5:
        self.dbbo.Destroy()
        self.bo.Destroy()
        self.x11.Destroy()
        self.x22.Destroy()
        self.x33.Destroy()
        self.x44.Destroy()
        self.x55.Destroy()
        a = 0 
      if a == 6:
        self.dbboo.Destroy()
        self.boo.Destroy()
        self.x1.Destroy()
        self.x2.Destroy()
        self.x3.Destroy()
        self.x4.Destroy()
        self.x5.Destroy()
        a = 0
      if a == 11:
        self.dbbp1.Destroy()
        self.dbbp2.Destroy()
        self.bp.Destroy()
        a = 0    
      self.dbb1 = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'建檔根目錄:', changeCallback = self.dbbCallback1)
      self.b1 = wx.Button(self, 40, u'開始建立xlsx檔', (355,70), style=wx.NO_BORDER)
      self.b1.SetToolTipString("xlsx.")
      self.Bind(wx.EVT_BUTTON, self.OnClick1, self.b1)
      sizer = wx.BoxSizer(wx.VERTICAL)        
      sizer.Add(self.dbb1, 0, wx.ALL, 5)
      box = wx.BoxSizer()
      box.Add(sizer, 0, wx.ALL, 20)
      self.b1.SetInitialSize()  
      self.SetSizer(box)
      a = 2

    if eid == 4 :
      self.Close()

    if eid == 5 :
      if a == 1:
        self.dbb.Destroy()
        self.b.Destroy()
        a = 0 
      if a == 2:
        self.dbb1.Destroy()
        self.b1.Destroy()
        a = 0
      if a == 5:
        self.dbbo.Destroy()
        self.bo.Destroy()
        self.x11.Destroy()
        self.x22.Destroy()
        self.x33.Destroy()
        self.x44.Destroy()
        self.x55.Destroy()
        a = 0   
      if a == 6:
        self.dbboo.Destroy()
        self.boo.Destroy()
        self.x1.Destroy()
        self.x2.Destroy()
        self.x3.Destroy()
        self.x4.Destroy()
        self.x5.Destroy()
        a = 0 
      if a == 11:
        self.dbbp1.Destroy()
        self.dbbp2.Destroy()
        self.bp.Destroy()
        a = 0 
      self.dbbo = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'圖片根目錄:', changeCallback = self.dbbCallbackocr)
      self.bo = wx.Button(self, 40, u'RUN', (500,28), style=wx.NO_BORDER)
      self.bo.SetToolTipString("Run OCR\n") 
      self.Bind(wx.EVT_BUTTON, self.OnClicko, self.bo)
      sizer = wx.BoxSizer(wx.VERTICAL)        
      sizer.Add(self.dbbo, 0, wx.ALL, 5)
      box = wx.BoxSizer()
      box.Add(sizer, 0, wx.ALL, 20)
      self.bo.SetInitialSize()  
      self.SetSizer(box)
      self.x11 = wx.CheckBox(self, 11, u"日期", (35, 100), (150, 20))  
      self.x22 = wx.CheckBox(self, 22, u"時間", (35, 120), (150, 20))  
      self.x33 = wx.CheckBox(self, 33, u"車牌", (35, 140), (150, 20))
      self.x44 = wx.CheckBox(self, 44, u"速限", (35, 160), (150, 20))
      self.x55 = wx.CheckBox(self, 55, u"速度", (35, 180), (150, 20))
      self.x11.SetValue(True)
      self.x22.SetValue(True)
      self.x33.SetValue(True)
      self.x44.SetValue(True)
      self.x55.SetValue(True)
      self.Bind(wx.EVT_CHECKBOX, self.OnClicko2, self.x11)
      self.Bind(wx.EVT_CHECKBOX, self.OnClicko2, self.x22)
      self.Bind(wx.EVT_CHECKBOX, self.OnClicko2, self.x33)
      self.Bind(wx.EVT_CHECKBOX, self.OnClicko2, self.x44)
      self.Bind(wx.EVT_CHECKBOX, self.OnClicko2, self.x55)
      a = 5
    if eid == 6 :
      if a == 1:
        self.dbb.Destroy()
        self.b.Destroy()
        a = 0 
      if a == 2:
        self.dbb1.Destroy()
        self.b1.Destroy()
        a = 0     
      if a == 5:
        self.dbbo.Destroy()
        self.bo.Destroy()
        self.x11.Destroy()
        self.x22.Destroy()
        self.x33.Destroy()
        self.x44.Destroy()
        self.x55.Destroy()
        a = 0 
      if a == 6:
        self.dbboo.Destroy()
        self.boo.Destroy()
        self.x1.Destroy()
        self.x2.Destroy()
        self.x3.Destroy()
        self.x4.Destroy()
        self.x5.Destroy()
        a = 0 
      if a == 11:
        self.dbbp1.Destroy()
        self.dbbp2.Destroy()
        self.bp.Destroy()
        a = 0 
      self.dbboo = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'圖片根目錄:', changeCallback = self.dbbCallbackoo)
      self.boo = wx.Button(self, 40, u'RUN', (500,28), style=wx.NO_BORDER)
      self.boo.SetToolTipString("Run OCR\n")
      self.Bind(wx.EVT_BUTTON, self.OnClickoo, self.boo)
      sizer = wx.BoxSizer(wx.VERTICAL)        
      sizer.Add(self.dbboo, 0, wx.ALL, 5)
      box = wx.BoxSizer()
      box.Add(sizer, 0, wx.ALL, 20)
      self.boo.SetInitialSize()  
      self.SetSizer(box)

      self.x1 = wx.CheckBox(self, 1, u"日期", (35, 100), (150, 20))  
      self.x2 = wx.CheckBox(self, 2, u"時間", (35, 120), (150, 20))  
      self.x3 = wx.CheckBox(self, 3, u"車牌", (35, 140), (150, 20))
      self.x4 = wx.CheckBox(self, 4, u"速限", (35, 160), (150, 20))
      self.x5 = wx.CheckBox(self, 5, u"速度", (35, 180), (150, 20))
      self.x1.SetValue(True)
      self.x2.SetValue(True)
      self.x3.SetValue(True)
      self.x4.SetValue(True)
      self.x5.SetValue(True)
      self.Bind(wx.EVT_CHECKBOX, self.OnClick2oo, self.x1)
      self.Bind(wx.EVT_CHECKBOX, self.OnClick2oo, self.x2)
      self.Bind(wx.EVT_CHECKBOX, self.OnClick2oo, self.x3)
      self.Bind(wx.EVT_CHECKBOX, self.OnClick2oo, self.x4)
      self.Bind(wx.EVT_CHECKBOX, self.OnClick2oo, self.x5)
      a = 6
    if eid == 10:
      subprocess.Popen('test.bat', shell=True)
      sys.exit(0)
    if eid == 11:
      if a == 1:
        self.dbb.Destroy()
        self.b.Destroy()
        a = 0  
      if a == 2:
        self.dbb1.Destroy()
        self.b1.Destroy()
        a = 0 
      if a == 5:
        self.dbbo.Destroy()
        self.bo.Destroy()
        self.x11.Destroy()
        self.x22.Destroy()
        self.x33.Destroy()
        self.x44.Destroy()
        self.x55.Destroy()
        a = 0 
      if a == 6:
        self.dbboo.Destroy()
        self.boo.Destroy()
        self.x1.Destroy()
        self.x2.Destroy()
        self.x3.Destroy()
        self.x4.Destroy()
        self.x5.Destroy()
        a = 0 
      if a == 11:
        self.dbbp1.Destroy()
        self.dbbp2.Destroy()
        self.bp.Destroy()
        a = 0 
      self.dbbp1 = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'圖片檔案位置:', changeCallback = self.dbbCallbackp)
      self.dbbp2 = filebrowse.FileBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'DBF檔案位置:', changeCallback = self.dbbCallbackp)
      self.bp = wx.Button(self, 40, u'RUN', (500,28), style=wx.NO_BORDER)
      self.Bind(wx.EVT_BUTTON, self.OnClickp, self.bp)
      sizer = wx.BoxSizer(wx.VERTICAL)        
      sizer.Add(self.dbbp1, 0, wx.ALL, 5)
      sizer.Add(self.dbbp2, 0, wx.ALL, 5)
      box = wx.BoxSizer()
      box.Add(sizer, 0, wx.ALL, 20)
      self.bp.SetInitialSize()  
      self.SetSizer(box)
      a = 11
###########################################################################

###################################xlsx建檔####################################
  def dbbCallback1(self, event):
      Dir = self.dbb1.GetValue()
      self.path = Dir
      event.Skip()
  def OnClick1(self, event):         
      self.listpathx(self.path)
  def listpathx(e,dirpath):
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
      workbook = xlsxwriter.Workbook(dirpath+u'\\'+xlsxname)
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
###########################################################################
###################################OCR 忽略有 txt####################################
  def dbbCallbackocr(self, event):
      Dir = self.dbbo.GetValue()
      self.path = Dir
      event.Skip()

  def OnClicko2(self, event):
      global da_v,ti_v,li_v,up_v,do_v   
      eid = event.GetId()
      if eid == 11 and da_v == False:
          da_v = True
      elif eid == 22 and ti_v == False:
          ti_v = True
      elif eid == 33 and li_v == False:
          li_v = True
      elif eid == 44 and up_v == False:
          up_v = True
      elif eid == 55 and do_v == False:
          do_v = True
      elif eid == 11 and da_v == True:
          da_v = False
      elif eid == 22 and ti_v == True:
          ti_v = False
      elif eid == 33 and li_v == True:
          li_v = False
      elif eid == 44 and up_v == True:
          up_v = False
      elif eid == 55 and do_v == True:
          do_v = False


  def OnClicko(self, event): 
      self.listpathocr(self.path)
  def listpathocr(e,dirpath):
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
                                 # if strsu[len(strsu)-2:len(strsu)] == "F1" :
                                  os.mkdir(pathr)
                                  count = count+1
                                  mainomli(strsu,strf,count)
                              if u'快速道路' in strf : 
                                 # if strsu[len(strsu)-2:len(strsu)] == "F1" :
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
                                  #if strsu[len(strsu)-2:len(strsu)] == "F1" :
                                  mainomli(strsu,strf,count)
                              if u'快速道路' in strf : 
                                 # if  strsu[len(strsu)-2:len(strsu)] == "F1" :
                                  mainomli2(strsu,strf,count)
                              if  u'紅燈' not in strf and  u'標線' not in strf and u'號誌' not in strf : 
                                  mainom(strsu,strf,count)

          tEnd = time.time()
          print "Produce  "+str(count)+"  files  spend  "+str((tEnd - tStart)//1)+"  second" 

###########################################################################
###################################  OCR  ####################################
  def dbbCallbackoo(self, event):
      Dir = self.dbboo.GetValue()
      self.path = Dir
      event.Skip()
  def OnClick2oo(self, event):
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

  def OnClickoo(self, event):         
      self.listpathoo(self.path)
  def listpathoo(e,dirpath):
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
                          # if strsu[len(strsu)-2:len(strsu)] == "F1" :
                          os.mkdir(pathr)
                          count = count+1
                          mainomli(strsu,strf,count)
                      if u'快速道路' in strf : 
                          # if strsu[len(strsu)-2:len(strsu)] == "F1" :
                          os.mkdir(pathr)
                          count = count+1
                          mainomli2(strsu,strf,count)
                      if  u'紅燈' not in strf and  u'標線' not in strf and u'號誌' not in strf : 
                          count = count+1
                          os.mkdir(pathr)
                          mainom(strsu,strf,count) 
                  else:
                      if u'紅燈' in strf or u'不依標誌標線' in strf or u'未依標線之指示' in strf : 
                          # if strsu[len(strsu)-2:len(strsu)] == "F1" :
                          count = count+1
                          mainomli(strsu,strf,count)
                      if u'快速道路' in strf : 
                          # if  strsu[len(strsu)-2:len(strsu)] == "F1" :
                          count = count+1
                          mainomli2(strsu,strf,count)
                      if  u'紅燈' not in strf and  u'標線' not in strf and u'號誌' not in strf : 
                          count = count+1
                          mainom(strsu,strf,count)
      tEnd = time.time()
      print "Produce  "+str(count)+"  files  spend  "+str((tEnd - tStart)//1)+"  second" 
###########################################################################
###################################格式更正####################################
  def dbbCallback(self, e):
    Dir = self.dbb.GetValue()
    self.path = Dir
    e.Skip()
  def OnClick(self, e):    
    self.listpath(self.path)   
  def listpath(e,dirpath): 
    num1 = True   
    tStart = time.time() 
    c = 0                  
    pa = "%^$&&$^"  
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
    for root, dirs, files in os.walk(dirpath):               
        for f1 in files:
            strph =  os.path.join(root , f1)
            finame = strph[strph.rfind('\\'):len(strph)].replace('.jpg','')
            if ".jpg" in strph :
                if len(finame) < 12 :
                    wx.MessageBox(u"發現異常 請更正 "+strph)
                    sys.exit(0)
                if  u'紅燈' not in strph and u'標線' not in strph and u'標誌' not in strph and u'快速道路' not in strph :
                    strph1 = strph[0:strph.rfind('\\')]
                    patnum = strph1[strph1.rfind(u'\\'):len(strph1)]
                    patnum1 = strph1.replace(patnum, u'&')
                    patnum2 = patnum1[patnum1.rfind(u'\\')+1:patnum1.rfind(u'&')-11]
                    if pa != patnum2:
                        if patnum2.encode('utf-8') not in piece1 :
                            pa = patnum2 
                            PATH.append(patnum2)
                            PATH1.append(strph)
                            print patnum2
                            print strph   
                if  u'紅燈' in strph or u'標線' in strph or u'標誌' in strph or u'快速道路' in strph : 
                    strph2 = strph[0:strph.rfind('\\')]
                    strph2 = strph2[0:strph2.rfind('\\')]
                    patnum = strph2[strph2.rfind(u'\\'):len(strph2)]
                    patnum1 = strph2.replace(patnum, u'&')
                    patnum2 = patnum1[patnum1.rfind(u'\\')+1:patnum1.rfind(u'&')-11]
                    if pa != patnum2:
                        if patnum2.encode('utf-8') not in piece1 :
                            pa = patnum2 
                            PATH.append(patnum2)
                            PATH1.append(strph)
                            print patnum2
                            print strph  
    if len(PATH)>0 and len(PATH1)>0 :
        contPATH = len(PATH)
        apppath = Apppath()
        apppath.MainLoop()
    tEnd = time.time()
    print "Success"
    print " spend  "+str((tEnd - tStart)//1)+"  second" 
      
###########################################################################
############################### DBF--->PDF ############################################
  def dbbCallbackp(self, event):
    Dir = self.dbbp1.GetValue()
    Dir1 = self.dbbp2.GetValue()
    self.path = Dir
    self.path1 = Dir1
    event.Skip()
  def OnClickp(self, event):         
    self.listpathpdf(self.path,self.path1)
  
  def listpathpdf(e,dirpath,dirpath1):
    tStart = time.time()
    global story
    cont = 0
    cont1 = 0
    Story=[]
    fig = False
    ost = ""
    pdfpath1 = ""
    ddrddd = dirpath1[0:dirpath1.rfind('\\')+1]
    if ".dbf" not in dirpath1 and ".DBF" not in dirpath1:
        wx.MessageBox(u"不存在DBF檔 請重新選取",u"提示訊息")
    else :
        print 'Wait for PDF......'
        db = dbf.Dbf(dirpath1) 
        for record in db:
            # print record['Plt_no'], record['Vil_dt'], record['Vil_time'],record['Bookno'],record['Vil_addr'],record['Rule_1'],record['Truth_1'],record['Rule_2'],record['Truth_2'],record['color'],record['A_owner1']
            filename2 = record['Plt_no'].decode('big5') + "." +  record['Vil_dt'] + "." + record['Vil_time'] #檔名
            piece2.append( filename2 )            #檔名
            piece2.append( record['Plt_no'] )    #車牌       
            piece2.append( record['Vil_dt'] )    #日期
            piece2.append( record['Vil_time'] )  #時間
            piece2.append( record['Bookno'] )    #冊頁號  
            piece2.append( record['Vil_addr'] )  #違規地點 
            piece2.append( record['Rule_1'] )    #法條1 
            piece2.append( record['Truth_1'] )   #法條1事實 
            piece2.append( record['Rule_2'] )    #法條2 
            piece2.append( record['Truth_2'] )   #法條2事實 
            piece2.append( record['color'] )     #車顏色 
            piece2.append( record['A_owner1'] )  #車廠牌
            record.store()
        print 'Wait for PDF......'
        for root, dirs, files in os.walk(dirpath):
            for f in files:
                if ".jpg" in os.path.join(root , f) :
                    strsu =  os.path.join(f).replace('-1.jpg','').replace('_1.jpg','')
                    if strsu in piece2 :
                        cc = os.path.join(root , f)
                        cont = cont +1
                        photo_path.append(cc) 
        # print photo_path
        if u'-1.jpg' in photo_path[0] :
            ost = '-1'
        if u'_1.jpg' in photo_path[0] :
            ost = '_1'

        for record in db:                
            filename = record['Plt_no'] + "." +  record['Vil_dt'] + "." + record['Vil_time']+ost#檔名
            piece1.append( filename )            #檔名
            piece1.append( record['Plt_no'] )    #車牌       
            piece1.append( record['Vil_dt'] )    #日期
            piece1.append( record['Vil_time'] )  #時間
            piece1.append( record['Bookno'] )    #冊頁號  
            piece1.append( record['Vil_addr'] )  #違規地點 
            piece1.append( record['Rule_1'] )    #法條1 
            piece1.append( record['Truth_1'] )   #法條1事實 
            piece1.append( record['Rule_2'] )    #法條2 
            piece1.append( record['Truth_2'] )   #法條2事實 
            piece1.append( record['color'] )     #車顏色 
            piece1.append( record['A_owner1'] )  #車廠牌  
            record.store()
        
        x = 0
        wxc = 0
        wxc2 = 0
        v22 = ''
        v1 = ''
        c = False
        e = len(photo_path)
        wxwx = len(piece1)//12 
        for x in range(wxwx):
            for t in range(e):
                phtotpath = photo_path[t]
                phtotpath = phtotpath[phtotpath.rfind(u'\\')+1:len(phtotpath)].replace('.jpg','')
                e1 = piece1[x*12+1]
                if e1[0:len(e1)-5] == phtotpath[0:len(e1)-5] and piece1[x*12+11] != '' and piece1[x*12+12] != '':
                    fig = True
                    break
                # if  e1[0:len(e1)-5] == phtotpath[0:len(e1)-5] and piece1[x*12+11] == '' and piece1[x*12+12] == '':
                #     pho = photo_path[t]
                #     print '123'
                #     print pho.encode('utf-8')
                #     break
            if fig == False and piece1[x*12+11] != '' and piece1[x*12+12] != '':
                wxc2 = 1
                print e1
                print piece1[x*12+5+1]
            if piece1[x*12+11] == '' and piece1[x*12+12] == '':
                pdfmetrics.registerFont(TTFont('song', 'simsun.ttf'))
                fonts.addMapping('song', 0, 0, 'song')
                fonts.addMapping('song', 0, 1, 'song')
                stylesheet=getSampleStyleSheet()
                normalStyle = copy.deepcopy(stylesheet['Normal'])
                normalStyle.fontName ='song'
                normalStyle.size = '13'
                # im2 = Image(pho,400,300)
                # story.append(im2)
                story2.append(Paragraph(u'<font size=15 color=red>車牌: '+piece2[x*12+1+1].decode('big5')+'</font><font size=13 color=white>-</font>'+u'<font size=13>冊頁號: </font><font size=13 color=blue>'+piece2[x*12+1+4]+'</font><font size=13 color=white>-----</font>'+u'<font size=13 color=blue>-1</font>', normalStyle))
                story2.append(Paragraph('----------------------------------------------------------------------------------------------------------------',normalStyle))
                story2.append(Paragraph(u'<font size=13>日期: </font><font size=13 color=blue>'+piece2[x*12+1+2]+'</font><font size=13 color=white>--</font>'+u'   <font size=13>時間: </font><font size=13 color=blue>'+piece2[x*12+1+3]+'</font><font size=13 color=white>--</font>'+u'   <font size=13>違規地點: </font><font size=13 color=blue>'+piece2[x*12+1+5].decode('big5')+'</font>',normalStyle))
                story2.append(Paragraph(u'<font size=13 color=blue>錯誤訊息:  交換不到車種、顏色，皆為空</font>', normalStyle))
                story2.append(Paragraph('----------------------------------------------------------------------------------------------------------------',normalStyle))
                wxc+=1
            fig =  False 
        if story2 != []:
            bbt = piece2[x*12+1+4]
            doc2 = SimpleDocTemplate(dirpath1[0:dirpath1.rfind('.')]+'_exception.pdf',rightMargin=1,leftMargin=1,topMargin=1,bottomMargin=1)
            print 'Wait for PDF_exception......'         
            doc2.build(story2)
            print 'Mission accomplished'     
        if wxc2 == 1: 
            print '請修正錯誤後重來'              
        if wxc2 == 0 :
            for x in range(wxwx):
                for t in range(e):
                    # print piece1[x*12+1]
                    # print photo_path[t].encode('utf-8')
                    phtotpath = photo_path[t]
                    phtotpath = phtotpath[phtotpath.rfind(u'\\')+1:len(phtotpath)].replace('.jpg','')
                    
                    e1 = piece1[x*12+1]
                    if e1[0:len(e1)-5] == phtotpath[0:len(e1)-5] and piece1[x*12+11] != '' and piece1[x*12+12] != '':

                        # if x%2 == 0:
                        #     v22 = v1
                        fig = True
                        cont1 = cont1+1
                        pathr = photo_path[t]
                        pathrr = photo_path[t].replace('1.jpg','2.jpg')
                        isexists = os.path.exists(pathr)
                        isexistss2 = os.path.exists(pathrr)
                        pdfpath1 = pathr[0:pathr.rfind(u'\\')]
                        pdfpath1 = pdfpath1[0:pdfpath1.rfind(u'\\')+1]
                        if  not isexists:
                            print pathr
                        if  isexists:
                            Pnum = piece2[x*12+1+4]
                            tryr = Pnum[0:9]
                            photo = pathr
                            photo3 = pathrr
                            pdfmetrics.registerFont(TTFont('song', 'simsun.ttf'))
                            fonts.addMapping('song', 0, 0, 'song')
                            fonts.addMapping('song', 0, 1, 'song')
                            #-----------------------------------------------------
                            db1 = dbf.Dbf('DBF\\COLOR_CODE.DBF') 

                            for record in db1:
                                co = piece2[x*12+1+10]
                                if len(co) > 0:
                                    if record['Color_id'] == co[0]:
                                        color = record['Color']
                                else:
                                    color = ""
                            for record in db1:
                                co = piece2[x*12+1+10]
                                if len(co) == 2:
                                    if record['Color_id'] == co[1]:
                                        color = color+record['Color']
                            #-----------------------------------------------------
                            stylesheet=getSampleStyleSheet()
                            normalStyle = copy.deepcopy(stylesheet['Normal'])
                            normalStyle.fontName ='song'
                            normalStyle.size = '13'
                            isexistsss = os.path.exists(photo3)
                            im = Image(photo,400,300)
                            story.append(im)
                            story.append(Paragraph(u'<font size=15 color=red>車牌: '+piece2[x*12+1+1]+'</font><font size=13 color=white>-</font>'+u'<font size=13>廠牌: </font><font size=13 color=blue>'+piece2[x*12+1+11].decode('big5')+'</font><font size=13 color=white>-</font>'+u'   <font size=13>顏色: </font><font size=13 color=blue>'+color.decode('big5')+'</font><font size=13 color=white>-</font>'+u'<font size=13>冊頁號: </font><font size=13 color=blue>'+piece2[x*12+1+4]+'</font><font size=13 color=white>-----</font>'+u'<font size=13 color=blue>-1</font>', normalStyle))
                            story.append(Paragraph('----------------------------------------------------------------------------------------------------------------',normalStyle))
                            story.append(Paragraph(u'<font size=13>日期: </font><font size=13 color=blue>'+piece2[x*12+1+2]+'</font><font size=13 color=white>--</font>'+u'   <font size=13>時間: </font><font size=13 color=blue>'+piece2[x*12+1+3]+'</font><font size=13 color=white>--</font>'+u'   <font size=13>違規地點: </font><font size=13 color=blue>'+piece2[x*12+1+5].decode('big5')+'</font>',normalStyle))
                            story.append(Paragraph(u'<font size=13>違規法條1: </font><font size=13 color=blue>'+piece2[x*12+1+6]+'</font>',normalStyle))
                            story.append(Paragraph(u'<font size=13>違規事實1: </font><font size=13 color=blue>'+piece2[x*12+1+7].decode('big5').replace(' ','')+'</font>',normalStyle))
                            if len(piece2[x*12+1+8])>1:
                                story.append(Paragraph(u'<font size=13>違規法條2: </font><font size=13 color=blue>'+piece2[x*12+1+8]+u'</font><font size=13 color=white>----</font><font size=13>違規事實2: </font><font size=13 color=blue>'+piece2[x*12+1+9].decode('big5')+'</font>',normalStyle))
                            story.append(Paragraph('----------------------------------------------------------------------------------------------------------------',normalStyle))
                            if isexistsss:
                                im2 = Image(photo3,400,300)
                                story.append(im2)
                                story.append(Paragraph(u'<font size=15 color=red>車牌: '+piece2[x*12+1+1]+'</font><font size=13 color=white>-</font>'+u'<font size=13>廠牌: </font><font size=13 color=blue>'+piece2[x*12+1+11].decode('big5')+'</font><font size=13 color=white>-</font>'+u'   <font size=13>顏色: </font><font size=13 color=blue>'+color.decode('big5')+'</font><font size=13 color=white>-</font>'+u'<font size=13>冊頁號: </font><font size=13 color=blue>'+piece2[x*12+1+4]+'</font><font size=13 color=white>-----</font>'+u'<font size=13 color=blue>-2</font>', normalStyle))
                                story.append(Paragraph('----------------------------------------------------------------------------------------------------------------',normalStyle))
                                story.append(Paragraph(u'<font size=13>日期: </font><font size=13 color=blue>'+piece2[x*12+1+2]+'</font><font size=13 color=white>--</font>'+u'   <font size=13>時間: </font><font size=13 color=blue>'+piece2[x*12+1+3]+'</font><font size=13 color=white>--</font>'+u'   <font size=13>違規地點: </font><font size=13 color=blue>'+piece2[x*12+1+5].decode('big5')+'</font>',normalStyle))
                                story.append(Paragraph(u'<font size=13>違規法條1: </font><font size=13 color=blue>'+piece2[x*12+1+6]+'</font>',normalStyle))
                                story.append(Paragraph(u'<font size=13>違規事實1: </font><font size=13 color=blue>'+piece2[x*12+1+7].decode('big5').replace(' ','')+'</font>',normalStyle))
                                if len(piece2[x*12+1+8])>1:
                                    story.append(Paragraph(u'<font size=13>違規法條2: </font><font size=13 color=blue>'+piece2[x*12+1+8]+u'</font><font size=13 color=white>----</font><font size=13>違規事實2: </font><font size=13 color=blue>'+piece2[x*12+1+9].decode('big5')+'</font>',normalStyle))
                                story.append(Paragraph('----------------------------------------------------------------------------------------------------------------',normalStyle))

      
                            print 'Progress  '+str(cont1)+'/'+str(cont-wxc)+'/'+str(wxwx)+'....exception:'+str(wxc)
                        break
                if x < wxwx-1:
                    g1 = piece2[x*12+1+4]
                    g1 = g1[0:len(g1)-3]
                    g2 = piece2[x*12+1+4+12]
                    g2 = g2[0:len(g2)-3]
                if g1 != g2 or x == wxwx-1:
                    bbt = piece2[x*12+1+4]
                    doc = SimpleDocTemplate(ddrddd+bbt[0:len(bbt)-3]+piece2[x*12+1+5].decode('big5').replace('?','')+'.pdf',rightMargin=1,leftMargin=1,topMargin=1,bottomMargin=1)
                    print 'Wait for PDF......'         
                    doc.build(story)
                    story=[]
                    print 'Mission accomplished'               
    tEnd = time.time()
    print "Spend  "+str((tEnd - tStart)//1)+"  second" 
###########################################################################
def scale_bitmap(bitmap, width, height):
    image = wx.ImageFromBitmap(bitmap)
    image = image.Scale(width, height, wx.IMAGE_QUALITY_HIGH)
    result = wx.BitmapFromImage(image)
    return result
class TestPanelpath(wx.Panel):
    def __init__(self, parent, path):
        super(TestPanelpath, self).__init__(parent, -1)
        bitmap = wx.Bitmap(path)
        bitmap = scale_bitmap(bitmap, 1024, 800)
        control = wx.StaticBitmap(self, -1, bitmap)
        control.SetPosition((10, 60))

        self.txtx = wx.StaticText(self,label=u"請輸入違規地點: ")  
        self.text = wx.TextCtrl(self,value = PATH[0] , size=(200, 20)) 
        b = wx.Button(self, 40, u'加入', (250,28), style=wx.NO_BORDER)
        self.Bind(wx.EVT_BUTTON, self.OnClick, b)
        sizer = wx.BoxSizer(wx.VERTICAL) 
        sizer.Add(self.txtx, 0, wx.ALL, 5,100)
        sizer.Add(self.text, 0, wx.ALL, 0,100)
        box = wx.BoxSizer()
        box.Add(sizer, 0, wx.ALL, 10)
        self.SetSizer(box)

    def OnClick(self, event): 
        tyyu = self.text.GetValue()
        # print PATH[0].encode('utf-8')
        # print tyyu.encode('utf-8')
        f = file(ocrfilepath, 'a+') 
        f.write(";"+PATH[0].encode('utf-8')+";"+tyyu.encode('utf-8')+"\n") # write text to file
        f.close()       
        wx.MessageBox(u"輸入成功 確認後關閉程式 請重新開啟")
        sys.exit(0)

class Framepath( wx.Frame ):
    def __init__( self, parent ):
        wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = u"寫入新資料", pos = wx.DefaultPosition, size = wx.Size(1100,1100))
        panel = TestPanelpath(self, PATH1[0])

class Apppath(wx.App):
    def OnInit(self):
        self.frame = Framepath(parent=None)
        icon = wx.EmptyIcon()
        icon.CopyFromBitmap(wx.Bitmap(icopath, wx.BITMAP_TYPE_ANY))
        self.frame.SetIcon(icon)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True
###########################################################################
###########################################################################
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
def omocrwr(path1):
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
def main(): 
  ex = wx.App()      #生成一个应用程序 
  Example(None, id=-1, title="main")  #调用我们的类 
  ex.MainLoop()#消息循环 
  
if __name__ == "__main__": 
  main() 
