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
import xlsxwriter
import Image
import sys
import PIL 
import time
import copy
from dbfpy import dbf 
import wx.lib.filebrowsebutton as filebrowse
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




Dir = ""
Dir1 = ""
piece1  = ['start']
piece2  = ['start']
story = []
Color = ""
# ------------------------------------Generate TXT-----------------------------------------------
class TestPanel(wx.Panel):
    def listpath(e,dirpath,dirpath1):
        tStart = time.time()
        if ".dbf" not in dirpath1 and ".DBF" not in dirpath1:
            wx.MessageBox(u"不存在DBF檔 請重新選取",u"提示訊息")
        else :
            db = dbf.Dbf(dirpath1) 
            for record in db:
                # print record['Plt_no'], record['Vil_dt'], record['Vil_time'],record['Bookno'],record['Vil_addr'],record['Rule_1'],record['Truth_1'],record['Rule_2'],record['Truth_2'],record['color'],record['A_owner1']
                filename = record['Plt_no'] + "." +  record['Vil_dt'] + "." + record['Vil_time'] +u'-1'#檔名
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

            print ('Mission accomplished')                  
        tEnd = time.time()
        print ("Spend  "+str((tEnd - tStart)//1)+"  second" )
                        
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent)
        self.log = log
        self.dbb = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'檔案位置:', changeCallback = self.dbbCallback)
        b = wx.Button(self, 40, u'RUN', (500,28), style=wx.NO_BORDER)
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
        self.listpath(self.path,self.path1)

        
class Frame ( wx.Frame ):
    def __init__( self, parent ):
        wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = u"碰出命中率", pos = wx.DefaultPosition, size = wx.Size(700,300))
        panel = TestPanel(self, -1)

class App(wx.App):
    def OnInit(self):
        self.frame = Frame(parent=None)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True
if __name__ == '__main__':
    app = App()
    app.MainLoop()
