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
        doc = SimpleDocTemplate("form_letter1.pdf",rightMargin=1,leftMargin=1,topMargin=1,bottomMargin=1)
        tStart = time.time()
        global story
        count = 0
        cont = 0
        cont1 = 0
        Story=[]
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
                filename2 = record['Plt_no'] + "." +  record['Vil_dt'] + "." + record['Vil_time'] +u'_1'#檔名
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
                    if ".jpg" in os.path.join(root , f) and "_sm.jpg" not in os.path.join(root , f):
                        strsu =  os.path.join(f).replace('.jpg','')
                        if strsu in piece1 or strsu in piece2:
                            cont = cont +1
            
            for root, dirs, files in os.walk(dirpath):
                for f in files:
                    if ".jpg" in os.path.join(root , f) and "_sm.jpg" not in os.path.join(root , f):
                        pathr = os.path.join(root , f)
                        pathr2 = os.path.join(root , f).replace('.jpg','')+'_sm.jpg'
                        strsu =  os.path.join(f).replace('.jpg','')
                        isexists = os.path.exists(pathr2)
                        pdfpath1 = pathr[0:pathr.rfind(u'\\')]
                        if strsu in piece1 :
                            cont1 = cont1 +1
                            x =  piece1.index(strsu) 
                            if not isexists:
                                resize_img(pathr,pathr2, 600) 
                            if  isexists:
                                photo = pathr2
                                pdfmetrics.registerFont(TTFont('song', 'simsun.ttf'))
                                fonts.addMapping('song', 0, 0, 'song')
                                fonts.addMapping('song', 0, 1, 'song')
                                #-----------------------------------------------------
                                db1 = dbf.Dbf('DBF\\COLOR_CODE.DBF') 

                                for record in db1:
                                    co = piece2[x+10]
                                    if len(co) > 0:
                                        if record['Color_id'] == co[0]:
                                            color = record['Color']
                                    else:
                                        color = ""
                                for record in db1:
                                    co = piece2[x+10]
                                    if len(co) == 2:
                                        if record['Color_id'] == co[1]:
                                            color = color+record['Color']
                                #-----------------------------------------------------
                                stylesheet=getSampleStyleSheet()
                                normalStyle = copy.deepcopy(stylesheet['Normal'])
                                normalStyle.fontName ='song'
                                normalStyle.size = '13'
                                
                                im = Image(photo,400,300)
                                story.append(im)
                                story.append(Paragraph(u'<font size=15 color=red>車牌: '+piece2[x+1]+'</font><font size=13 color=white>-</font>'+u'<font size=13>廠牌: </font><font size=13 color=blue>'+piece2[x+11].decode('big5')+'</font><font size=13 color=white>-</font>'+u'   <font size=13>顏色: </font><font size=13 color=blue>'+color.decode('big5')+'</font><font size=13 color=white>-</font>'+u'<font size=13>冊號: </font><font size=13 color=blue>'+piece2[x+4]+'</font><font size=13 color=white>-</font>'+u'<font size=13>檔名: </font><font size=13 color=blue>'+piece2[x]+'</font>', normalStyle))
                                story.append(Paragraph('----------------------------------------------------------------------------------------------------------------',normalStyle))
                                story.append(Paragraph(u'<font size=13>日期: </font><font size=13 color=blue>'+piece2[x+2]+'</font><font size=13 color=white>--</font>'+u'   <font size=13>時間: </font><font size=13 color=blue>'+piece2[x+3]+'</font><font size=13 color=white>--</font>'+u'   <font size=13>違規地點: </font><font size=13 color=blue>'+piece2[x+5].decode('big5')+'</font>',normalStyle))
                                story.append(Paragraph(u'<font size=13>違規法條1: </font><font size=13 color=blue>'+piece2[x+6]+'</font>',normalStyle))
                                story.append(Paragraph(u'<font size=13>違規事實1: </font><font size=13 color=blue>'+piece2[x+7].decode('big5')+'</font>',normalStyle))
                                story.append(Paragraph(u'<font size=13>違規法條2: </font><font size=13 color=blue>'+piece2[x+8]+u'</font><font size=13 color=white>----</font><font size=13>違規事實2: </font><font size=13 color=blue>'+piece2[x+9].decode('big5')+'</font>',normalStyle))
                                story.append(Paragraph('----------------------------------------------------------------------------------------------------------------',normalStyle))
                                print 'Progress  '+str(cont1)+'/'+str(cont)
                                doc = SimpleDocTemplate(pdfpath1+'.pdf',rightMargin=1,leftMargin=1,topMargin=1,bottomMargin=1)

                        if strsu in piece2 :
                            cont1 = cont1 +1
                            x =  piece2.index(strsu) 
                            if not isexists:
                                resize_img(pathr,pathr2, 600)
                            if  isexists: 
                                photo = pathr2
                                pdfmetrics.registerFont(TTFont('song', 'simsun.ttf'))
                                fonts.addMapping('song', 0, 0, 'song')
                                fonts.addMapping('song', 0, 1, 'song')
                                #-----------------------------------------------------
                                db1 = dbf.Dbf('DBF\\COLOR_CODE.DBF') 

                                for record in db1:
                                    co = piece2[x+10]
                                    if len(co) > 0:
                                        if record['Color_id'] == co[0]:
                                            color = record['Color']
                                    else:
                                        color = ""
                                for record in db1:
                                    co = piece2[x+10]
                                    if len(co) == 2:
                                        if record['Color_id'] == co[1]:
                                            color = color+record['Color']


                                #-----------------------------------------------------
                                stylesheet=getSampleStyleSheet()
                                normalStyle = copy.deepcopy(stylesheet['Normal'])
                                normalStyle.fontName ='song'
                                normalStyle.size = '13'
                                
                                im = Image(photo,400,300)
                                story.append(im)
                                story.append(Paragraph(u'<font size=15 color=red>車牌: '+piece2[x+1]+'</font><font size=13 color=white>-</font>'+u'<font size=13>廠牌: </font><font size=13 color=blue>'+piece2[x+11].decode('big5')+'</font><font size=13 color=white>-</font>'+u'   <font size=13>顏色: </font><font size=13 color=blue>'+color.decode('big5')+'</font><font size=13 color=white>-</font>'+u'<font size=13>冊號: </font><font size=13 color=blue>'+piece2[x+4]+'</font><font size=13 color=white>-</font>'+u'<font size=13>檔名: </font><font size=13 color=blue>'+piece2[x]+'</font>', normalStyle))
                                story.append(Paragraph('----------------------------------------------------------------------------------------------------------------',normalStyle))
                                story.append(Paragraph(u'<font size=13>日期: </font><font size=13 color=blue>'+piece2[x+2]+'</font><font size=13 color=white>--</font>'+u'   <font size=13>時間: </font><font size=13 color=blue>'+piece2[x+3]+'</font><font size=13 color=white>--</font>'+u'   <font size=13>違規地點: </font><font size=13 color=blue>'+piece2[x+5].decode('big5')+'</font>',normalStyle))
                                story.append(Paragraph(u'<font size=13>違規法條1: </font><font size=13 color=blue>'+piece2[x+6]+'</font>',normalStyle))
                                story.append(Paragraph(u'<font size=13>違規事實1: </font><font size=13 color=blue>'+piece2[x+7].decode('big5')+'</font>',normalStyle))
                                story.append(Paragraph(u'<font size=13>違規法條2: </font><font size=13 color=blue>'+piece2[x+8]+u'</font><font size=13 color=white>----</font><font size=13>違規事實2: </font><font size=13 color=blue>'+piece2[x+9].decode('big5')+'</font>',normalStyle))
                                story.append(Paragraph('----------------------------------------------------------------------------------------------------------------',normalStyle))
                                print 'Progress  '+str(cont1)+'/'+str(cont)
                                doc = SimpleDocTemplate(pdfpath1+'.pdf',rightMargin=1,leftMargin=1,topMargin=1,bottomMargin=1)

                print 'Wait for PDF......'         
                doc.build(story)
                story=[]
            print 'Mission accomplished'                  
        tEnd = time.time()
        print "Spend  "+str((tEnd - tStart)//1)+"  second" 
                        
    def __init__(self, parent, log):
        wx.Panel.__init__(self, parent)
        self.log = log
        self.dbb = filebrowse.DirBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'圖片檔案位置:', changeCallback = self.dbbCallback)
        self.dbb1 = filebrowse.FileBrowseButton(self, -1, size=(450, -1), buttonText=u'瀏覽', labelText=u'DBF檔案位置:', changeCallback = self.dbbCallback)
        b = wx.Button(self, 40, u'RUN', (500,28), style=wx.NO_BORDER)
        self.Bind(wx.EVT_BUTTON, self.OnClick, b)
        sizer = wx.BoxSizer(wx.VERTICAL)        
        sizer.Add(self.dbb, 0, wx.ALL, 5)
        sizer.Add(self.dbb1, 0, wx.ALL, 5)
        box = wx.BoxSizer()
        box.Add(sizer, 0, wx.ALL, 20)
        b.SetInitialSize()  
        self.SetSizer(box)

    def dbbCallback(self, event):
        Dir = self.dbb.GetValue()
        Dir1 = self.dbb1.GetValue()
        self.path = Dir
        self.path1 = Dir1
        event.Skip()
    def OnClick(self, event):         
        self.listpath(self.path,self.path1)

        
class Frame ( wx.Frame ):
    def __init__( self, parent ):
        wx.Frame.__init__(self, parent, id = wx.ID_ANY, title = u"產生PDF檔", pos = wx.DefaultPosition, size = wx.Size(700,300))
        panel = TestPanel(self, -1)

class App(wx.App):
    def OnInit(self):
        self.frame = Frame(parent=None)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True
# ------------------------------------------------------------------------------------------------- 
def resize_img(img_path, out_path, new_width):
    #读取图像
    im = PIL.Image.open(img_path)
    #获得图像的宽度和高度
    width,height = im.size
    #计算高宽比
    ratio = 1.0 * height / width
    #计算新的高度
    new_height = int(new_width * ratio)
    new_size = (new_width, new_height)
    #插值缩放图像，
    out = im.resize(new_size, PIL.Image.ANTIALIAS)
    #保存图像
    out.save(out_path) 
# -------------------------------------------------------------------------------------------------
if __name__ == '__main__':
    app = App()
    app.MainLoop()
