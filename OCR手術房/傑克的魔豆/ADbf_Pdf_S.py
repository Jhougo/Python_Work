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
from PIL import Image
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
photo_path = []
story = []
Color = ""
# ------------------------------------Generate TXT-----------------------------------------------
class TestPanel(wx.Panel):
    def listpath(e,dirpath,dirpath1):
        doc = SimpleDocTemplate("form_letter1.pdf",rightMargin=1,leftMargin=1,topMargin=1,bottomMargin=1)
        tStart = time.time()
        global story
        cont = 0
        cont1 = 0
        Story=[]
        ost = ""
        pdfpath1 = ""
        if ".dbf" not in dirpath1 and ".DBF" not in dirpath1:
            wx.MessageBox(u"不存在DBF檔 請重新選取",u"提示訊息")
        else :
            db = dbf.Dbf(dirpath1) 
            for record in db:
                # print record['Plt_no'], record['Vil_dt'], record['Vil_time'],record['Bookno'],record['Vil_addr'],record['Rule_1'],record['Truth_1'],record['Rule_2'],record['Truth_2'],record['color'],record['A_owner1']
                filename2 = record['Plt_no'] + "." +  record['Vil_dt'] + "." + record['Vil_time'] #檔名
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
            print '正在轉換圖檔大小......'
            for root, dirs, files in os.walk(dirpath):
                for f in files:
                    if ".jpg" in os.path.join(root , f) :
                        strsu =  os.path.join(f).replace('-1.jpg','').replace('_1.jpg','')
                        if strsu in piece2 :
                            cc = os.path.join(root , f)
                            cont = cont +1
                            photo_path.append(cc) 
            if u'-1.jpg' in photo_path[0] :
                ost = u'-1'
            if u'_1.jpg' in photo_path[0] :
                ost = u'_1'
            e = len(photo_path)
            for t1 in range(e):
                pathr = photo_path[t1]
                pathrr = photo_path[t1].replace('1.jpg','2.jpg')
                isexists = os.path.exists(pathr)
                isexistss2 = os.path.exists(pathrr)

                img = PIL.Image.open(pathr)
                imgSize = img.size 
                # print imgSize
                maxSize = max(imgSize) 
                # print maxSize
                # minSize = min(imgSize) 
                # print minSize
                if maxSize > 900 :
                    if  isexists:
                        resize_img(pathr,pathr, 900) 
                    if  isexistss2:
                        resize_img(pathrr,pathrr, 900) 

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
            print 'Wait for PDF......'
            x = 0
            c = False

            wxwx = (len(piece1))//12 
            for x in range(wxwx):
                for t in range(e):
                    # print piece1[x*12+1]
                    # print photo_path[t].encode('utf-8')
                    phtotpath = photo_path[t]
                    phtotpath = phtotpath[phtotpath.rfind(u'\\')+1:len(phtotpath)].replace('.jpg','')
                    e1 = piece1[x*12+1]
                    if e1[0:len(e1)-5] == phtotpath[0:len(e1)-5]:

                        cont1 = cont1+1
                        pathr2 = photo_path[t]
                        pathrr2 = photo_path[t].replace('1.jpg','2.jpg')
                        pdfpath1 = pathr2[0:pathr2.rfind(u'\\')]
                        pdfpath1 = pdfpath1[0:pdfpath1.rfind(u'\\')+1]
                        # cc2isexists = os.path.exists(pathrr2)
                        # if cc2isexists:
                        #     cont1 = cont1+1
                        Pnum = piece2[x*12+1+4]
                        tryr = Pnum[0:9]
                        photo = pathr2
                        photo3 = pathrr2
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

  
                        print 'Progress  '+str(cont1)+'/'+str(cont)+'/'+str(wxwx)
                        break
                            
            doc = SimpleDocTemplate(pdfpath1+'repor.pdf',rightMargin=1,leftMargin=1,topMargin=1,bottomMargin=1)
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
