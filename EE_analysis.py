"""
Created on Wed Feb 28 12:59:27 2018
@author: dell
"""
import os
import wx
import wx.grid as grd
import xlrd
import xlwt
from matplotlib.backends.backend_wxagg import FigureCanvasWxAgg as FigureCanvas
from matplotlib.backends.backend_wx import NavigationToolbar2Wx as NavigationToolbar
from matplotlib.font_manager import FontProperties
from matplotlib.figure import Figure
from matplotlib.widgets import Cursor
from matplotlib.widgets import MultiCursor
from pylab import *



class MyApp(wx.App):
    def OnInit(self):
        frame = MyFrame(parent=None, id=-1, title='EE analysis')
        frame.Show(True)
        return True


class MyFrame(wx.Frame):
    def __init__(self, parent, id, title):
        wx.Frame.__init__(self, parent, id, title, size=(1300, 750),
                          style=wx.DEFAULT_FRAME_STYLE | wx.RESIZE_BORDER | wx.MAXIMIZE_BOX | wx.MAXIMIZE)
                
        self.panel = wx.Panel(self, -1)
        self.grid = wx.grid.Grid(self.panel, -1)
        self.grid.CreateGrid(5, 10)
        self.figure = Figure()
        self.canvas = FigureCanvas(self.panel, -1, self.figure)
        self.toolbar = NavigationToolbar(self.canvas)
        self.radio_1_button= wx.RadioButton(self.panel,-1,"EE1.2")
        self.radio_2_button= wx.RadioButton(self.panel,-1,"EE2.0") 
        for each in [self.radio_1_button,self.radio_2_button]:
            self.Bind(wx.EVT_RADIOBUTTON,self.model_data_input,each)            
        self.EE1_path = wx.TextCtrl(self.panel, -1)
        self.EE2_path = wx.TextCtrl(self.panel, -1)
        self.message_out1 = wx.TextCtrl(self.panel, -1,style=wx.TE_MULTILINE)
        self.message_out2 = wx.TextCtrl(self.panel, -1,style=wx.TE_MULTILINE)
        self.Data_xy = wx.TextCtrl(self.panel,-1,"",(20, 60), (233, 20),style=wx.TE_CENTER)
        self.EE1_open_button = wx.Button(self.panel, 0, label='原型EE')
        self.EE2_open_button = wx.Button(self.panel, 1, label='派生EE')
        self.EE_plot_button = wx.Button(self.panel, 2, label='EE绘图')
        self.EE_comp_button = wx.Button(self.panel, 4, label='EE对比')
        self.EE_fre_button=wx.Button(self.panel,5,label='禁止频率检查')
        self.EE_fun_button=wx.Button(self.panel,6,label='功能检查')
        
        self.Bind(wx.EVT_BUTTON, lambda evt, i=self.EE1_open_button.GetId(): self.EE_data_read(evt, i),
                  self.EE1_open_button)
        self.Bind(wx.EVT_BUTTON, lambda evt, i=self.EE2_open_button.GetId(): self.EE_data_read(evt, i),
                  self.EE2_open_button)
        self.Bind(wx.EVT_BUTTON,self.EE_data_plot, self.EE_plot_button)
        self.Bind(wx.EVT_BUTTON,self.EE_data_comp, self.EE_comp_button)
        self.Bind(wx.EVT_BUTTON,self.EE_fre_check, self.EE_fre_button)
        self.Bind(wx.EVT_BUTTON,self.EE_fun_check, self.EE_fun_button)
        
        h1_box_sizer = wx.BoxSizer(wx.VERTICAL)        
        h2_box_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h3_box_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h4_box_sizer = wx.BoxSizer(wx.VERTICAL)
        h5_box_sizer = wx.GridSizer(2, 4, 0.01, 0.01)
        h6_box_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h7_box_sizer = wx.BoxSizer(wx.HORIZONTAL)
        h1_box_sizer.Add(self.radio_1_button, proportion=1, flag=wx.EXPAND | wx.ALL, border=2)
        h1_box_sizer.Add(self.radio_2_button, proportion=1, flag=wx.EXPAND | wx.ALL, border=2)
        h2_box_sizer.Add(self.Data_xy, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        h2_box_sizer.Add(self.EE1_path, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        h2_box_sizer.Add(self.EE1_open_button, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        h3_box_sizer.Add(self.toolbar, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        h3_box_sizer.Add(self.EE2_path, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        h3_box_sizer.Add(self.EE2_open_button, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        h4_box_sizer.Add(h2_box_sizer, proportion=0, flag=wx.EXPAND | wx.ALL)
        h4_box_sizer.Add(h3_box_sizer, proportion=0, flag=wx.EXPAND | wx.ALL)
        h6_box_sizer.Add(h4_box_sizer, proportion=1, flag=wx.EXPAND | wx.ALL, border=5)
        h6_box_sizer.Add(h1_box_sizer, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        h6_box_sizer.Add(h5_box_sizer, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        h5_box_sizer.Add(self.EE_plot_button, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        h5_box_sizer.Add(self.EE_comp_button, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        h5_box_sizer.Add(self.EE_fre_button, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        h5_box_sizer.Add(self.EE_fun_button, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        h7_box_sizer.Add(self.message_out1,proportion=1,flag=wx.EXPAND, border=2)
        h7_box_sizer.Add(self.message_out2,proportion=1,flag=wx.EXPAND, border=2)
        h7_box_sizer.Add(self.grid,proportion=3,flag= wx.LEFT | wx.RIGHT , border=2)               
        v_box_sizer = wx.BoxSizer(wx.VERTICAL)
        v_box_sizer.Add(h6_box_sizer, proportion=0, flag=wx.EXPAND, border=2)        
        v_box_sizer.Add(self.canvas, proportion=1, flag=wx.EXPAND, border=2)
        v_box_sizer.Add(h7_box_sizer, proportion=0, flag=wx.LEFT | wx.RIGHT, border=2)
        self.panel.SetSizer(v_box_sizer)       

    def model_data_input(self, event): 
        '''
        导入EE模板
        '''        
        radioselected=event.GetEventObject()
        self.str_sel=radioselected.GetLabel()        
        fname = 'EE2.0.xlsx'        
        data = xlrd.open_workbook(fname)
        if self.str_sel=="EE1.2":
            tab = data.sheets()[1]            
            col_num_start=3            
        else:
            tab = data.sheets()[2]            
            col_num_start=2
        self.grid.SetColSize(0, 150)
        self.grid.SetColSize(1, 55)
        self.grid.SetColSize(2, 55)
        self.grid.SetColSize(3, 75)
        self.grid.SetColSize(4, 55)
        self.grid.SetColSize(5, 55)
        self.grid.SetColSize(6, 55)
        self.grid.SetColSize(7, 60)        
        self.grid.SetColSize(8, 60)
        self.grid.SetColSize(9, 500)
        nrows = tab.nrows
        ncols = tab.ncols        
        self.grid.ClearGrid()                
        for x in range(0, 10):
            cv= str(tab.cell(0, x +col_num_start).value).strip().replace(' ','').replace("\n",'')
            self.grid.SetColLabelValue(x, cv )
            pass
        for x in range(2, nrows+1):
            for y in range(col_num_start, ncols):
                if self.grid.GetNumberRows() < x:
                    self.grid.AppendRows(numRows=1)
                if self.grid.GetNumberCols() <= y:
                    self.grid.AppendCols(numCols=1)
                self.grid.SetCellValue(x - 2, y - col_num_start, str(tab.cell(x-1, y).value))  
        self.grid.Refresh()
        
    def EE_data_read(self, event, index):
        '''
        打开EE文件，读入EE设定值
        '''        
        filterFile = "Excel files(*.xlsx)|*.xlsx|Excel files(*.xls)|*.xls"
        dlg = wx.FileDialog(self, "Open Excel file...", os.getcwd(), style=wx.FD_OPEN, wildcard=filterFile)
        if dlg.ShowModal() == wx.ID_OK:
            filename = dlg.GetPath()
            if index == 0:
                self.EE1_path.SetValue(filename)
            elif index == 1:
                self.EE2_path.SetValue(filename)
        dlg.Destroy()
        data = xlrd.open_workbook(filename)
        dia_1 = wx.SingleChoiceDialog(None, '选择哪个sheet？', '提示', data.sheet_names())
        if dia_1.ShowModal() == wx.ID_OK:
            num_1=dia_1.GetSelection()
        tab = data.sheets()[num_1]
        nrows = tab.nrows
        ncols = tab.ncols
        array_1=[]         #读入EE数据列表
        array_1_rota=[]    #读入EE数据转置后列表
        array_2=[]         #模板EE数据列表
        self.message_out1.Clear()
        self.message_out2.Clear()
        for i in range(nrows):
            for j in range(ncols):
                cv= str(tab.cell(i,j).value).strip().replace(' ','').replace("\n",'')
                if cv == '参数代号':
                    row_1=i
                    col_1=j
                if cv  == '设定值':
                    col_2=j
                if cv == '单位':
                    col_3=j
                if cv == '参数名称':
                    col_4 =j
        for i in range(row_1+1,nrows):
            cv1 = str(tab.cell(i, col_1).value)
            cv2 = str(tab.cell(i, col_2).value)
            if cv2!="":
                cv2=cv2.rstrip('0').strip('.') if '.' in cv2 else cv2
            cv3 = str(tab.cell(i, col_3).value)
            cv4 = str(tab.cell(i, col_4).value)
            if cv4 != "":
                array_1.append([cv1,cv2,cv3,cv4])
        array_1_rota=list(map(list,zip(*array_1)))
        n=0        
        while self.grid.GetCellValue(n,0) != "":
            self.grid.SetCellBackgroundColour(n, 7 + index, 'white')
            self.grid.SetCellValue(n, 7 + index, "")
            array_2.append(self.grid.GetCellValue(n,0))
            n = n + 1        
        array_1_set=set(array_1_rota[3])
        array_2_set=set(array_2)
        if array_1_rota[3]==array_2 :
            for n in range(len(array_1)):                
                self.grid.SetCellValue(n, 7+index,array_1[n][1])                
        elif array_1_set < array_2_set :            
            dlg = wx.MessageDialog(None, "EE格式不对应，是否继续？", "错误提示", wx.YES_NO | wx.ICON_QUESTION)
            self.message_out.AppendText("缺少部分EE数据\r\n")
            for da in array_2_set.difference(array_1_set):
                self.message_out1.AppendText(da+"\r\n")
            if dlg.ShowModal() == wx.ID_YES:
                for j in range(len(array_1)):
                    for k in range(len(array_2)):
                       if array_1[j][3] == self.grid.GetCellValue(k,0):                           
                           self.grid.SetCellValue(k, 7+index,array_1[j][1])                           
            dlg.Destroy()
        elif array_1_set > array_2_set :            
            dlg = wx.MessageDialog(None, "EE格式不对应，是否继续？", "错误提示", wx.YES_NO | wx.ICON_QUESTION)
            self.message_out.AppendText("多余部分EE数据\r\n")
            for da in array_1_set.difference(array_2_set):
                self.message_out2.AppendText(da+"\r\n")
            if dlg.ShowModal() == wx.ID_YES: 
                for j in range(len(array_1)):
                    for k in range(len(array_2)):
                       if array_1[j][3] == self.grid.GetCellValue(k,0):                           
                           self.grid.SetCellValue(k, 7+index,array_1[j][1])                                       
            dlg.Destroy()
        else:
            dlg = wx.MessageDialog(None, "EE格式不对应，是否继续？", "错误提示", wx.YES_NO | wx.ICON_QUESTION)
            self.message_out1.AppendText("缺少部分EE数据\r\n")
            for da in array_2_set.difference(array_1_set):
                self.message_out1.AppendText(da+"\r\n")
            self.message_out2.AppendText("多余部分EE数据\r\n")
            for da in array_1_set.difference(array_2_set):
                self.message_out2.AppendText(da+"\r\n")
            if dlg.ShowModal() == wx.ID_YES:                   
                for j in range(len(array_1)):
                    for k in range(len(array_2)):
                       if array_1[j][3] == self.grid.GetCellValue(k,0):                           
                           self.grid.SetCellValue(k, 7+index,array_1[j][1])                                
            dlg.Destroy()
      
    def EE_data_comp(self,event): 
        '''
        EE数据对比并输出
        '''
        nrows=self.grid.GetNumberRows()        
        for n in range(nrows):
            self.grid.SetCellBackgroundColour(n,7,"white")
            self.grid.SetCellBackgroundColour(n,8,"white")
            if self.grid.GetCellValue(n,7) != self.grid.GetCellValue(n,8):                
                self.grid.SetCellBackgroundColour(n,7,"red")
                self.grid.SetCellBackgroundColour(n,8,"red")          
        self.grid.Refresh()
        workbook = xlwt.Workbook(encoding='ascii')
        worksheet = workbook.add_sheet('EE差异对比')
        for i in range(self.grid.GetNumberCols()):
            worksheet.write(0, i, self.grid.GetColLabelValue(i))
        n=0
        m=1
        print()
        while n<nrows:
            if self.grid.GetCellBackgroundColour(n,8)=="red":
                for i in range(self.grid.GetNumberCols()):
                    worksheet.write(m, i, self.grid.GetCellValue(n,i))
                m+=1
            n+=1
        workbook.save('temp_EE_comp.xls')
        dlg = wx.MessageDialog(None, "EE差异点已导出到temp_EE_comp.xls文件", "提示", wx.OK)
        if dlg.ShowModal() == wx.OK:
          dlg.Destroy()
          
    def EE_fre_check(self,event):
        shield_frec_set=set()
        shield_freh_set=set()
        self.message_out1.Clear()
        self.message_out2.Clear()
        self.message_out1.AppendText("制冷屏蔽点：\r\n")
        for n in range(self.grid.GetNumberRows()):
            if "屏蔽频率"  in  self.grid.GetCellValue(n,0):
                if "制冷"  in  self.grid.GetCellValue(n,0):
                    if int(self.grid.GetCellValue(n,8)) != 0 :
                        shield_frec_set.add(self.grid.GetCellValue(n,8))
                        self.message_out1.AppendText(self.grid.GetCellValue(n,8)+"、")
        self.message_out1.AppendText("\r\n制热屏蔽点：\r\n")
        for n in range(self.grid.GetNumberRows()):
            if "屏蔽频率"  in  self.grid.GetCellValue(n,0) :
                if "制热"  in  self.grid.GetCellValue(n,0):
                    if int(self.grid.GetCellValue(n,8)) != 0 :
                        shield_freh_set.add(self.grid.GetCellValue(n,8))
                        self.message_out1.AppendText(self.grid.GetCellValue(n,8)+"、")
        for n in range(self.grid.GetNumberRows()):
            self.grid.SetCellBackgroundColour(n,7,"white")
            self.grid.SetCellBackgroundColour(n,8,"white")
            if "屏蔽频率"  not in  self.grid.GetCellValue(n,0):
               if "Hz" in self.grid.GetCellValue(n,5) :
                   if self.grid.GetCellValue(n,8) in shield_frec_set :
                       self.grid.SetCellBackgroundColour(n,8,"yellow")
                       self.message_out2.AppendText("制冷频率冲突\r\n  "+self.grid.GetCellValue(n,0)+":"+self.grid.GetCellValue(n,8)+"\r\n")
                   elif self.grid.GetCellValue(n,8) in shield_freh_set :
                       self.grid.SetCellBackgroundColour(n,8,"red")
                       self.message_out2.AppendText("制热频率冲突\r\n  "+self.grid.GetCellValue(n,0)+":"+self.grid.GetCellValue(n,8)+"\r\n")
        self.grid.Refresh()                                  
        pass
    def EE_fun_check(self,event):
        pass
    def EE_data_plot(self, event):
        '''
        EE数据绘图
        '''  
        if self.str_sel=="EE1.2":
            self.data_plot_1()
        elif self.str_sel == "EE2.0" :
            self.data_plot_2()
        self.axlist=[self.ax1,self.ax2]
        self.figure.canvas.mpl_connect('axes_enter_event', self.enter_axes)
        self.figure.canvas.mpl_connect('axes_leave_event', self.leave_axes)
        self.figure.canvas.mpl_connect('motion_notify_event', self.move)        
        self.figure.Mult = MultiCursor(self.figure.canvas, self.axlist, color='red', lw=1.5, vertOn=True,horizOn=True)
        self.canvas.draw()
    def data_plot_1(self):
        '''
        EE1.2版软件绘图
        '''
        #以下为制冷运行曲线
        list_uplimit_fc=[]#上限
        list_uplimit_tc=[]#上限
        list_downlimit_fc=[]#下限
        list_downlimit_tc=[]#下限
        list_commonlimit_tc=[]#常规线
        list_commonlimit_fc=[]#常规线
        list_fastlimit_tc=[]#快速线
        list_fastlimit_fc=[]#快速线
        list_rated_c=[]
        tc_lowlimit = -20
        tc_highlimit = 60
        n=0
        while self.grid.GetCellValue(n,0) != "【TA制冷】":
            n=n+1
        #上限点列表
        for i in range(n+1,n+10,2):
            list_uplimit_fc.append(int(self.grid.GetCellValue(i,8)))
        for i in range(n,n+9,2):
            list_uplimit_tc.append(int(self.grid.GetCellValue(i,8)))
        list_uplimit_tc.append(60)
        list_uplimit_fc.append(int(self.grid.GetCellValue(n+9,8)))         
        #下限点列表
        list_downlimit_fc.append(int(self.grid.GetCellValue(n+9,8)))
        list_downlimit_tc.append(int(self.grid.GetCellValue(n+8,8)))
        list_downlimit_fc.append(int(self.grid.GetCellValue(n+15,8)))
        list_downlimit_tc.append(int(self.grid.GetCellValue(n+14,8)))
        list_downlimit_fc.append(int(self.grid.GetCellValue(n+13,8)))
        list_downlimit_tc.append(int(self.grid.GetCellValue(n+12,8)))
        list_downlimit_fc.append(int(self.grid.GetCellValue(n+13,8)))
        list_downlimit_tc.append(int(self.grid.GetCellValue(n,8)))
        list_downlimit_fc.append(int(self.grid.GetCellValue(n+1,8)))
        list_downlimit_tc.append(int(self.grid.GetCellValue(n,8)))
        list_downlimit_fc.append(int(self.grid.GetCellValue(n+11,8)))
        list_downlimit_tc.append(int(self.grid.GetCellValue(n+10,8)))
        k_fa_c=(list_downlimit_fc[-1]-list_downlimit_fc[-2])/(list_downlimit_tc[-1]-list_downlimit_tc[-2])
        fc_temp=list_downlimit_fc[-1]+(tc_lowlimit-list_downlimit_tc[-1])*k_fa_c
        list_downlimit_tc.append(tc_lowlimit)
        list_downlimit_fc.append(fc_temp)        
        #常规线列表
        list_rated_c.append(35)
        list_rated_c.append(int(self.grid.GetCellValue(n+16,8)))
        k_ab_c=(list_uplimit_fc[1]-list_uplimit_fc[0])/(list_uplimit_tc[1]-list_uplimit_tc[0])
        k_cd_c=(list_uplimit_fc[3]-list_uplimit_fc[2])/(list_uplimit_tc[3]-list_uplimit_tc[2])
        k_de_c=(list_uplimit_fc[4]-list_uplimit_fc[3])/(list_uplimit_tc[4]-list_uplimit_tc[3])        
        k_gh_c=(list_downlimit_fc[2] - list_downlimit_fc[1]) / (list_downlimit_tc[2] - list_downlimit_tc[1])        
        y_a1_c=list_uplimit_fc[0]-(list_uplimit_fc[1]-int(self.grid.GetCellValue(n+16,8)))
        x_a1_c=list_uplimit_tc[0]
        if y_a1_c<list_downlimit_fc[2]:
            y_a1_c=list_downlimit_fc[2]
            x_a1_c=list_uplimit_tc[1]-(int(self.grid.GetCellValue(n+16,8))-y_a1_c)/k_ab_c
        list_commonlimit_tc.append(x_a1_c)
        list_commonlimit_fc.append(y_a1_c)
        for i in range(3):
            list_commonlimit_tc.append(list_uplimit_tc[i+1])
            list_commonlimit_fc.append(list_uplimit_fc[i+1]-(list_uplimit_fc[1]-int(self.grid.GetCellValue(n+16,8))))
        if list_commonlimit_fc[-1] <list_uplimit_fc[4]:
            list_commonlimit_fc.pop()
            list_commonlimit_fc.append(list_uplimit_fc[4])
            list_commonlimit_tc.pop()
            x_d1_c=list_commonlimit_tc[-1]-(list_commonlimit_fc[-2]-list_commonlimit_fc[-1])/k_cd_c
            list_commonlimit_tc.append(x_d1_c)            
        else:
            list_commonlimit_fc.append(list_downlimit_fc[0])
            x_e1_c=list_commonlimit_tc[-1]-(list_commonlimit_fc[-2]-list_commonlimit_fc[-1])/k_de_c
            list_commonlimit_tc.append(x_e1_c)        
        #快速线列表
        y_a2_c=list_uplimit_fc[0]-(list_uplimit_fc[1]-list_commonlimit_fc[1]-int(self.grid.GetCellValue(n+17,8)))
        x_a2_c=list_uplimit_tc[0]
        if y_a2_c<list_downlimit_fc[2]:
            y_a2_c=list_downlimit_fc[2]
            x_a2_c=list_uplimit_tc[1]-(int(self.grid.GetCellValue(n+16,8))+int(self.grid.GetCellValue(n+17,8))-y_a2_c)/k_ab_c
        list_fastlimit_tc.append(x_a2_c)
        list_fastlimit_fc.append(y_a2_c)     
        for i in range(3):
            list_fastlimit_tc.append(list_uplimit_tc[i+1])
            list_fastlimit_fc.append(list_uplimit_fc[i+1]-(list_uplimit_fc[1]-list_commonlimit_fc[1]-int(self.grid.GetCellValue(n+17,8))))        
        if list_fastlimit_fc[-1]<list_downlimit_fc[0] :
            list_fastlimit_fc.pop()
            list_fastlimit_tc.pop()
            list_fastlimit_fc.append(list_downlimit_fc[0])
            x_d2_c=list_fastlimit_tc[-1]-(list_fastlimit_fc[-2]-list_fastlimit_fc[-1])/k_cd_c
            list_fastlimit_tc.append(x_d2_c)
        else:
            list_fastlimit_fc.append(list_downlimit_fc[0])
            x_e2_c=list_fastlimit_tc[-1]-(list_fastlimit_fc[-2]-list_fastlimit_fc[-1])/k_de_c
            list_fastlimit_tc.append(x_e2_c)
        
        plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
        plt.rcParams['axes.unicode_minus'] = False  # 用来显示正常负号        
        self.figure.clf()
        self.figure.subplots_adjust(left=0.05, right=0.9, top=0.95, bottom=0.12, hspace=0.1, wspace=0.20)
        self.ax1=self.figure.add_subplot(121)
        self.ax1.set_xlim(list_downlimit_tc[-1],list_uplimit_tc[-1])
        self.ax1.set_ylim(0,list_uplimit_fc[1]+10)
        xmajorLocator_c = MultipleLocator(5)
        ymajorLocator_c = MultipleLocator(10)
        xminorLocator_c = MultipleLocator(1)
        yminorLocator_c = MultipleLocator(2)
        self.ax1.xaxis.set_major_locator(xmajorLocator_c)
        self.ax1.yaxis.set_major_locator(ymajorLocator_c)
        self.ax1.xaxis.set_minor_locator(xminorLocator_c)
        self.ax1.yaxis.set_minor_locator(yminorLocator_c)
        self.ax1.spines['bottom'].set_linewidth(2)
        self.ax1.spines['left'].set_linewidth(2)
        self.ax1.tick_params(length=10,width=2,labelsize=16)
        self.ax1.tick_params(which='minor',length=5,width=1)
        self.ax1.set_ylabel('压机频率(Hz)',fontsize=16)
        self.ax1.set_xlabel('外环境温度(℃)',fontsize=16)
        self.ax1.plot(list_uplimit_tc,list_uplimit_fc,color='red',label="制冷上限频率",linewidth=3)
        self.ax1.plot(list_downlimit_tc,list_downlimit_fc,color='blue',label="制冷下限频率",linewidth=3)
        self.ax1.plot(list_commonlimit_tc, list_commonlimit_fc, color='black',label="常规制冷频率", linewidth=3)
        self.ax1.plot(list_fastlimit_tc, list_fastlimit_fc, color='pink', label="快速制冷频率",linewidth=3)
        for i in range(len(list_uplimit_fc)):
            self.ax1.plot([list_uplimit_tc[i],list_uplimit_tc[i]],[0,list_uplimit_fc[i]],'--',color='black',linewidth=0.5)
            self.ax1.plot([tc_lowlimit,list_uplimit_tc[i]],[list_uplimit_fc[i],list_uplimit_fc[i]],'--',color='black',linewidth=0.5)
        for i in range(len(list_downlimit_fc)):
            self.ax1.plot([list_downlimit_tc[i],list_downlimit_tc[i]],[0,list_downlimit_fc[i]],'--',color='black',linewidth=0.5)
            self.ax1.plot([tc_lowlimit,list_downlimit_tc[i]],[list_downlimit_fc[i],list_downlimit_fc[i]],'--',color='black',linewidth=0.5)
        for i in range(len(list_commonlimit_fc)):
            self.ax1.plot([tc_lowlimit,list_commonlimit_tc[i]],[list_commonlimit_fc[i],list_commonlimit_fc[i]],'--',color='black',linewidth=0.5)
        self.ax1.plot([tc_lowlimit, list_downlimit_tc[i]], [list_downlimit_fc[i], list_downlimit_fc[i]], '--',
                      color='black', linewidth=0.5)
        self.ax1.plot([list_rated_c[0],list_rated_c[0]],[0,list_rated_c[1]],'--',color='black',linewidth=0.5)
        self.ax1.plot([list_uplimit_tc[0], list_uplimit_tc[0]], [0, list_uplimit_fc[1]+10],  color='black', linewidth=2)
        self.ax1.legend(loc='upper right', fontsize=10,ncol=1)
        font0 = FontProperties()
        font1 = font0.copy()
        font1.set_size('xx-large')
        fami_c=['A','B','C','D','E','H','G','','','F','','O']
        for i in range(0,5):
            if i==0:
                self.ax1.text(list_uplimit_tc[i]-3,list_uplimit_fc[i]-3,fami_c[i], fontproperties=font1)
            else:
                self.ax1.text(list_uplimit_tc[i], list_uplimit_fc[i]+1, fami_c[i], fontproperties=font1)
        for i in range(5,10):
            if i==7 or i==8:
                pass
            elif i==9:
                self.ax1.text(list_downlimit_tc[i-4], list_downlimit_fc[i-4]+1, fami_c[i], fontproperties=font1)
            else:
                self.ax1.text(list_downlimit_tc[i - 4] + 1, list_downlimit_fc[i - 4] - 3, fami_c[i], fontproperties=font1)
        self.ax1.text(list_rated_c[0]+1, list_rated_c[1] - 3, fami_c[-1], fontproperties=font1)
        self.ax1.text(list_uplimit_tc[0] + 1, list_uplimit_fc[1]+6, "【常规】制冷运行", fontproperties=font1)
        alignment = {'horizontalalignment': 'right', 'verticalalignment': 'baseline'}
        self.ax1.text(list_uplimit_tc[0] -1, list_uplimit_fc[1] + 6, "【低温】制冷运行", fontproperties=font1,**alignment)

        #以下为制热运行曲线
        list_uplimit_fh = []  # 上限
        list_uplimit_th = []  # 上限
        list_downlimit_fh = []  # 下限
        list_downlimit_th = []  # 下限
        list_fastlimit_th = []  # 快速线
        list_fastlimit_fh = []  # 快速线
        list_rated_h = []
        th_lowlimit = -30
        th_highlimit = 33
        m = 0
        while self.grid.GetCellValue(m, 0) != "【TA制热】":
            m = m + 1
        # 上限点列表
        for i in range(m+3, m + 12, 2):
            list_uplimit_fh.append(int(self.grid.GetCellValue(i, 8)))
        for i in range(m + 2, m + 11, 2):
            list_uplimit_th.append(int(self.grid.GetCellValue(i, 8)))
        list_uplimit_fh.pop(2)
        list_uplimit_th.pop(2)
        list_uplimit_fh.insert(0,int(self.grid.GetCellValue(m+3, 8)))
        list_uplimit_th.insert(0,0)
        list_uplimit_fh.insert(0, (int(self.grid.GetCellValue(m+3, 8))+int(self.grid.GetCellValue(m-24, 8)))/2)
        list_uplimit_th.insert(0, 0)
        list_uplimit_fh.insert(0, (int(self.grid.GetCellValue(m + 3, 8)) + int(self.grid.GetCellValue(m - 24, 8))) / 2)
        list_uplimit_th.insert(0, -5)
        list_uplimit_fh.insert(0, int(self.grid.GetCellValue(m - 24, 8)))
        list_uplimit_th.insert(0, -5)
        list_uplimit_fh.insert(0, int(self.grid.GetCellValue(m - 24, 8)))
        list_uplimit_th.insert(0, th_lowlimit)
        k_ef_h = (list_uplimit_fh[-2] - list_uplimit_fh[-1]) / (list_uplimit_th[-2] - list_uplimit_th[-1])
        x_ef_h=list_uplimit_th[-2]-(list_uplimit_fh[-2]-int(self.grid.GetCellValue(m -26, 8)))/k_ef_h
        list_uplimit_th.pop()
        list_uplimit_th.append(x_ef_h)
        list_uplimit_fh.pop()
        list_uplimit_fh.append(int(self.grid.GetCellValue(m -26, 8)))
        #下限列表
        list_downlimit_th.append(th_highlimit)
        list_downlimit_fh.append(int(self.grid.GetCellValue(m -26, 8)))
        list_downlimit_th.append(24)
        list_downlimit_fh.append(int(self.grid.GetCellValue(m - 26, 8)))
        list_downlimit_th.append(18)
        list_downlimit_fh.append(int(self.grid.GetCellValue(m - 27, 8)))
        list_downlimit_th.append(int(self.grid.GetCellValue(m +6, 8)))
        list_downlimit_fh.append(int(self.grid.GetCellValue(m +7, 8)))
        list_downlimit_th.append(int(self.grid.GetCellValue(m , 8)))
        list_downlimit_fh.append(int(self.grid.GetCellValue(m +1, 8)))
        k_ad_h=(list_downlimit_fh[-1]-list_downlimit_fh[-2])/(list_downlimit_th[-1]-list_downlimit_th[-2])
        k_temp_h=(list_downlimit_fh[2]-list_downlimit_fh[1])/(list_downlimit_th[2]-list_downlimit_th[1])
        fh_temp = list_downlimit_fh[-1] + (-25 - list_downlimit_th[-1]) * k_ad_h
        list_downlimit_th.append(-25)
        list_downlimit_fh.append(fh_temp)
        list_downlimit_th.append(th_lowlimit)
        list_downlimit_fh.append(fh_temp)
        # 快速线列表
        list_rated_h.append(7)
        list_rated_h.append(int(self.grid.GetCellValue(m + 12, 8)))
        list_fastlimit_th.append(list_uplimit_th[-4])
        list_fastlimit_fh.append(list_uplimit_fh[-4])
        list_fastlimit_th.append(7)
        list_fastlimit_fh.append(int(self.grid.GetCellValue(m + 12, 8)))
        list_fastlimit_th.append(list_uplimit_th[-3])
        list_fastlimit_fh.append(int(self.grid.GetCellValue(m + 12, 8)))
        list_fastlimit_th.append(list_uplimit_th[-2])
        list_fastlimit_fh.append(list_uplimit_fh[-2]-(int(self.grid.GetCellValue(m + 5, 8))-int(self.grid.GetCellValue(m + 12, 8))))
        list_fastlimit_fh.append(int(self.grid.GetCellValue(m -26, 8)))
        x_f1_h = list_fastlimit_th[-1] - (list_fastlimit_fh[-2] - list_fastlimit_fh[-1]) / k_ef_h
        list_fastlimit_th.append(x_f1_h)
        if x_f1_h<list_downlimit_th[1]:
            x_f1_h = ((list_downlimit_fh[1] - k_temp_h * list_downlimit_th[1]) - (
                        list_fastlimit_fh[-2] - k_ef_h * list_fastlimit_th[-2])) / (k_ef_h - k_temp_h)
            y_f1_h = list_fastlimit_fh[-2] - k_ef_h * (list_fastlimit_th[-2] - x_f1_h)
            list_fastlimit_th.pop()
            list_fastlimit_fh.pop()
            list_fastlimit_th.append(x_f1_h)
            list_fastlimit_fh.append(y_f1_h)
        self.ax2 = self.figure.add_subplot(122)
        self.ax2.set_xlim(list_downlimit_th[-1], list_downlimit_th[0])
        self.ax2.set_ylim(0, list_uplimit_fh[1] + 10)
        xmajorLocator_h = MultipleLocator(5)
        ymajorLocator_h = MultipleLocator(10)
        xminorLocator_h = MultipleLocator(1)
        yminorLocator_h = MultipleLocator(2)
        self.ax2.xaxis.set_major_locator(xmajorLocator_h)
        self.ax2.yaxis.set_major_locator(ymajorLocator_h)
        self.ax2.xaxis.set_minor_locator(xminorLocator_h)
        self.ax2.yaxis.set_minor_locator(yminorLocator_h)
        self.ax2.spines['bottom'].set_linewidth(2)
        self.ax2.spines['left'].set_linewidth(2)
        self.ax2.tick_params(length=10, width=2, labelsize=16)
        self.ax2.tick_params(which='minor', length=5, width=1)
        self.ax2.set_ylabel('压机频率(Hz)', fontsize=16)
        self.ax2.set_xlabel('外环境温度(℃)', fontsize=16)
        self.ax2.plot(list_uplimit_th, list_uplimit_fh, color='red', label="制热上限频率", linewidth=3)
        self.ax2.plot(list_downlimit_th, list_downlimit_fh, color='blue', label="制热下限频率", linewidth=3)
        self.ax2.plot(list_fastlimit_th, list_fastlimit_fh, color='pink', label="快速制热频率", linewidth=3)
        self.ax2.plot([list_uplimit_th[-1],int(self.grid.GetCellValue(m+10, 8))],[list_uplimit_fh[-1],int(self.grid.GetCellValue(m+11, 8))],'--',color='red',linewidth=3)
        for i in range(len(list_uplimit_fh)):
            self.ax2.plot([list_uplimit_th[i],list_uplimit_th[i]],[0,list_uplimit_fh[i]],'--',color='black',linewidth=0.5)
            self.ax2.plot([th_lowlimit,list_uplimit_th[i]],[list_uplimit_fh[i],list_uplimit_fh[i]],'--',color='black',linewidth=0.5)
        for i in range(len(list_downlimit_fh)):
            self.ax2.plot([list_downlimit_th[i],list_downlimit_th[i]],[0,list_downlimit_fh[i]],'--',color='black',linewidth=0.5)
            self.ax2.plot([th_lowlimit,list_downlimit_th[i]],[list_downlimit_fh[i],list_downlimit_fh[i]],'--',color='black',linewidth=0.5)
        self.ax2.plot([th_lowlimit, 7], [int(self.grid.GetCellValue(m+12, 8)), int(self.grid.GetCellValue(m+12, 8))], '--',color='black', linewidth=0.5)
        self.ax2.plot([7, 7],[0, int(self.grid.GetCellValue(m + 12, 8))], '--',color='black', linewidth=0.5)
        self.ax2.plot([int(self.grid.GetCellValue(m + 10,8)), int(self.grid.GetCellValue(m + 10, 8))], [0, int(self.grid.GetCellValue(m + 11, 8))], '--', color='black', linewidth=0.5)
        self.ax2.plot([th_lowlimit, int(self.grid.GetCellValue(m + 10, 8))],
                      [int(self.grid.GetCellValue(m + 11, 8)), int(self.grid.GetCellValue(m + 11, 8))], '--', color='black', linewidth=0.5)
        self.ax2.legend(loc='upper right', fontsize=10, ncol=1)
        font0 = FontProperties()
        font1 = font0.copy()
        font1.set_size('xx-large')
        fami_h = ['B',"C", 'E',"D"]
        for i in range(0,3):
            self.ax2.text(list_uplimit_th[i+5],list_uplimit_fh[i+5]+1,fami_h[i], fontproperties=font1)
        self.ax2.text(int(self.grid.GetCellValue(m+10, 8))+1, int(self.grid.GetCellValue(m + 11, 8)), 'F', fontproperties=font1)
        self.ax2.text(7 , int(self.grid.GetCellValue(m+12, 8)), 'O', fontproperties=font1)
        self.ax2.text(list_downlimit_th[3], list_downlimit_fh[3]+1, 'D', fontproperties=font1)
        self.ax2.text(list_downlimit_th[4], list_downlimit_fh[4]+1, 'A', fontproperties=font1)
        
    def data_plot_2(self):
        '''
        EE2.0版软件绘图
        '''
        #以下为制冷运行曲线
        list_uplimit_fc=[]#上限
        list_uplimit_tc=[]#上限
        list_downlimit_fc=[]#下限
        list_downlimit_tc=[]#下限
        list_commonlimit_tc=[]#常规线
        list_commonlimit_fc=[]#常规线
        list_fastlimit_tc=[]#快速线
        list_fastlimit_fc=[]#快速线
        list_rated_c=[]
        n=0
        while self.grid.GetCellValue(n,0) != "【FA制冷】":
            n=n+1
        #上限点列表
        for i in range(n,n+11,2):
            list_uplimit_fc.append(int(self.grid.GetCellValue(i,8)))
        for i in range(n+1,n+12,2):
            list_uplimit_tc.append(int(self.grid.GetCellValue(i,8)))
        list_uplimit_tc.append(int(self.grid.GetCellValue(n+26,8)))
        list_uplimit_fc.append(int(self.grid.GetCellValue(n+10,8)))
        #下限点列表
        for i in range(n+10,n+21,2):
            list_downlimit_fc.append(int(self.grid.GetCellValue(i,8)))
        for i in range(n+11,n+22,2):
            list_downlimit_tc.append(int(self.grid.GetCellValue(i,8)))
        list_downlimit_tc.insert(5,int(self.grid.GetCellValue(n+1,8)))
        list_downlimit_fc.insert(5,int(self.grid.GetCellValue(n+18,8)))
        list_downlimit_tc.insert(6, int(self.grid.GetCellValue(n + 1, 8)))
        list_downlimit_fc.insert(6, int(self.grid.GetCellValue(n, 8)))
        tc_lowlimit=int(self.grid.GetCellValue(n+25,8))
        tc_highlimit=int(self.grid.GetCellValue(n+26,8))
        k_ka_c=(list_downlimit_fc[-1]-list_downlimit_fc[-2])/(list_downlimit_tc[-1]-list_downlimit_tc[-2])
        fc_temp=list_downlimit_fc[-1]+(tc_lowlimit-list_downlimit_tc[-1])*k_ka_c
        list_downlimit_tc.append(tc_lowlimit)
        list_downlimit_fc.append(fc_temp)
        #常规线列表
        list_rated_c.append(35)
        list_rated_c.append(int(self.grid.GetCellValue(n+2,8))-int(self.grid.GetCellValue(n+22,8)))
        k_ab_c=(list_uplimit_fc[1]-list_uplimit_fc[0])/(list_uplimit_tc[1]-list_uplimit_tc[0])
        k_cd_c=(list_uplimit_fc[3]-list_uplimit_fc[2])/(list_uplimit_tc[3]-list_uplimit_tc[2])
        k_de_c=(list_uplimit_fc[4]-list_uplimit_fc[3])/(list_uplimit_tc[4]-list_uplimit_tc[3])
        k_ef_c = (list_uplimit_fc[5] - list_uplimit_fc[4]) / (list_uplimit_tc[5] - list_uplimit_tc[4])
        k_gh_c=(list_downlimit_fc[2] - list_downlimit_fc[1]) / (list_downlimit_tc[2] - list_downlimit_tc[1])
        k_ij_c=(list_downlimit_fc[4] - list_downlimit_fc[3]) / (list_downlimit_tc[4] - list_downlimit_tc[3])
        y_a1_c=list_uplimit_fc[0]-int(self.grid.GetCellValue(n+22,8))
        x_a1_c=list_uplimit_tc[0]
        if y_a1_c<list_downlimit_fc[5]:
            y_a1_c=list_downlimit_fc[5]
            x_a1_c=list_uplimit_tc[1]-(list_uplimit_fc[1]-int(self.grid.GetCellValue(n+22,8))-y_a1_c)/k_ab_c
        list_commonlimit_tc.append(x_a1_c)
        list_commonlimit_fc.append(y_a1_c)
        for i in range(4):
            list_commonlimit_tc.append(list_uplimit_tc[i+1])
            list_commonlimit_fc.append(list_uplimit_fc[i+1]-int(self.grid.GetCellValue(n+22,8)))
        if (list_commonlimit_fc[-1]-int(self.grid.GetCellValue(n+22,8))) <list_uplimit_fc[5]:
            list_commonlimit_fc.pop()
            list_commonlimit_fc.append(list_uplimit_fc[5])
            list_commonlimit_tc.pop()
            x_e1_c=list_commonlimit_tc[-1]-(list_commonlimit_fc[-2]-list_commonlimit_fc[-1])/k_de_c
            list_commonlimit_tc.append(x_e1_c)
            if list_commonlimit_tc[-1]<list_downlimit_tc[1]:
                x_e1_c=((list_downlimit_fc[1]-k_gh_c*list_downlimit_tc[1])-(list_commonlimit_fc[-2]-k_de_c*list_commonlimit_tc[-2]))/(k_de_c-k_gh_c)
                y_e1_c=list_commonlimit_fc[-2]-k_de_c*(list_commonlimit_tc[-2]-x_e1_c)
                list_commonlimit_fc.pop()
                list_commonlimit_fc.append(y_e1_c)
                list_commonlimit_tc.pop()
                list_commonlimit_tc.append(x_e1_c)
        else:
            list_commonlimit_fc.append(list_downlimit_fc[0])
            x_f1_c=list_commonlimit_tc[-1]-(list_commonlimit_fc[-2]-list_commonlimit_fc[-1])/k_ef_c
            list_commonlimit_tc.append(x_f1_c)
        #快速线列表
        for i in range(5):
            list_fastlimit_tc.append(list_uplimit_tc[i])
            list_fastlimit_fc.append(list_uplimit_fc[i]-int(self.grid.GetCellValue(n+22,8))*float(self.grid.GetCellValue(n+23,8)))
        if list_fastlimit_fc[-1]<list_downlimit_fc[0] :
            list_fastlimit_fc.pop()
            list_fastlimit_tc.pop()
            list_fastlimit_fc.append(list_downlimit_fc[0])
            x_e2_c=list_fastlimit_tc[-1]-(list_fastlimit_fc[-2]-list_fastlimit_fc[-1])/k_de_c
            list_fastlimit_tc.append(x_e2_c)
        else:
            list_fastlimit_fc.append(list_downlimit_fc[0])
            x_f2_c=list_fastlimit_tc[-1]-(list_fastlimit_fc[-2]-list_fastlimit_fc[-1])/k_ef_c
            list_fastlimit_tc.append(x_f2_c)

        plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
        plt.rcParams['axes.unicode_minus'] = False  # 用来显示正常负号        
        self.figure.clf()
        self.figure.subplots_adjust(left=0.05, right=0.9, top=0.95, bottom=0.12, hspace=0.1, wspace=0.20)
        self.ax1=self.figure.add_subplot(121)
        self.ax1.set_xlim(list_downlimit_tc[-1],list_uplimit_tc[-1])
        self.ax1.set_ylim(0,list_uplimit_fc[1]+10)
        xmajorLocator_c = MultipleLocator(5)
        ymajorLocator_c = MultipleLocator(10)
        xminorLocator_c = MultipleLocator(1)
        yminorLocator_c = MultipleLocator(2)
        self.ax1.xaxis.set_major_locator(xmajorLocator_c)
        self.ax1.yaxis.set_major_locator(ymajorLocator_c)
        self.ax1.xaxis.set_minor_locator(xminorLocator_c)
        self.ax1.yaxis.set_minor_locator(yminorLocator_c)
        self.ax1.spines['bottom'].set_linewidth(2)
        self.ax1.spines['left'].set_linewidth(2)
        self.ax1.tick_params(length=10,width=2,labelsize=16)
        self.ax1.tick_params(which='minor',length=5,width=1)
        self.ax1.set_ylabel('压机频率(Hz)',fontsize=16)
        self.ax1.set_xlabel('外环境温度(℃)',fontsize=16)
        self.ax1.plot(list_uplimit_tc,list_uplimit_fc,color='red',label="制冷上限频率",linewidth=3)
        self.ax1.plot(list_downlimit_tc,list_downlimit_fc,color='blue',label="制冷下限频率",linewidth=3)
        self.ax1.plot(list_commonlimit_tc, list_commonlimit_fc, color='black',label="常规制冷频率", linewidth=3)
        self.ax1.plot(list_fastlimit_tc, list_fastlimit_fc, color='pink', label="快速制冷频率",linewidth=3)
        for i in range(len(list_uplimit_fc)):
            self.ax1.plot([list_uplimit_tc[i],list_uplimit_tc[i]],[0,list_uplimit_fc[i]],'--',color='black',linewidth=0.5)
            self.ax1.plot([tc_lowlimit,list_uplimit_tc[i]],[list_uplimit_fc[i],list_uplimit_fc[i]],'--',color='black',linewidth=0.5)
        for i in range(len(list_downlimit_fc)):
            self.ax1.plot([list_downlimit_tc[i],list_downlimit_tc[i]],[0,list_downlimit_fc[i]],'--',color='black',linewidth=0.5)
            self.ax1.plot([tc_lowlimit,list_downlimit_tc[i]],[list_downlimit_fc[i],list_downlimit_fc[i]],'--',color='black',linewidth=0.5)
        for i in range(len(list_commonlimit_fc)):
            self.ax1.plot([tc_lowlimit,list_commonlimit_tc[i]],[list_commonlimit_fc[i],list_commonlimit_fc[i]],'--',color='black',linewidth=0.5)
        self.ax1.plot([tc_lowlimit, list_downlimit_tc[i]], [list_downlimit_fc[i], list_downlimit_fc[i]], '--',
                      color='black', linewidth=0.5)
        self.ax1.plot([list_rated_c[0],list_rated_c[0]],[0,list_rated_c[1]],'--',color='black',linewidth=0.5)
        self.ax1.plot([list_uplimit_tc[0], list_uplimit_tc[0]], [0, list_uplimit_fc[1]+10],  color='black', linewidth=2)
        self.ax1.legend(loc='upper right', fontsize=10,ncol=1)
        font0 = FontProperties()
        font1 = font0.copy()
        font1.set_size('xx-large')
        fami_c=['A','B','C','D','E','F','G','H','I','J','','','K','O']
        for i in range(0,6):
            if i==0:
                self.ax1.text(list_uplimit_tc[i]-3,list_uplimit_fc[i]-3,fami_c[i], fontproperties=font1)
            else:
                self.ax1.text(list_uplimit_tc[i], list_uplimit_fc[i]+1, fami_c[i], fontproperties=font1)
        for i in range(6,13):
            if i==10 or i==11:
                pass
            elif i==12:
                self.ax1.text(list_downlimit_tc[i-5], list_downlimit_fc[i-5]+1, fami_c[i], fontproperties=font1)
            else:
                self.ax1.text(list_downlimit_tc[i - 5] + 1, list_downlimit_fc[i - 5] - 3, fami_c[i], fontproperties=font1)
        self.ax1.text(list_rated_c[0]+1, list_rated_c[1] - 3, fami_c[-1], fontproperties=font1)
        self.ax1.text(list_uplimit_tc[0] + 1, list_uplimit_fc[1]+6, "【常规】制冷运行", fontproperties=font1)
        alignment = {'horizontalalignment': 'right', 'verticalalignment': 'baseline'}
        self.ax1.text(list_uplimit_tc[0] -1, list_uplimit_fc[1] + 6, "【低温】制冷运行", fontproperties=font1,**alignment)

        #以下为制热运行曲线
        list_uplimit_fh = []  # 上限
        list_uplimit_th = []  # 上限
        list_downlimit_fh = []  # 下限
        list_downlimit_th = []  # 下限
        list_commonlimit_th = []  # 常规线
        list_commonlimit_fh = []  # 常规线
        list_fastlimit_th = []  # 快速线
        list_fastlimit_fh = []  # 快速线
        list_rated_h = []
        m = 0
        while self.grid.GetCellValue(m, 0) != "【FA制热】":
            m = m + 1
        # 上限点列表
        for i in range(m, m + 14, 2):
            list_uplimit_fh.append(int(self.grid.GetCellValue(i, 8)))
        for i in range(m + 1, m + 15, 2):
            list_uplimit_th.append(int(self.grid.GetCellValue(i, 8)))
        th_lowlimit = -25
        list_uplimit_th.insert(0,th_lowlimit)
        list_uplimit_fh.insert(0, list_uplimit_fh[0])
        list_uplimit_fh.insert(5, list_uplimit_fh[5])
        list_uplimit_th.insert(5, list_uplimit_th[4])
        list_uplimit_fh.insert(4,list_uplimit_fh[4])
        list_uplimit_th.insert(4, list_uplimit_th[3])
        list_uplimit_fh.insert(3, list_uplimit_fh[3])
        list_uplimit_th.insert(3, list_uplimit_th[2])
        list_uplimit_fh.insert(2, list_uplimit_fh[2])
        list_uplimit_th.insert(2, list_uplimit_th[1])
        k_fg_h = (list_uplimit_fh[-2] - list_uplimit_fh[-1]) / (list_uplimit_th[-2] - list_uplimit_th[-1])
        x_fg_h=list_uplimit_th[-2]-(list_uplimit_fh[-2]-int(self.grid.GetCellValue(m + 14, 8)))/k_fg_h
        list_uplimit_th.pop()
        list_uplimit_th.append(x_fg_h)
        list_uplimit_fh.pop()
        list_uplimit_fh.append(int(self.grid.GetCellValue(m + 14, 8)))
        #下限列表
        list_downlimit_th.append(int(self.grid.GetCellValue(m + 28, 8)))
        list_downlimit_fh.append(int(self.grid.GetCellValue(m + 14, 8)))
        for i in range(m+14, m + 26, 2):
            list_downlimit_fh.append(int(self.grid.GetCellValue(i, 8)))
        for i in range(m + 15, m + 27, 2):
            list_downlimit_th.append(int(self.grid.GetCellValue(i, 8)))
        k_lm_h=(list_downlimit_fh[-1]-list_downlimit_fh[-2])/(list_downlimit_th[-1]-list_downlimit_th[-2])
        fh_temp = list_downlimit_fh[-1] + (th_lowlimit - list_downlimit_th[-1]) * k_lm_h
        list_downlimit_th.append(th_lowlimit)
        list_downlimit_fh.append(fh_temp)
        # 常规线列表
        list_rated_h.append(8)
        list_rated_h.append(int(self.grid.GetCellValue(m + 8, 8)) - int(self.grid.GetCellValue(m + 26, 8)))
        k_hi_h=(list_downlimit_fh[1] - list_downlimit_fh[2]) / (list_downlimit_th[1] - list_downlimit_th[2])
        for i in range(len(list_uplimit_th)-1):
            list_commonlimit_fh.append(list_uplimit_fh[i]-int(self.grid.GetCellValue(m + 26, 8)))
            list_commonlimit_th.append(list_uplimit_th[i])
        list_commonlimit_fh.append(list_uplimit_fh[-1])
        x_g1_h = list_commonlimit_th[-1] + (list_commonlimit_fh[-1] - list_commonlimit_fh[-2]) / k_fg_h
        list_commonlimit_th.append(x_g1_h)
        if x_g1_h<list_downlimit_th[1]:
            x_g1_h = ((list_downlimit_fh[1] - k_hi_h * list_downlimit_th[1]) - (
                        list_commonlimit_fh[-2] - k_fg_h * list_commonlimit_th[-2])) / (k_fg_h - k_hi_h)
            y_g1_h = list_commonlimit_fh[-2] - k_fg_h * (list_commonlimit_th[-2] - x_g1_h)
            list_commonlimit_th.pop()
            list_commonlimit_fh.pop()
            list_commonlimit_th.append(x_g1_h)
            list_commonlimit_fh.append(y_g1_h)
        # 快速线列表
        for i in range(len(list_uplimit_th) - 1):
            list_fastlimit_fh.append(list_uplimit_fh[i] - int(self.grid.GetCellValue(m + 26, 8))*float(self.grid.GetCellValue(m + 27, 8)))
            list_fastlimit_th.append(list_uplimit_th[i])
        list_fastlimit_fh.append(list_uplimit_fh[-1])
        x_g2_h = list_fastlimit_th[-1] + (list_fastlimit_fh[-1] - list_fastlimit_fh[-2]) / k_fg_h
        list_fastlimit_th.append(x_g2_h)
        if x_g2_h < list_downlimit_th[1]:
            x_g2_h = ((list_downlimit_fh[1] - k_hi_h * list_downlimit_th[1]) - (
                    list_fastlimit_fh[-2] - k_fg_h * list_fastlimit_th[-2])) / (k_fg_h - k_hi_h)
            y_g2_h = list_fastlimit_fh[-2] - k_fg_h * (list_fastlimit_th[-2] - x_g2_h)
            list_fastlimit_th.pop()
            list_fastlimit_fh.pop()
            list_fastlimit_th.append(x_g2_h)
            list_fastlimit_fh.append(y_g2_h)
        self.ax2 = self.figure.add_subplot(122)
        self.ax2.set_xlim(list_downlimit_th[-1], list_downlimit_th[0])
        self.ax2.set_ylim(0, list_uplimit_fh[1] + 10)
        xmajorLocator_h = MultipleLocator(5)
        ymajorLocator_h = MultipleLocator(10)
        xminorLocator_h = MultipleLocator(1)
        yminorLocator_h = MultipleLocator(2)
        self.ax2.xaxis.set_major_locator(xmajorLocator_h)
        self.ax2.yaxis.set_major_locator(ymajorLocator_h)
        self.ax2.xaxis.set_minor_locator(xminorLocator_h)
        self.ax2.yaxis.set_minor_locator(yminorLocator_h)
        self.ax2.spines['bottom'].set_linewidth(2)
        self.ax2.spines['left'].set_linewidth(2)
        self.ax2.tick_params(length=10, width=2, labelsize=16)
        self.ax2.tick_params(which='minor', length=5, width=1)
        self.ax2.set_ylabel('压机频率(Hz)', fontsize=16)
        self.ax2.set_xlabel('外环境温度(℃)', fontsize=16)
        self.ax2.plot(list_uplimit_th, list_uplimit_fh, color='red', label="制热上限频率", linewidth=3)
        self.ax2.plot(list_downlimit_th, list_downlimit_fh, color='blue', label="制热下限频率", linewidth=3)
        self.ax2.plot(list_commonlimit_th, list_commonlimit_fh, color='black', label="常规制热频率", linewidth=3)
        self.ax2.plot(list_fastlimit_th, list_fastlimit_fh, color='pink', label="快速制热频率", linewidth=3)
        self.ax2.plot([list_uplimit_th[-1],int(self.grid.GetCellValue(m+13, 8))],[list_uplimit_fh[-1],int(self.grid.GetCellValue(m+12, 8))],'--',color='red',linewidth=3)
        for i in range(len(list_uplimit_fh)):
            self.ax2.plot([list_uplimit_th[i],list_uplimit_th[i]],[0,list_uplimit_fh[i]],'--',color='black',linewidth=0.5)
            self.ax2.plot([th_lowlimit,list_uplimit_th[i]],[list_uplimit_fh[i],list_uplimit_fh[i]],'--',color='black',linewidth=0.5)
        for i in range(len(list_downlimit_fh)):
            self.ax2.plot([list_downlimit_th[i],list_downlimit_th[i]],[0,list_downlimit_fh[i]],'--',color='black',linewidth=0.5)
            self.ax2.plot([th_lowlimit,list_downlimit_th[i]],[list_downlimit_fh[i],list_downlimit_fh[i]],'--',color='black',linewidth=0.5)
        for i in range(len(list_commonlimit_fh)):
            self.ax2.plot([th_lowlimit,list_commonlimit_th[i]],[list_commonlimit_fh[i],list_commonlimit_fh[i]],'--',color='black',linewidth=0.5)
        self.ax2.plot([th_lowlimit, int(self.grid.GetCellValue(m+13, 8))], [int(self.grid.GetCellValue(m+12, 8)), int(self.grid.GetCellValue(m+12, 8))], '--',color='black', linewidth=0.5)
        self.ax2.plot([int(self.grid.GetCellValue(m+13, 8)), int(self.grid.GetCellValue(m+13, 8))],[0, int(self.grid.GetCellValue(m + 12, 8))], '--',color='black', linewidth=0.5)
        self.ax2.legend(loc='upper right', fontsize=10, ncol=1)
        font0 = FontProperties()
        font1 = font0.copy()
        font1.set_size('xx-large')
        fami_h = ['A',"", 'B',"", 'C',"", 'D',"", 'E', 'F', "", 'H', 'I', 'J', 'K', 'L', 'M', '']
        for i in range(1,12):
            if fami_h[i-1]!="":
                self.ax2.text(list_uplimit_th[i],list_uplimit_fh[i]+1,fami_h[i-1], fontproperties=font1)
        for i in range(1,8):
            if i<4:
                self.ax2.text(list_downlimit_th[i],list_downlimit_fh[i]-4,fami_h[i+10], fontproperties=font1)
            else:
                self.ax2.text(list_downlimit_th[i]+1, list_downlimit_fh[i] , fami_h[i +10], fontproperties=font1)
        self.ax2.text(int(self.grid.GetCellValue(m+13, 8)) , int(self.grid.GetCellValue(m+12, 8)), 'G', fontproperties=font1)
     
    def enter_axes(self, event):  # 鼠标进入事件
        if event.inaxes in self.axlist:
            self.figure.cursor = Cursor(event.inaxes, useblit=True, color='red', lw=1.5)
        self.canvas.draw()
    def leave_axes(self, event):  # 鼠标离开事件
        self.figure.cursor = Cursor(event.inaxes, useblit=False, color='black', lw=1.5)
        self.canvas.draw()
    def move(self,event):
        if event.inaxes in self.axlist:
            self.Data_xy.SetValue("X:"+str(round(event.xdata,1))+"  "+"Y:"+str(round(event.ydata,1)))
            pass  

def main():
    app = MyApp()
    app.MainLoop()


if __name__ == '__main__':
    main()
