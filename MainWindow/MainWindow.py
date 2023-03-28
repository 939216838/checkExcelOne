# -*- coding: utf-8 -*-


import wx
import wx.xrc

from report import check


class MainWindow(wx.Frame):

    def __init__(self, parent):
        self.i = 0
        wx.Frame.__init__(self, parent, id=-1, title=u"check三个表", pos=wx.Point(3300 - 1980, 110),
                          size=wx.Size(380, 110),
                          style=wx.CAPTION | wx.CLOSE_BOX | wx.FRAME_TOOL_WINDOW | wx.STAY_ON_TOP | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)
        self.SetFont(wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "宋体"))
        self.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))
        self.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_APPWORKSPACE))

        bSizer_布局1 = wx.BoxSizer(wx.VERTICAL)

        bSizer_布局1.SetMinSize(wx.Size(420, 110))

        bSizer_布局1.Add((5, 5), 0, 0, 0)

        gbSizer1_布局2 = wx.GridBagSizer(0, 0)
        # gbSizer1_布局2.SetFlexibleDirection(wx.BOTH)
        gbSizer1_布局2.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)

        gbSizer1_布局2.SetMinSize(wx.Size(100, 30))

        # self.m_filePicker4 = wx.FilePickerCtrl(self, wx.ID_ANY, wx.EmptyString, u"选择名单",
        #                                    u"EXCEL文件(*.xls;*.xlsx)|*.xls;*.xlsx|所有文件(*.*)|*.*",
        #                                    wx.DefaultPosition, wx.DefaultSize,
        #                                    wx.FLP_DEFAULT_STYLE | wx.FLP_FILE_MUST_EXIST, wx.DefaultValidator, u"浏览")
        self.m_dirPicker1_文件路径选择器 = wx.DirPickerCtrl(self, wx.ID_ANY,
                                                            u"C:\\Users\\ShanHongFeng\\Desktop\\龙井",
                                                            u"请选择工作表目录",
                                                            wx.DefaultPosition, wx.Size(305, 30),
                                                            wx.DIRP_USE_TEXTCTRL | wx.CLIP_CHILDREN)
        self.m_dirPicker1_文件路径选择器.SetFont(
            wx.Font(9, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "宋体"))

        gbSizer1_布局2.Add(self.m_dirPicker1_文件路径选择器, wx.GBPosition(0, 0), wx.GBSpan(1, 1), wx.ALL, 5)

        self.m_button1 = wx.Button(self, wx.ID_ANY, u"开始", wx.DefaultPosition, wx.Size(50, 30), 0)
        self.m_button1.SetFont(
            wx.Font(9, wx.FONTFAMILY_SWISS, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, "微软雅黑"))
        self.m_button1.SetForegroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_CAPTIONTEXT))
        self.m_button1.SetBackgroundColour(wx.SystemSettings.GetColour(wx.SYS_COLOUR_WINDOW))

        gbSizer1_布局2.Add(self.m_button1, wx.GBPosition(0, 1), wx.GBSpan(1, 1), wx.ALIGN_CENTER_VERTICAL, 1)

        bSizer_布局1.Add(gbSizer1_布局2, 0, 0, 0)

        bSizer2 = wx.BoxSizer(wx.VERTICAL)

        bSizer2.SetMinSize(wx.Size(110, 30))
        self.m_gauge_进度条 = wx.Gauge(self, wx.ID_ANY, 100, wx.Point(-1, -1), wx.Size(380, 20), wx.GA_HORIZONTAL)
        self.m_gauge_进度条.SetValue(0)
        bSizer2.Add(self.m_gauge_进度条, 0, wx.ALL, 5)

        bSizer_布局1.Add(bSizer2, 0, 0, 0)

        self.SetSizer(bSizer_布局1)
        self.Layout()

        # self.Centre(wx.BOTH)

        # Connect Events
        self.m_dirPicker1_文件路径选择器.Bind(wx.EVT_DIRPICKER_CHANGED, self.selectPath)
        self.m_button1.Bind(wx.EVT_BUTTON, self.start)

    def __del__(self):
        pass

    # Virtual event handlers, override them in your derived class
    def selectPath(self, event):
        path = self.m_dirPicker1_文件路径选择器.GetPath()
        print(path)
        event.Skip()

    def start(self, event):
        self.m_gauge_进度条.SetValue(20)

        print("点击了开始按钮")
        # print(type(self))
        check.start(self, self.m_dirPicker1_文件路径选择器.GetPath(), wx)
        event.Skip()


if __name__ == '__main__':
    # 创建程序对象
    app = wx.App()
    # 创建窗口对象
    window = MainWindow(parent=None)

    # 显示窗口
    window.Show()
    # 进入主事件循环
    app.MainLoop()
