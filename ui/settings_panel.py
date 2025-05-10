import wx

class SettingsPanel(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)
        self._init_ui()

    def _init_ui(self):
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 平台选择
        hbox_platform = wx.BoxSizer(wx.HORIZONTAL)
        lbl_platform = wx.StaticText(self, label="平台：")
        self.choice_platform = wx.Choice(self, choices=["自动检测", "Windows", "macOS"])
        self.choice_platform.SetSelection(0)
        hbox_platform.Add(lbl_platform, 0, wx.ALIGN_CENTER_VERTICAL|wx.RIGHT, 8)
        hbox_platform.Add(self.choice_platform, 1)
        vbox.Add(hbox_platform, 0, wx.EXPAND|wx.ALL, 10)

        # 自动化参数配置（占位）
        hbox_delay = wx.BoxSizer(wx.HORIZONTAL)
        lbl_delay = wx.StaticText(self, label="操作延迟(ms)：")
        self.spin_delay = wx.SpinCtrl(self, min=0, max=2000, initial=200)
        hbox_delay.Add(lbl_delay, 0, wx.ALIGN_CENTER_VERTICAL|wx.RIGHT, 8)
        hbox_delay.Add(self.spin_delay, 1)
        vbox.Add(hbox_delay, 0, wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, 10)

        # 预留更多参数
        vbox.AddStretchSpacer()

        self.SetSizer(vbox) 