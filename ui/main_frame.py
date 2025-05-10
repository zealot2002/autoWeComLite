import wx
from ui.send_panel import SendPanel
from ui.settings_panel import SettingsPanel

class MainFrame(wx.Frame):
    def __init__(self, parent, title):
        super().__init__(parent, title=title, size=(900, 600))
        self.SetMinSize((700, 500))
        self._init_ui()
        self.Centre()
        self.Show(True)

    def _init_ui(self):
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 顶部按钮栏
        hbox_nav = wx.BoxSizer(wx.HORIZONTAL)
        self.btn_send = wx.Button(panel, label="消息发送")
        self.btn_settings = wx.Button(panel, label="设置")
        self.btn_send.Bind(wx.EVT_BUTTON, lambda evt: self.show_panel('send'))
        self.btn_settings.Bind(wx.EVT_BUTTON, lambda evt: self.show_panel('settings'))
        hbox_nav.Add(self.btn_send, 0, wx.RIGHT, 5)
        hbox_nav.Add(self.btn_settings, 0)
        vbox.Add(hbox_nav, 0, wx.EXPAND|wx.ALL, 8)

        # 内容区：两个页面
        self.panels = {}
        self.panels['send'] = SendPanel(panel, on_send_callback=self.on_send_message)
        self.panels['settings'] = SettingsPanel(panel)
        for p in self.panels.values():
            p.Hide()
            vbox.Add(p, 1, wx.EXPAND|wx.ALL, 0)
        self.current_panel = None
        self.show_panel('send')

        # 状态栏
        self.CreateStatusBar()
        self.SetStatusText("Ready")

        panel.SetSizer(vbox)

    def show_panel(self, name):
        if self.current_panel:
            self.current_panel.Hide()
        self.current_panel = self.panels[name]
        self.current_panel.Show()
        self.Layout()

    def on_send_message(self, contact, message):
        # 这里后续集成自动化逻辑
        self.panels['send'].add_log(f"[模拟] 发送给 {contact}：{message}") 