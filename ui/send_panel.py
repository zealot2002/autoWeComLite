import wx
from automation.wechat_auto import WeChatAutomation

class SendPanel(wx.Panel):
    def __init__(self, parent, on_send_callback=None):
        super().__init__(parent)
        self.on_send_callback = on_send_callback
        self._init_ui()
        self.automation = WeChatAutomation(logger=self.add_log)

    def _init_ui(self):
        vbox = wx.BoxSizer(wx.VERTICAL)

        # 联系人输入
        hbox_contact = wx.BoxSizer(wx.HORIZONTAL)
        lbl_contact = wx.StaticText(self, label="联系人：")
        self.txt_contact = wx.TextCtrl(self)
        hbox_contact.Add(lbl_contact, 0, wx.ALIGN_CENTER_VERTICAL|wx.RIGHT, 8)
        hbox_contact.Add(self.txt_contact, 1)
        vbox.Add(hbox_contact, 0, wx.EXPAND|wx.ALL, 10)

        # 消息内容输入
        hbox_msg = wx.BoxSizer(wx.HORIZONTAL)
        lbl_msg = wx.StaticText(self, label="消息内容：")
        self.txt_msg = wx.TextCtrl(self, style=wx.TE_MULTILINE, size=(-1, 100))
        hbox_msg.Add(lbl_msg, 0, wx.ALIGN_TOP|wx.RIGHT, 8)
        hbox_msg.Add(self.txt_msg, 1, wx.EXPAND)
        vbox.Add(hbox_msg, 0, wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, 10)

        # 发送按钮
        self.btn_send = wx.Button(self, label="发送")
        self.btn_send.Bind(wx.EVT_BUTTON, self._on_send)
        vbox.Add(self.btn_send, 0, wx.ALIGN_RIGHT|wx.RIGHT|wx.BOTTOM, 10)

        # 日志显示
        self.log_ctrl = wx.TextCtrl(self, style=wx.TE_MULTILINE|wx.TE_READONLY|wx.HSCROLL, size=(-1, 120))
        vbox.Add(self.log_ctrl, 0, wx.EXPAND|wx.ALL, 10)

        self.SetSizer(vbox)

    def _on_send(self, event):
        contact = self.txt_contact.GetValue().strip()
        message = self.txt_msg.GetValue().strip()
        if not contact or not message:
            self.add_log("[警告] 联系人和消息内容不能为空！")
            return
        self.btn_send.Disable()
        try:
            self.automation.send_message(contact, message)
        except Exception as e:
            self.add_log(f"[异常] {e}")
        self.btn_send.Enable()

    def add_log(self, msg):
        self.log_ctrl.AppendText(msg + "\n") 