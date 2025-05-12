import wx
from core.config_manager import ConfigManager

class SettingsPanel(wx.Panel):
    def __init__(self, parent):
        super().__init__(parent)
        self.config_manager = ConfigManager()
        self._init_ui()
        
        # 确保初始布局正确渲染
        self.Layout()
        self.Refresh()

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

        # 微信窗口类名设置
        hbox_wechat_class = wx.BoxSizer(wx.HORIZONTAL)
        lbl_wechat_class = wx.StaticText(self, label="微信窗口类名：")
        
        # 获取当前配置
        main_window_config = self.config_manager.get_control_config("main_window")
        default_class = main_window_config.get("class_name", "mmui::MainWindow") if main_window_config else "mmui::MainWindow"
        
        self.txt_wechat_class = wx.TextCtrl(self, value=default_class)
        self.btn_save_class = wx.Button(self, label="保存", size=(60, -1))
        self.btn_save_class.Bind(wx.EVT_BUTTON, self.on_save_wechat_class)
        
        hbox_wechat_class.Add(lbl_wechat_class, 0, wx.ALIGN_CENTER_VERTICAL|wx.RIGHT, 8)
        hbox_wechat_class.Add(self.txt_wechat_class, 1, wx.RIGHT, 8)
        hbox_wechat_class.Add(self.btn_save_class, 0)
        vbox.Add(hbox_wechat_class, 0, wx.EXPAND|wx.LEFT|wx.RIGHT|wx.BOTTOM, 10)

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
        
    def on_save_wechat_class(self, event):
        wechat_class = self.txt_wechat_class.GetValue().strip()
        if wechat_class:
            # 更新配置
            if "windows" not in self.config_manager.config:
                self.config_manager.config["windows"] = {}
            if "main_window" not in self.config_manager.config["windows"]:
                self.config_manager.config["windows"]["main_window"] = {
                    "control_type": "Window",
                    "description": "微信主窗口类名"
                }
            
            self.config_manager.config["windows"]["main_window"]["class_name"] = wechat_class
            self.config_manager.save_config()
            wx.MessageBox("微信窗口类名设置已保存", "成功", wx.OK | wx.ICON_INFORMATION)
        else:
            wx.MessageBox("微信窗口类名不能为空", "错误", wx.OK | wx.ICON_ERROR)

    def refresh_config_data(self):
        """刷新配置数据，在面板显示时调用"""
        # 重新加载配置
        self.config_manager = ConfigManager()
        
        # 更新界面控件的值
        main_window_config = self.config_manager.get_control_config("main_window")
        current_class = main_window_config.get("class_name", "mmui::MainWindow") if main_window_config else "mmui::MainWindow"
        self.txt_wechat_class.SetValue(current_class)
        
        # 强制刷新布局
        self.Layout()
        self.Refresh()
    
    def Show(self, show=True):
        """重写Show方法，在显示面板时刷新配置"""
        if show:
            self.refresh_config_data()
        return super().Show(show) 