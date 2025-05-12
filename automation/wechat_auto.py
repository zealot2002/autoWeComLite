import platform
import pygetwindow as gw
import pyperclip
from pywinauto import Desktop
import time
import win32gui
import win32con
from core.config_manager import ConfigManager

# Windows only
def try_import_pywinauto():
    try:
        import pywinauto
        return pywinauto
    except ImportError:
        return None

class WeChatAutomation:
    def __init__(self, logger=None, config_path=None):
        self.logger = logger or (lambda msg: print(msg))
        self.is_mac = platform.system() == "Darwin"
        self.is_win = platform.system() == "Windows"
        self.pywinauto = try_import_pywinauto() if self.is_win else None
        self.window_keywords = ["weixin","WeCom","WeChat","qiyeweixin"]
        
        # 加载配置
        self.config_manager = ConfigManager(config_path)
        self.control_configs = {}
        self.timeouts = {}
        self.strategies = {}
        self._load_configs()
        
    def _load_configs(self):
        """加载所有相关配置"""
        # Windows平台控件配置
        if self.is_win:
            self.control_configs["search_box"] = self.config_manager.get_control_config("search_box")
            self.control_configs["message_input"] = self.config_manager.get_control_config("message_input")
            self.control_configs["search_result_list"] = self.config_manager.get_control_config("search_result_list")
            self.control_configs["search_result_item"] = self.config_manager.get_control_config("search_result_item")
            self.control_configs["chat_title"] = self.config_manager.get_control_config("chat_title")
        
        # 超时设置
        self.timeouts["search_result_wait"] = self.config_manager.get_timeout("search_result_wait")
        self.timeouts["chat_window_load"] = self.config_manager.get_timeout("chat_window_load")
        self.timeouts["input_focus"] = self.config_manager.get_timeout("input_focus")
        self.timeouts["typing_pause"] = self.config_manager.get_timeout("typing_pause")
        
        # 策略设置
        self.strategies["search_result_selection"] = self.config_manager.get_strategy("search_result_selection")
        self.strategies["alternative_search_result_selection"] = self.config_manager.get_strategy("alternative_search_result_selection")

    def log(self, msg):
        if self.logger:
            self.logger(msg)

    def focus_wechat_window(self):
        all_windows = gw.getAllTitles()
        self.log(f"[窗口枚举] 当前所有窗口名: {all_windows}")
        candidates = []
        for w in all_windows:
            if w and any(key.lower() in w.lower() for key in self.window_keywords):
                candidates.append(w)
        if not candidates:
            raise RuntimeError("未找到微信/企业微信窗口或激活失败")
        # 优先第一个匹配
        win = gw.getWindowsWithTitle(candidates[0])[0]
        win.activate()
        # 强制前台
        try:
            hwnd = win._hWnd
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            self.log(f"[Win32] 已调用 SetForegroundWindow: {hwnd}")
        except Exception as e:
            self.log(f"[Win32] SetForegroundWindow 失败: {e}")
        time.sleep(0.5)
        active = gw.getActiveWindow()
        self.log(f"尝试激活窗口: {win.title}, 当前活动窗口: {active.title if active else None}")
        if active and (active.title == win.title):
            self.log(f"[确认] 已激活窗口: {win.title}")
            return win
        else:
            self.log(f"[警告] 激活失败，当前活动窗口为: {active.title if active else None}")
            raise RuntimeError("未找到微信/企业微信窗口或激活失败")

    def send_message(self, contact, message):
        try:
            win = self.focus_wechat_window()
            if self.is_win and self.pywinauto:
                self._send_message_windows(win, contact, message)
            else:
                raise RuntimeError("不支持的操作系统")
        except Exception as e:
            self.log(f"[错误] {e}")
            raise

    def _send_message_windows(self, win, contact, message):
        from pywinauto.application import Application
        windows = Desktop(backend="uia").windows()
        self.log("[pywinauta窗口枚举] 所有UIA窗口：")
        for w in windows:
            self.log(f"  title={w.window_text()} class={w.element_info.class_name} handle={w.handle} visible={w.is_visible()} enabled={w.is_enabled()} process={w.element_info.process_id}")
        try:
            app = Application(backend="uia").connect(title=win.title, timeout=5)
        except Exception as e:
            self.log(f"[pywinauta连接异常] {e}")
            raise
        main_win = app.window(title=win.title)
        main_win.set_focus()
        
        # 搜索联系人
        try:
            # 获取搜索框配置
            search_box_config = self.control_configs.get("search_box", {})
            search_box_class = search_box_config.get("class_name", "mmui::XLineEdit")
            search_box_type = search_box_config.get("control_type", "Edit")
            
            # 使用class_name精确查找搜索框
            self.log(f"[搜索框] 使用配置的class_name='{search_box_class}'查找搜索框")
            search_box = main_win.child_window(class_name=search_box_class, control_type=search_box_type)
            self.log(f"[搜索框] 精确定位搜索框 class={search_box.element_info.class_name}, handle={search_box.handle}")
            
            # 聚焦并清空搜索框
            search_box.set_focus()
            time.sleep(self.timeouts.get("typing_pause", 0.3))
            search_box.type_keys('^a{BACKSPACE}', set_foreground=True)
            time.sleep(self.timeouts.get("typing_pause", 0.3))
            
            # 输入联系人名称
            self.log(f"[搜索框] 输入联系人: '{contact}'")
            pyperclip.copy(contact)
            search_box.type_keys('^v', set_foreground=True)  # 先只输入文本，不按回车
            
            # 等待搜索结果显示
            self.log(f"[搜索框] 等待搜索结果加载")
            time.sleep(self.timeouts.get("search_result_wait", 0.5))  # 给足够时间加载搜索结果
            self.log(f"[搜索框] 使用回车键选择联系人")
            search_box.type_keys('{ENTER}', set_foreground=True)
            
            # 等待聊天窗口加载
            self.log(f"[搜索框] 等待聊天窗口加载")
            time.sleep(self.timeouts.get("chat_window_load", 0.5))  # 给足够时间加载聊天窗口
            self.log("[搜索结果] ===== 搜索结果处理完成 =====")
        except Exception as e:
            # 严格错误处理：列出所有Edit控件信息以便调试，但不使用备选方案
            self.log(f"[搜索框查找失败] {e}")
            try:
                edits = main_win.descendants(control_type="Edit")
                edit_info = [f"{i}: 文本:'{e.window_text()}' 类名:{e.element_info.class_name}" for i, e in enumerate(edits)]
                self.log(f"[调试] 所有Edit控件: {edit_info}")
            except Exception as e2:
                self.log(f"[调试信息获取失败] {e2}")
            
            self.log(f"[错误] 无法精确定位搜索框(class_name=\"{search_box_class}\")，操作终止")
            raise RuntimeError(f"无法精确定位搜索框，请确认微信版本兼容性或更新配置文件中的class_name")
        # 输入消息
        try:
            # 获取消息输入框配置
            msg_input_config = self.control_configs.get("message_input", {})
            msg_input_class = msg_input_config.get("class_name", "")
            
            # 找到输入框（根据配置或通用方法）
            edits = main_win.descendants(control_type="Edit")
            input_box = main_win.child_window(class_name=msg_input_class, control_type="Edit")
            
            if input_box:
                self.log("[消息框] 开始输入消息")
                input_box.set_focus()
                time.sleep(self.timeouts.get("input_focus", 0.5))  # 等待聚焦
                input_box.type_keys('^a{BACKSPACE}', set_foreground=True)  # 清空输入框
                time.sleep(self.timeouts.get("typing_pause", 0.3))
                pyperclip.copy(message)  # 复制消息到剪贴板
                input_box.type_keys('^v', set_foreground=True)  # 粘贴
                time.sleep(self.timeouts.get("typing_pause", 0.3))  # 等待消息输入完成
                input_box.type_keys('{ENTER}', set_foreground=True)  # 按回车发送
                self.log(f"[消息] 已发送消息: {message}")
            else:
                raise RuntimeError("未找到消息输入框")
        except Exception as e:
            self.log(f"[消息发送] 失败: {e}")
            raise RuntimeError("发送消息失败")