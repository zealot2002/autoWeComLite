import platform
import pyautogui
import pygetwindow as gw
import pyperclip
from pywinauto import Desktop, Application
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
        # 修改关键词列表，使其更精确
        self.window_keywords = ["微信", "Weixin", "WeChat", "企业微信", "WeCom"]
        # 自动化工具窗口的关键词，用于排除
        self.exclude_keywords = ["autoWeComLite", "automation"]
        
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
            self.control_configs["main_window"] = self.config_manager.get_control_config("main_window")
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
        """直接通过配置的类名查找并激活微信窗口"""
        self.log("[窗口查找] 开始查找微信窗口")
        
        if not self.is_win or not self.pywinauto:
            raise RuntimeError("仅支持Windows平台")
        
        # 获取配置的微信窗口类名
        main_window_config = self.control_configs.get("main_window", {})
        wechat_class_name = main_window_config.get("class_name", "mmui::MainWindow")
        
        self.log(f"[窗口查找] 使用类名 '{wechat_class_name}' 查找微信窗口")
        
        # 直接查找特定类名的窗口
        try:
            desktop = Desktop(backend="uia")
            wechat_windows = []
            
            # 查找所有匹配类名的窗口
            for win in desktop.windows():
                try:
                    title = win.window_text()
                    class_name = win.element_info.class_name
                    
                    if wechat_class_name in class_name:
                        self.log(f"[微信窗口] 通过类名找到: '{title}', class='{class_name}'")
                        wechat_windows.append((title, win.handle))
                except Exception as e:
                    continue
            
            if not wechat_windows:
                self.log(f"[错误] 未找到类名为 '{wechat_class_name}' 的微信窗口")
                raise RuntimeError(f"未找到微信窗口，请确保微信已打开")
            
            # 如果找到多个，优先选择第一个
            title, handle = wechat_windows[0]
            self.log(f"[选择] 将激活窗口: '{title}'")
            
            # 通过handle获取窗口对象
            win = None
            for w in gw.getAllWindows():
                if w._hWnd == handle:
                    win = w
                    break
            
            if not win:
                # 备选方案：使用标题查找
                win = gw.getWindowsWithTitle(title)[0]
            
            # 激活窗口
            win.activate()
            
            # 强制前台
            try:
                win32gui.ShowWindow(handle, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(handle)
                self.log(f"[Win32] 已调用 SetForegroundWindow: {handle}")
            except Exception as e:
                self.log(f"[Win32] SetForegroundWindow 失败: {e}")
            
            time.sleep(0.5)
            active = gw.getActiveWindow()
            
            if active and (active.title == win.title):
                self.log(f"[确认] 已激活窗口: {win.title}")
                return win
            else:
                self.log(f"[警告] 激活失败，当前活动窗口为: {active.title if active else None}")
                raise RuntimeError("激活微信窗口失败")
                
        except Exception as e:
            self.log(f"[窗口查找失败] {e}")
            raise RuntimeError(f"查找微信窗口失败: {e}")

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
            # 精确定位搜索框
            search_box = main_win.child_window(class_name="mmui::XLineEdit", control_type="Edit")
            self.log(f"[搜索框] 精确定位 class={search_box.element_info.class_name}")
            
            # 聚焦并清空搜索框
            search_box.set_focus()
            time.sleep(0.1)
            search_box.type_keys('^a{BACKSPACE}', set_foreground=True)
            time.sleep(0.1)
            
            # 输入联系人名称
            self.log(f"[搜索框] 输入联系人: '{contact}'")
            pyperclip.copy(contact)
            search_box.type_keys('^v', set_foreground=True)
            
            # 等待搜索结果显示
            search_result_wait = self.timeouts["search_result_wait"]
            self.log(f"[搜索框] 等待搜索结果加载 (等待 {search_result_wait} 秒)")
            time.sleep(search_result_wait)
            
            
            # 直接按回车选择第一个结果
            self.log(f"[搜索框] 按回车选择联系人")
            search_box.type_keys('{ENTER}', set_foreground=True)
            
            # 等待聊天窗口加载
            chat_window_wait = self.timeouts["chat_window_load"]
            self.log(f"[搜索框] 等待聊天窗口加载 (等待 {chat_window_wait} 秒)")
            time.sleep(chat_window_wait)
        except Exception as e:
            # 严格错误处理：列出所有Edit控件信息以便调试，但不使用备选方案
            self.log(f"[搜索框查找失败] {e}")
            try:
                edits = main_win.descendants(control_type="Edit")
                edit_info = [f"{i}: 文本:'{e.window_text()}' 类名:{e.element_info.class_name}" for i, e in enumerate(edits)]
                self.log(f"[调试] 所有Edit控件: {edit_info}")
            except Exception as e2:
                self.log(f"[调试信息获取失败] {e2}")
            
            self.log(f"[错误] 无法精确定位搜索框(class_name=\"mmui::XLineEdit\")，操作终止")
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

    def _send_message_mac(self, win, contact, message):
        win.activate()
        pyautogui.hotkey('command', 'f')
        pyperclip.copy(contact)
        pyautogui.hotkey('command', 'v')
        
        # 等待搜索结果显示
        search_result_wait = self.timeouts["search_result_wait"]
        self.log(f"[搜索框] 等待搜索结果加载 (等待 {search_result_wait} 秒)")
        time.sleep(search_result_wait)
        
        pyautogui.press('enter')
        
        # 等待聊天窗口加载
        chat_window_wait = self.timeouts["chat_window_load"]
        self.log(f"[搜索框] 等待聊天窗口加载 (等待 {chat_window_wait} 秒)")
        time.sleep(chat_window_wait)
        
        pyautogui.hotkey('command', 'l')  # 聚焦输入框（如支持）
        pyperclip.copy(message)
        pyautogui.hotkey('command', 'v')
        pyautogui.press('enter')
        self.log(f"[成功] 已发送消息给 {contact}")