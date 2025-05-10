import platform
import pyautogui
import pygetwindow as gw
import pyperclip
from pywinauto import Desktop
import time
import win32gui
import win32con

# Windows only
def try_import_pywinauto():
    try:
        import pywinauto
        return pywinauto
    except ImportError:
        return None

class WeChatAutomation:
    def __init__(self, logger=None):
        self.logger = logger or (lambda msg: print(msg))
        self.is_mac = platform.system() == "Darwin"
        self.is_win = platform.system() == "Windows"
        self.pywinauto = try_import_pywinauto() if self.is_win else None
        self.window_keywords = ["微信", "weixin", "企业微信"]

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
            elif self.is_mac:
                self._send_message_mac(win, contact, message)
            else:
                raise RuntimeError("不支持的操作系统")
        except Exception as e:
            self.log(f"[错误] {e}")
            raise

    def _send_message_windows(self, win, contact, message):
        from pywinauto.application import Application
        windows = Desktop(backend="uia").windows()
        self.log("[pywinauto窗口枚举] 所有UIA窗口：")
        for w in windows:
            self.log(f"  title={w.window_text()} class={w.element_info.class_name} handle={w.handle} visible={w.is_visible()} enabled={w.is_enabled()} process={w.element_info.process_id}")
        try:
            app = Application(backend="uia").connect(title=win.title, timeout=5)
        except Exception as e:
            self.log(f"[pywinauto连接异常] {e}")
            raise
        main_win = app.window(title=win.title)
        main_win.set_focus()
        # 列出所有控件名和详细属性
        try:
            all_ctrls = main_win.descendants()
            ctrl_info = [f"{c.window_text()} ({c.element_info.control_type}) class={c.element_info.class_name} auto_id={c.element_info.automation_id} handle={c.handle}" for c in all_ctrls]
            self.log(f"[控件枚举] 当前窗口所有控件: {ctrl_info}")
        except Exception as e:
            self.log(f"[控件枚举异常] {e}")
        # 搜索联系人
        try:
            # 修改：使用class_name而非title来查找搜索框
            search_box = main_win.child_window(class_name="mmui::XLineEdit", control_type="Edit")
            self.log(f"[发现搜索框] class={search_box.element_info.class_name}, handle={search_box.handle}")
            search_box.set_focus()
            search_box.type_keys('^a{BACKSPACE}', set_foreground=True)
            pyperclip.copy(contact)
            search_box.type_keys('^v{ENTER}', set_foreground=True)
        except Exception as e:
            self.log(f"[搜索框查找失败] {e}")
            # 备选方案：尝试查找所有Edit控件并使用第一个
            try:
                edits = main_win.descendants(control_type="Edit")
                edit_info = [f"{e.window_text()} class={e.element_info.class_name} auto_id={e.element_info.automation_id} handle={e.handle}" for e in edits]
                self.log(f"[调试] 所有Edit控件: {edit_info}")
                
                # 优先查找mmui::XLineEdit类名的控件
                xline_edits = [e for e in edits if e.element_info.class_name == "mmui::XLineEdit"]
                if xline_edits:
                    search_box = xline_edits[0]
                    self.log(f"[备选方案] 使用第一个mmui::XLineEdit控件: {search_box.element_info.class_name}")
                else:
                    # 如果没有找到指定类名的控件，使用第一个Edit控件
                    search_box = edits[0]
                    self.log(f"[备选方案] 使用第一个Edit控件: {search_box.element_info.class_name}")
                
                search_box.set_focus()
                search_box.type_keys('^a{BACKSPACE}', set_foreground=True)
                pyperclip.copy(contact)
                search_box.type_keys('^v{ENTER}', set_foreground=True)
            except Exception as e2:
                self.log(f"[备选方案失败] {e2}")
                raise RuntimeError("未找到搜索框控件")
        
        # 等待联系人加载
        time.sleep(1)
        
        # 检查聊天窗口标题
        try:
            chat_title = main_win.child_window(control_type="Text", found_index=0).window_text()
            self.log(f"[当前聊天窗口] 标题: {chat_title}")
            if contact not in chat_title:
                self.log(f"[警告] 未检测到目标联系人'{contact}'在聊天标题'{chat_title}'中")
        except Exception as e:
            self.log(f"[标题检查失败] {e}")
            texts = main_win.descendants(control_type="Text")
            text_info = [f"{t.window_text()} class={t.element_info.class_name} auto_id={t.element_info.automation_id} handle={t.handle}" for t in texts]
            self.log(f"[调试] 所有Text控件: {text_info}")
            self.log("[继续] 即使未确认聊天窗口标题也继续尝试发送消息")
        
        # 输入消息
        try:
            # 查找所有Edit控件
            edits = main_win.descendants(control_type="Edit")
            edit_info = [f"{e.window_text()} class={e.element_info.class_name} auto_id={e.element_info.automation_id} handle={e.handle}" for e in edits]
            self.log(f"[调试] 所有Edit控件: {edit_info}")
            
            # 尝试找到消息输入框
            # 通常消息输入框是除搜索框外的另一个Edit控件，可能是第二个
            input_box = None
            for edit in edits:
                # 排除搜索框（通常是mmui::XLineEdit）
                if edit.element_info.class_name != "mmui::XLineEdit":
                    input_box = edit
                    break
            
            if not input_box and len(edits) > 1:
                # 如果上面的条件无法找到，就使用第二个Edit控件
                input_box = edits[1]
            
            if not input_box:
                # 如果只有一个Edit控件，使用第一个
                input_box = edits[0]
            
            self.log(f"[选定输入框] class={input_box.element_info.class_name}, handle={input_box.handle}")
            input_box.set_focus()
            time.sleep(0.5)
            pyperclip.copy(message)
            input_box.type_keys('^v', set_foreground=True)
            time.sleep(0.5)
            input_box.type_keys('{ENTER}', set_foreground=True)
        except Exception as e:
            self.log(f"[消息发送失败] {e}")
            raise RuntimeError("未找到消息输入框控件或发送失败")
        
        self.log(f"[成功] 已发送消息给 {contact}")

    def _send_message_mac(self, win, contact, message):
        win.activate()
        pyautogui.hotkey('command', 'f')
        pyperclip.copy(contact)
        pyautogui.hotkey('command', 'v')
        pyautogui.press('enter')
        pyautogui.hotkey('command', 'l')  # 聚焦输入框（如支持）
        pyperclip.copy(message)
        pyautogui.hotkey('command', 'v')
        pyautogui.press('enter')
        self.log(f"[成功] 已发送消息给 {contact}") 