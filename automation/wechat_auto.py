import platform
import pyautogui
import pygetwindow as gw
import pyperclip

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
        self.window_titles = ["微信", "WeChat", "企业微信"]

    def log(self, msg):
        if self.logger:
            self.logger(msg)

    def focus_wechat_window(self):
        for title in self.window_titles:
            wins = gw.getWindowsWithTitle(title)
            if wins:
                win = wins[0]
                win.activate()
                self.log(f"已激活窗口: {win.title}")
                return win
        raise RuntimeError("未找到微信/企业微信窗口")

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
        app = Application(backend="uia").connect(title=win.title)
        main_win = app.window(title=win.title)
        main_win.set_focus()
        # 搜索联系人
        try:
            search_box = main_win.child_window(title="搜索", control_type="Edit")
            search_box.set_focus()
            search_box.type_keys('^a{BACKSPACE}', set_foreground=True)
            pyperclip.copy(contact)
            search_box.type_keys('^v{ENTER}', set_foreground=True)
        except Exception:
            raise RuntimeError("未找到搜索框控件")
        # 检查聊天窗口标题
        try:
            chat_title = main_win.child_window(control_type="Text", found_index=0).window_text()
            if contact not in chat_title:
                raise RuntimeError(f"未切换到目标联系人: {contact}")
        except Exception:
            raise RuntimeError("未找到聊天窗口标题控件")
        # 输入消息
        try:
            input_box = main_win.child_window(control_type="Edit", found_index=1)
            input_box.set_focus()
            pyperclip.copy(message)
            input_box.type_keys('^v', set_foreground=True)
            input_box.type_keys('{ENTER}', set_foreground=True)
        except Exception:
            raise RuntimeError("未找到消息输入框控件")
        self.log(f"[成功] 已发送消息给 {contact}")

    def _send_message_mac(self, win, contact, message):
        # macOS 下只能用快捷键和 pyautogui，控件名不可用
        # 仅用窗口激活+快捷键流程，找不到控件直接报错
        win.activate()
        pyautogui.hotkey('command', 'f')
        pyperclip.copy(contact)
        pyautogui.hotkey('command', 'v')
        pyautogui.press('enter')
        # 理论上此时已切换到联系人
        pyautogui.hotkey('command', 'l')  # 聚焦输入框（如支持）
        pyperclip.copy(message)
        pyautogui.hotkey('command', 'v')
        pyautogui.press('enter')
        self.log(f"[成功] 已发送消息给 {contact}") 