import platform
import pyautogui
import pygetwindow as gw
import pyperclip
from pywinauto import Desktop, Application
import time
import win32gui
import win32con
from core.config_manager import ConfigManager
import json

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
        
        # 输出当前加载的配置信息
        self.log("[配置] ===== 加载的配置信息 =====")
        main_window_class = self.control_configs.get('main_window', {}).get('class_name', '未配置')
        self.log(f"[配置] 主窗口类名: {main_window_class}")
        self.log(f"[配置] 搜索框类名: {self.control_configs.get('search_box', {}).get('class_name', '未配置')}")
        self.log(f"[配置] 消息框类名: {self.control_configs.get('message_input', {}).get('class_name', '未配置')}")
        self.log(f"[配置] 搜索结果等待时间: {self.timeouts.get('search_result_wait', 0.5)} 秒")
        self.log(f"[配置] 聊天窗口加载等待时间: {self.timeouts.get('chat_window_load', 0.5)} 秒")

    def log(self, msg):
        if self.logger:
            self.logger(msg)

    def focus_wechat_window(self):
        """查找并激活微信窗口，考虑兼容性"""
        self.log("[窗口查找] 开始查找微信窗口")
        
        if not self.is_win or not self.pywinauto:
            raise RuntimeError("仅支持Windows平台")
        
        # 输出系统环境信息
        self.log(f"[系统信息] 操作系统: {platform.system()} {platform.release()}, Python版本: {platform.python_version()}")
        
        # 输出所有窗口信息，帮助诊断
        self.log("[窗口枚举] ===== 开始枚举所有可见窗口 =====")
        all_windows = []
        try:
            desktop = Desktop(backend="uia")
            for win in desktop.windows():
                try:
                    title = win.window_text()
                    class_name = win.element_info.class_name
                    handle = win.handle
                    process_id = win.element_info.process_id
                    # 只记录有标题的窗口，避免太多输出
                    if title:
                        win_info = {
                            "title": title,
                            "class_name": class_name,
                            "handle": handle,
                            "process_id": process_id
                        }
                        all_windows.append(win_info)
                        self.log(f"  窗口: title='{title}', class='{class_name}', handle={handle}, pid={process_id}")
                except Exception as e:
                    continue
        except Exception as e:
            self.log(f"[警告] 枚举窗口时出错: {e}")
        
        self.log(f"[窗口枚举] 找到 {len(all_windows)} 个窗口")
        self.log("[窗口枚举] ===== 结束窗口枚举 =====")
        
        # 查找微信窗口的多种方式
        wechat_windows = []
        
        # 方法1: 通过配置的类名查找
        main_window_config = self.control_configs.get("main_window", {})
        wechat_class_name = main_window_config.get("class_name", "")
        
        if wechat_class_name:
            self.log(f"[窗口查找] 方法1: 使用配置的类名 '{wechat_class_name}' 查找")
            try:
                desktop = Desktop(backend="uia")
                for win in desktop.windows():
                    try:
                        title = win.window_text()
                        class_name = win.element_info.class_name
                        handle = win.handle
                        
                        if class_name == wechat_class_name:  # 使用完全匹配而不是部分匹配
                            self.log(f"[微信窗口] 通过类名完全匹配: '{title}', class='{class_name}', handle={handle}")
                            wechat_windows.append((title, class_name, handle, 1))  # 优先级1
                        elif wechat_class_name in class_name:  # 部分匹配作为备选
                            self.log(f"[微信窗口] 通过类名部分匹配: '{title}', class='{class_name}', handle={handle}")
                            wechat_windows.append((title, class_name, handle, 2))  # 优先级2
                    except Exception:
                        pass
            except Exception as e:
                self.log(f"[警告] 通过类名查找失败: {e}")
        
        # 方法2: 通过窗口标题关键词查找
        if len(wechat_windows) == 0:
            self.log(f"[窗口查找] 方法2: 通过窗口标题关键词查找")
            try:
                desktop = Desktop(backend="uia")
                for win in desktop.windows():
                    try:
                        title = win.window_text()
                        class_name = win.element_info.class_name
                        handle = win.handle
                        
                        if any(key.lower() in title.lower() for key in self.window_keywords) and \
                           not any(ex.lower() in title.lower() for ex in self.exclude_keywords):
                            self.log(f"[微信窗口] 通过标题匹配: '{title}', class='{class_name}', handle={handle}")
                            wechat_windows.append((title, class_name, handle, 3))  # 优先级3
                    except Exception:
                        pass
            except Exception as e:
                self.log(f"[警告] 通过标题查找失败: {e}")
        
        # 如果未找到任何窗口，报错
        if not wechat_windows:
            self.log("[错误] 未找到微信窗口")
            raise RuntimeError(f"未找到微信窗口，请确保微信已打开")
        
        # 按优先级排序并选择第一个窗口
        wechat_windows.sort(key=lambda x: x[3])  # 按优先级排序
        title, class_name, handle, _ = wechat_windows[0]
        
        self.log(f"[选择] 将激活窗口: '{title}', class='{class_name}', handle={handle}")
        
        # 如果找到的窗口类名与配置不符，建议更新配置
        if wechat_class_name and class_name != wechat_class_name:
            self.log(f"[建议] 请更新配置文件中的微信窗口类名: main_window.class_name='{class_name}'")
        
        # 通过handle获取窗口对象
        win = None
        for w in gw.getAllWindows():
            if w._hWnd == handle:
                win = w
                break
        
        if not win:
            # 备选方案：使用标题查找
            self.log(f"[警告] 无法通过handle获取窗口对象，尝试使用标题查找")
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
        
        # 列出所有控件，帮助诊断
        self.log("[控件枚举] ===== 开始枚举窗口控件 =====")
        windows = Desktop(backend="uia").windows()
        for w in windows:
            if w.window_text() == win.title:
                self.log(f"[主窗口] title='{w.window_text()}' class='{w.element_info.class_name}' handle={w.handle}")
                # 枚举主窗口下的所有直接子控件
                children = w.children()
                self.log(f"[直接子控件] 数量: {len(children)}")
                for i, child in enumerate(children):
                    try:
                        ctrl_type = child.element_info.control_type
                        class_name = child.element_info.class_name
                        title = child.window_text()
                        self.log(f"  子控件[{i}]: type='{ctrl_type}', class='{class_name}', text='{title}'")
                    except:
                        pass
                break
        
        try:
            app = Application(backend="uia").connect(title=win.title, timeout=5)
        except Exception as e:
            self.log(f"[pywinauta连接异常] {e}")
            raise
            
        main_win = app.window(title=win.title)
        main_win.set_focus()
        
        # 列出所有Edit控件
        self.log("[编辑框枚举] ===== 开始查找所有Edit控件 =====")
        edits = main_win.descendants(control_type="Edit")
        self.log(f"[Edit控件] 找到 {len(edits)} 个Edit控件:")
        for i, edit in enumerate(edits):
            try:
                class_name = edit.element_info.class_name
                text = edit.window_text()
                rect = edit.rectangle()
                self.log(f"  Edit[{i}]: class='{class_name}', text='{text}', rect={rect}")
            except:
                pass
        self.log("[编辑框枚举] ===== 枚举完毕 =====")
        
        if len(edits) == 0:
            self.log("[错误] 未找到任何Edit控件，无法继续操作")
            raise RuntimeError("未找到任何编辑框控件，请检查微信窗口状态")
        
        # 搜索联系人
        try:
            # 查找搜索框的多种方式
            search_box = None
            
            # 方法1: 使用配置的类名查找
            search_box_config = self.control_configs.get("search_box", {})
            search_box_class = search_box_config.get("class_name", "")
            
            if search_box_class:
                self.log(f"[搜索框] 方法1: 使用配置的类名 '{search_box_class}' 查找")
                try:
                    search_box = main_win.child_window(class_name=search_box_class, control_type="Edit")
                    self.log(f"[搜索框] 通过类名找到: class='{search_box.element_info.class_name}'")
                except Exception as e:
                    self.log(f"[搜索框] 通过类名查找失败: {e}")
            
            # 方法2: 使用第一个Edit控件作为搜索框
            if not search_box and len(edits) > 0:
                self.log("[搜索框] 方法2: 使用第一个Edit控件作为搜索框")
                search_box = edits[0]
                self.log(f"[搜索框] 使用第一个Edit控件: class='{search_box.element_info.class_name}'")
                
                # 如果第一个Edit的类名与配置不符，建议更新配置
                if search_box.element_info.class_name != search_box_class:
                    self.log(f"[建议] 请更新配置文件: 搜索框class_name='{search_box.element_info.class_name}'")
            
            if not search_box:
                self.log("[错误] 未找到搜索框")
                raise RuntimeError("未找到搜索框，无法继续操作")
            
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
            
            # 搜索结果详细情况
            self.log("[搜索结果枚举] ===== 开始枚举搜索结果相关控件 =====")
            list_items = main_win.descendants(control_type="ListItem")
            self.log(f"[ListItem控件] 找到 {len(list_items)} 个ListItem控件:")
            for i, item in enumerate(list_items):
                try:
                    class_name = item.element_info.class_name
                    text = item.window_text()
                    self.log(f"  ListItem[{i}]: class='{class_name}', text='{text}'")
                    # 检查是否匹配联系人
                    if contact.lower() in text.lower():
                        self.log(f"  [匹配] ListItem[{i}] 匹配联系人 '{contact}'")
                except:
                    pass
            self.log("[搜索结果枚举] ===== 枚举完毕 =====")
            
            # 使用配置的策略选择搜索结果
            selection_strategy = self.strategies.get("search_result_selection", "enter_key")
            
            if selection_strategy == "enter_key" or len(list_items) == 0:
                # 直接按回车选择第一个结果
                self.log(f"[搜索框] 按回车选择联系人")
                search_box.type_keys('{ENTER}', set_foreground=True)
            elif selection_strategy == "click_first_item" and len(list_items) > 0:
                # 点击第一个列表项
                list_items[0].click_input()
                self.log(f"[搜索框] 点击第一个列表项选择联系人")
            elif selection_strategy == "click_matching_item" and len(list_items) > 0:
                # 查找匹配的列表项并点击
                matched_item = None
                for item in list_items:
                    try:
                        text = item.window_text()
                        if contact.lower() in text.lower():
                            matched_item = item
                            break
                    except:
                        pass
                
                if matched_item:
                    matched_item.click_input()
                    self.log(f"[搜索框] 点击匹配的列表项选择联系人")
                else:
                    # 如果没找到匹配项，回退到按回车
                    self.log(f"[搜索框] 未找到匹配项，回退到按回车")
                    search_box.type_keys('{ENTER}', set_foreground=True)
            
            # 等待聊天窗口加载
            chat_window_wait = self.timeouts["chat_window_load"]
            self.log(f"[搜索框] 等待聊天窗口加载 (等待 {chat_window_wait} 秒)")
            time.sleep(chat_window_wait)
            
            # 再次枚举Edit控件，此时应该包含消息输入框
            self.log("[聊天窗口] ===== 查找消息输入框 =====")
            chat_edits = main_win.descendants(control_type="Edit")
            self.log(f"[聊天窗口] 找到 {len(chat_edits)} 个Edit控件:")
            for i, edit in enumerate(chat_edits):
                try:
                    class_name = edit.element_info.class_name
                    text = edit.window_text()
                    self.log(f"  ChatEdit[{i}]: class='{class_name}', text='{text}'")
                except:
                    pass
            
            # 输入消息
            # 不再依赖类名查找，直接基于控件类型和序号查找
            input_box = None
            
            # 在联系人对话窗口中，第一个Edit控件是消息输入框
            if len(chat_edits) > 0:
                input_box = chat_edits[0]  # 使用第一个Edit控件作为消息输入框
                self.log(f"[消息框] 使用第一个Edit控件作为消息输入框: class='{input_box.element_info.class_name}'")
            else:
                self.log("[错误] 未找到任何可用作消息输入框的Edit控件")
                raise RuntimeError("未找到消息输入框")
            
            # 输入消息
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
                self.log("[错误] 无法找到消息输入框")
                raise RuntimeError("未找到消息输入框")
            
        except Exception as e:
            self.log(f"[错误] 发送消息失败: {e}")
            raise RuntimeError(f"发送消息失败: {e}")

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