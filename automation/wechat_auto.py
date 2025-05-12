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
        self.log("[配置] 已加载以下配置:")
        self.log(f"[配置] Windows控件配置: {json.dumps(self.control_configs, ensure_ascii=False, indent=2)}")
        self.log(f"[配置] 等待时间配置: {json.dumps(self.timeouts, ensure_ascii=False, indent=2)}")
        self.log(f"[配置] 策略配置: {json.dumps(self.strategies, ensure_ascii=False, indent=2)}")

    def log(self, msg):
        if self.logger:
            self.logger(msg)

    def focus_wechat_window(self):
        """直接通过配置的类名查找并激活微信窗口"""
        self.log("[窗口查找] 开始查找微信窗口")
        
        if not self.is_win or not self.pywinauto:
            raise RuntimeError("仅支持Windows平台")
        
        # 输出系统环境信息
        self.log(f"[系统信息] 操作系统: {platform.system()} {platform.release()}, Python版本: {platform.python_version()}")
        
        # 获取配置的微信窗口类名
        main_window_config = self.control_configs.get("main_window", {})
        wechat_class_name = main_window_config.get("class_name", "mmui::MainWindow")
        
        self.log(f"[窗口查找] 使用类名 '{wechat_class_name}' 查找微信窗口")
        
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
        
        # 直接查找特定类名的窗口
        try:
            desktop = Desktop(backend="uia")
            wechat_windows = []
            other_wechat_windows = []
            
            # 查找所有匹配类名的窗口
            for win in desktop.windows():
                try:
                    title = win.window_text()
                    class_name = win.element_info.class_name
                    handle = win.handle
                    
                    # 完全匹配的优先
                    if wechat_class_name == class_name:
                        self.log(f"[微信窗口] 精确匹配: '{title}', class='{class_name}', handle={handle}")
                        wechat_windows.append((title, class_name, handle))
                    # 部分匹配作为备选
                    elif wechat_class_name in class_name:
                        self.log(f"[微信窗口] 部分匹配: '{title}', class='{class_name}', handle={handle}")
                        wechat_windows.append((title, class_name, handle))
                    # 其他可能是微信的窗口
                    elif any(key.lower() in title.lower() for key in self.window_keywords):
                        self.log(f"[其他相关窗口] 标题匹配: '{title}', class='{class_name}', handle={handle}")
                        other_wechat_windows.append((title, class_name, handle))
                except Exception:
                    pass
            
            if not wechat_windows:
                if other_wechat_windows:
                    self.log(f"[提示] 未找到类名为 '{wechat_class_name}' 的窗口，但找到以下可能的微信相关窗口:")
                    for title, class_name, handle in other_wechat_windows:
                        self.log(f"  候选窗口: '{title}', class='{class_name}'")
                    self.log(f"[建议] 请修改配置文件中的 'main_window.class_name' 值为上述窗口的class_name")
                
                self.log(f"[错误] 未找到类名为 '{wechat_class_name}' 的微信窗口")
                raise RuntimeError(f"未找到微信窗口，请确保微信已打开或更新配置文件中的类名")
            
            # 如果找到多个，优先选择第一个
            title, class_name, handle = wechat_windows[0]
            self.log(f"[选择] 将激活窗口: '{title}', class='{class_name}', handle={handle}")
            
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
        
        # 搜索框详细情况
        self.log("[搜索框枚举] ===== 开始查找所有Edit控件 =====")
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
        self.log("[搜索框枚举] ===== 枚举完毕 =====")
        
        # 搜索联系人
        try:
            # 获取搜索框配置
            search_box_config = self.control_configs.get("search_box", {})
            search_box_class = search_box_config.get("class_name", "mmui::XLineEdit")
            search_box_type = search_box_config.get("control_type", "Edit")
            
            # 精确定位搜索框
            self.log(f"[搜索框] 使用类名='{search_box_class}', 类型='{search_box_type}'查找搜索框")
            search_box = main_win.child_window(class_name=search_box_class, control_type=search_box_type)
            self.log(f"[搜索框] 精确定位 class={search_box.element_info.class_name}, handle={search_box.handle}")
            
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
            
            # 直接按回车选择第一个结果
            self.log(f"[搜索框] 按回车选择联系人")
            search_box.type_keys('{ENTER}', set_foreground=True)
            
            # 等待聊天窗口加载
            chat_window_wait = self.timeouts["chat_window_load"]
            self.log(f"[搜索框] 等待聊天窗口加载 (等待 {chat_window_wait} 秒)")
            time.sleep(chat_window_wait)
        except Exception as e:
            # 严格错误处理：输出更详细的错误信息
            self.log(f"[搜索框查找失败] {e}")
            try:
                # 检查所有Edit控件，给出更明确的建议
                edits = main_win.descendants(control_type="Edit")
                if edits:
                    self.log(f"[诊断信息] 找到 {len(edits)} 个Edit控件，但没有匹配的搜索框:")
                    for i, edit in enumerate(edits):
                        try:
                            class_name = edit.element_info.class_name
                            text = edit.window_text()
                            rect = edit.rectangle()
                            self.log(f"  Edit[{i}]: class='{class_name}', text='{text}', rect={rect}")
                        except:
                            pass
                    
                    # 建议使用第一个Edit控件作为搜索框
                    if edits and edits[0].element_info.class_name != search_box_class:
                        suggest_class = edits[0].element_info.class_name
                        self.log(f"[建议] 请修改配置文件: 搜索框class_name应改为 '{suggest_class}'")
                else:
                    self.log("[诊断信息] 未找到任何Edit控件，请检查窗口状态")
            except Exception as e2:
                self.log(f"[调试信息获取失败] {e2}")
            
            self.log(f"[错误] 无法精确定位搜索框(class_name=\"{search_box_class}\")，操作终止")
            raise RuntimeError(f"无法精确定位搜索框，请确认微信版本兼容性或更新配置文件中的class_name")
        
        # 输入消息
        try:
            # 获取消息输入框配置
            msg_input_config = self.control_configs.get("message_input", {})
            msg_input_class = msg_input_config.get("class_name", "")
            msg_input_type = msg_input_config.get("control_type", "Edit")
            
            self.log(f"[消息框] 使用类名='{msg_input_class}', 类型='{msg_input_type}'查找消息输入框")
            
            # 再次枚举所有Edit控件，用于诊断
            chat_edits = main_win.descendants(control_type="Edit")
            self.log(f"[聊天界面] 找到 {len(chat_edits)} 个Edit控件:")
            for i, edit in enumerate(chat_edits):
                try:
                    class_name = edit.element_info.class_name
                    text = edit.window_text()
                    self.log(f"  ChatEdit[{i}]: class='{class_name}', text='{text}'")
                except:
                    pass
            
            # 找到输入框
            input_box = None
            
            # 方法1: 使用配置的类名精确查找
            if msg_input_class:
                try:
                    input_box = main_win.child_window(class_name=msg_input_class, control_type=msg_input_type)
                    self.log(f"[消息框] 通过class_name='{msg_input_class}'找到消息输入框")
                except Exception as e:
                    self.log(f"[消息框] 通过配置的class_name查找失败: {e}")
                    self.log("[消息框] 尝试备选查找方法...")
            
            # 方法2: 如果方法1失败，尝试排除法查找
            if not input_box and len(chat_edits) > 0:
                # 排除搜索框，找到其他Edit控件
                for edit in chat_edits:
                    try:
                        if edit.element_info.class_name != search_box_class:
                            input_box = edit
                            self.log(f"[消息框] 通过排除法找到消息输入框: class='{edit.element_info.class_name}'")
                            
                            # 建议更新配置
                            if edit.element_info.class_name != msg_input_class:
                                self.log(f"[建议] 更新配置文件: 消息输入框class_name应为 '{edit.element_info.class_name}'")
                            break
                    except:
                        pass
            
            # 方法3: 如果还是没找到，使用最简单的方法 - 假设第二个Edit控件是消息框
            if not input_box and len(chat_edits) > 1:
                input_box = chat_edits[1]
                self.log(f"[消息框] 使用第二个Edit控件作为消息输入框: class='{input_box.element_info.class_name}'")
                
                # 建议更新配置
                if input_box.element_info.class_name != msg_input_class:
                    self.log(f"[建议] 更新配置文件: 消息输入框class_name应为 '{input_box.element_info.class_name}'")
            
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
                self.log("[错误] 无法找到消息输入框，请检查控件配置")
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