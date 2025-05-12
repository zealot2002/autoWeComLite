#!/usr/bin/env python
"""
微信控件配置生成工具

用于分析微信窗口并生成控件配置文件，支持不同版本的微信。
使用方法:
    1. 打开微信
    2. 运行此脚本: python generate_controls_config.py
    3. 按照提示操作，程序会生成控件配置文件

生成的配置文件会保存在 config/wechat_controls.json
"""

import os
import sys
import time
import json
import platform

# 添加项目根目录到PATH
parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.append(parent_dir)

# 导入项目模块
from core.config_manager import ConfigManager
from automation.wechat_auto import WeChatAutomation


class ConfigGenerator:
    """微信控件配置生成器"""
    
    def __init__(self):
        self.config_manager = ConfigManager()
        self.is_windows = platform.system() == "Windows"
        self.is_mac = platform.system() == "Darwin"
        
        # 初始化日志函数
        self.log_messages = []
    
    def log(self, msg):
        """记录日志"""
        print(msg)
        self.log_messages.append(msg)
    
    def save_logs(self, filename="config_generator_log.txt"):
        """保存日志到文件"""
        log_path = os.path.join(parent_dir, "logs", filename)
        os.makedirs(os.path.dirname(log_path), exist_ok=True)
        
        with open(log_path, "w", encoding="utf-8") as f:
            f.write("\n".join(self.log_messages))
        print(f"日志已保存到: {log_path}")
    
    def analyze_wechat_window(self):
        """分析微信窗口，识别关键控件"""
        if not self.is_windows:
            self.log("当前仅支持Windows平台")
            return {}
        
        self.log("=== 开始分析微信窗口 ===")
        
        try:
            # 初始化自动化类但不执行任何操作
            automation = WeChatAutomation(logger=self.log)
            
            # 尝试激活窗口
            win = automation.focus_wechat_window()
            if not win:
                self.log("无法找到或激活微信窗口，请确保微信已打开")
                return {}
            
            self.log(f"成功激活窗口: {win.title}")
            
            # 导入pywinauto
            from pywinauto.application import Application
            from pywinauto import Desktop
            
            app = Application(backend="uia").connect(title=win.title, timeout=5)
            main_win = app.window(title=win.title)
            
            # 收集所有控件信息
            controls_info = {}
            
            # 按控件类型分类
            self.log("正在枚举控件...")
            
            # 搜索框 (Edit控件)
            edits = main_win.descendants(control_type="Edit")
            self.log(f"找到 {len(edits)} 个Edit控件")
            
            search_box = None
            for edit in edits:
                class_name = edit.element_info.class_name
                # 查找最可能的搜索框
                if "XLineEdit" in class_name:
                    search_box = edit
                    controls_info["search_box"] = {
                        "control_type": "Edit",
                        "class_name": class_name,
                        "description": "搜索框控件 (自动识别)"
                    }
                    self.log(f"找到可能的搜索框: class={class_name}")
                    break
            
            # 如果没找到明确的搜索框，使用第一个Edit控件
            if "search_box" not in controls_info and edits:
                controls_info["search_box"] = {
                    "control_type": "Edit",
                    "class_name": edits[0].element_info.class_name,
                    "description": "搜索框控件 (猜测)"
                }
                self.log(f"猜测搜索框: class={edits[0].element_info.class_name}")
            
            # 尝试识别消息输入框 (如果有多个Edit控件，通常第二个是消息输入框)
            if len(edits) > 1:
                input_box = edits[1]  # 假设第二个是消息输入框
                controls_info["message_input"] = {
                    "control_type": "Edit",
                    "class_name": input_box.element_info.class_name,
                    "description": "消息输入框控件 (猜测)"
                }
                self.log(f"猜测消息输入框: class={input_box.element_info.class_name}")
            
            # 列表控件 (搜索结果可能在List中)
            lists = main_win.descendants(control_type="List")
            self.log(f"找到 {len(lists)} 个List控件")
            if lists:
                controls_info["search_result_list"] = {
                    "control_type": "List",
                    "class_name": lists[0].element_info.class_name,
                    "description": "搜索结果列表控件 (猜测)"
                }
                self.log(f"猜测搜索结果列表: class={lists[0].element_info.class_name}")
            
            # 列表项控件 (搜索结果项可能是ListItem)
            list_items = main_win.descendants(control_type="ListItem")
            self.log(f"找到 {len(list_items)} 个ListItem控件")
            if list_items:
                controls_info["search_result_item"] = {
                    "control_type": "ListItem",
                    "class_name": list_items[0].element_info.class_name,
                    "description": "搜索结果列表项控件 (猜测)"
                }
                self.log(f"猜测搜索结果列表项: class={list_items[0].element_info.class_name}")
            
            # 聊天窗口标题 (通常是Text控件)
            texts = main_win.descendants(control_type="Text")
            self.log(f"找到 {len(texts)} 个Text控件")
            if texts:
                controls_info["chat_title"] = {
                    "control_type": "Text",
                    "class_name": texts[0].element_info.class_name,
                    "description": "聊天窗口标题控件 (猜测)"
                }
                self.log(f"猜测聊天窗口标题: class={texts[0].element_info.class_name}")
            
            self.log("=== 分析完成 ===")
            return controls_info
        
        except Exception as e:
            self.log(f"分析过程出错: {e}")
            import traceback
            self.log(traceback.format_exc())
            return {}
    
    def generate_config(self):
        """生成配置文件"""
        self.log("开始生成配置文件...")
        
        controls_info = self.analyze_wechat_window()
        if not controls_info:
            self.log("无法分析微信窗口，请确保微信已打开")
            return False
        
        # 加载现有配置
        config = self.config_manager.config
        
        # 更新Windows平台控件配置
        if "windows" not in config:
            config["windows"] = {}
        
        for control_name, info in controls_info.items():
            config["windows"][control_name] = info
        
        # 确保有默认的超时和策略配置
        if "timeouts" not in config:
            config["timeouts"] = {
                "search_result_wait": 1.5,
                "chat_window_load": 1.5,
                "input_focus": 0.5,
                "typing_pause": 0.3,
                "description": "各操作等待时间(秒)"
            }
        
        if "strategies" not in config:
            config["strategies"] = {
                "search_result_selection": "enter_key",
                "alternative_search_result_selection": "click_first_item",
                "description": "可选值: enter_key, click_first_item, click_matching_item"
            }
        
        # 保存配置
        self.config_manager.save_config(config)
        self.log(f"配置文件已生成: {self.config_manager.config_path}")
        
        self.save_logs()
        return True


def main():
    print("===== 微信控件配置生成工具 =====")
    print("此工具将分析当前打开的微信窗口，并生成控件配置文件")
    print("请确保微信已打开并位于前台")
    
    input("按Enter键继续...")
    
    generator = ConfigGenerator()
    generator.generate_config()
    
    print("\n配置生成完成！")
    print("您可以手动编辑配置文件以调整识别参数")
    print(f"配置文件路径: {os.path.join(parent_dir, 'config', 'wechat_controls.json')}")


if __name__ == "__main__":
    main() 