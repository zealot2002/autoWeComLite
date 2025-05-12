import json
import os
import platform

class ConfigManager:
    """配置管理器，负责加载和管理自动化控件配置"""
    
    def __init__(self, config_path=None):
        """初始化配置管理器
        
        Args:
            config_path: 配置文件路径，默认为 config/wechat_controls.json
        """
        self.is_windows = platform.system() == "Windows"
        self.is_mac = platform.system() == "Darwin"
        
        # 默认配置文件路径
        if config_path is None:
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            config_path = os.path.join(base_dir, "config", "wechat_controls.json")
        
        self.config_path = config_path
        self.config = self._load_config()
    
    def _load_config(self):
        """加载配置文件
        
        Returns:
            dict: 配置数据
        """
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            else:
                # 返回默认配置
                return self._get_default_config()
        except Exception as e:
            print(f"[警告] 加载配置文件失败: {e}，使用默认配置")
            return self._get_default_config()
    
    def _get_default_config(self):
        """获取默认配置
        
        Returns:
            dict: 默认配置数据
        """
        return {
            "windows": {
                "search_box": {
                    "control_type": "Edit",
                    "class_name": "mmui::XLineEdit",
                    "description": "搜索框控件"
                },
                "message_input": {
                    "control_type": "Edit",
                    "class_name": "",
                    "description": "消息输入框控件，留空表示使用排除法查找"
                }
            },
            "mac": {
                "search_shortcut": "command+f",
                "message_input_shortcut": "command+l"
            },
            "timeouts": {
                "search_result_wait": 1.5,
                "chat_window_load": 1.5,
                "input_focus": 0.5,
                "typing_pause": 0.3
            },
            "strategies": {
                "search_result_selection": "enter_key",
                "alternative_search_result_selection": "click_first_item"
            }
        }
    
    def save_config(self, config=None):
        """保存配置到文件
        
        Args:
            config: 要保存的配置，默认为当前配置
        """
        if config is None:
            config = self.config
        
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            print(f"[成功] 配置已保存到 {self.config_path}")
        except Exception as e:
            print(f"[错误] 保存配置失败: {e}")
    
    def get_control_config(self, control_name):
        """获取控件配置
        
        Args:
            control_name: 控件名称
        
        Returns:
            dict: 控件配置
        """
        if self.is_windows:
            return self.config.get("windows", {}).get(control_name, {})
        elif self.is_mac:
            return self.config.get("mac", {}).get(control_name, {})
        return {}
    
    def get_timeout(self, action_name):
        """获取超时设置
        
        Args:
            action_name: 动作名称
        
        Returns:
            float: 超时时间(秒)
        """
        return self.config.get("timeouts", {}).get(action_name, 1.0)
    
    def get_strategy(self, strategy_name):
        """获取策略设置
        
        Args:
            strategy_name: 策略名称
        
        Returns:
            str: 策略值
        """
        return self.config.get("strategies", {}).get(strategy_name, "")
    
    def update_control_class(self, control_name, class_name):
        """更新控件类名配置
        
        Args:
            control_name: 控件名称
            class_name: 类名
        """
        if self.is_windows:
            if "windows" in self.config and control_name in self.config["windows"]:
                self.config["windows"][control_name]["class_name"] = class_name
                self.save_config()
                return True
        elif self.is_mac:
            # Mac平台不使用类名
            pass
        return False
    
    def generate_config_from_controls(self, controls_info):
        """从控件信息生成配置
        
        Args:
            controls_info: 控件信息字典，格式为 {控件名: {control_type: xx, class_name: xx}}
        """
        if self.is_windows:
            for control_name, info in controls_info.items():
                if "windows" not in self.config:
                    self.config["windows"] = {}
                if control_name not in self.config["windows"]:
                    self.config["windows"][control_name] = {}
                
                self.config["windows"][control_name]["control_type"] = info.get("control_type", "")
                self.config["windows"][control_name]["class_name"] = info.get("class_name", "")
            
            self.save_config()
            return True
        return False 