import json
import os


class ConfigManager:
    """配置管理器，负责读取和保存配置到config.json"""
    CONFIG_FILE = "config.json"
    
    def __init__(self):
        self.config = self.load_config()
        if not os.path.exists(self.CONFIG_FILE):
            self.save_config()
    
    @classmethod
    def load_config(cls):
        """加载配置文件"""
        default_config = {
            # 发件人配置
            "sender_email": "ybjs3232678@126.com",
            "sender_password": "<PASSWORD>",
            "smtp_server": "smtp.126.com",
            "smtp_port": 25,
            
            # 文件路径配置
            "email_excel_path": "邮件对应表.xlsx",

            # 其他配置
            'sender_name': "四平市医保局结算科",
            'email_subject': "支付明细",
            'email_body': "请查收"
        }
        
        if os.path.exists(cls.CONFIG_FILE):
            try:
                with open(cls.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                # 合并默认配置，确保所有字段都存在
                for key, value in default_config.items():
                    if key not in config:
                        config[key] = value
                return config
            except Exception as e:
                print(f"读取配置文件失败：{e}")
                return default_config
        else:
            return default_config
    
    def get_config(self, key=None):
        """获取配置项"""
        if key is None:
            return self.config
        return self.config.get(key)
    
    def update_config(self, **kwargs):
        """更新配置项"""
        for key, value in kwargs.items():
            if key in self.config:
                self.config[key] = value
            else:
                print(f"警告：未知配置项 {key}")
        return self.save_config()
    
    def save_config(self):
        """保存配置到文件"""
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True, "配置保存成功"
        except Exception as e:
            return False, f"配置保存失败：{e}"
    
    def reset_to_default(self):
        """重置为默认配置"""
        self.config = self.load_config()
        return self.save_config()
    
    def get_email_sender_config(self):
        """获取EmailSender初始化所需的配置"""
        return {
            'sender_email': self.config['sender_email'],
            'sender_password': self.config['sender_password'],
            'sender_name': self.config['sender_name'],
            'smtp_server': self.config['smtp_server'],
            'smtp_port': self.config['smtp_port']
        }
    