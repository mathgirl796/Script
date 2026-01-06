import json
import os


class ConfigManager:
    """配置管理器，负责读取和保存配置到config.json"""
    CONFIG_FILE = "config.json"
    
    def __init__(self):
        config, message = self.load_config()
        if config is None:
            raise ValueError(f"配置加载失败：{message}")
        self.config = config
    
    @classmethod
    def load_config(cls):
        """加载配置文件"""
        if not os.path.exists(cls.CONFIG_FILE):
            return None, f"配置文件 {cls.CONFIG_FILE} 不存在"
        
        try:
            with open(cls.CONFIG_FILE, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            # 检查必需的配置项
            required_keys = [
                'sender_email', 'sender_password', 'smtp_server', 'smtp_port',
                'email_excel_path', 'sender_name', 'email_subject', 'email_body'
            ]
            
            missing_keys = []
            for key in required_keys:
                if key not in config:
                    missing_keys.append(key)
            
            if missing_keys:
                return None, f"配置文件缺少必需的配置项：{', '.join(missing_keys)}"
            
            # 检查配置项的值是否有效
            validation_errors = []
            if not config['sender_email']:
                validation_errors.append("发件人邮箱不能为空")
            if not config['sender_password']:
                validation_errors.append("授权码不能为空")
            if not config['smtp_server']:
                validation_errors.append("SMTP服务器不能为空")
            if not isinstance(config['smtp_port'], int) or config['smtp_port'] <= 0:
                validation_errors.append("SMTP端口必须是正整数")
            
            if validation_errors:
                return None, "配置验证失败：" + "; ".join(validation_errors)
            
            return config, "配置加载成功"
            
        except json.JSONDecodeError as e:
            return None, f"配置文件格式错误（JSON解析失败）：{str(e)}"
        except UnicodeDecodeError as e:
            return None, f"配置文件编码错误：{str(e)}"
        except Exception as e:
            return None, f"读取配置文件时发生未知错误：{str(e)}"
    
    def get_config(self, key=None):
        """获取配置项"""
        if key is None:
            return self.config
        return self.config.get(key)
    
    def update_config(self, **kwargs):
        """更新配置项"""
        # 允许添加新的配置项，但要验证基本格式
        for key, value in kwargs.items():
            if key == 'smtp_port':
                if isinstance(value, str) and value.isdigit():
                    self.config[key] = int(value)
                elif isinstance(value, int):
                    self.config[key] = value
                else:
                    return False, f"SMTP端口必须是数字，当前值：{value}"
            else:
                self.config[key] = str(value) if value is not None else ""
        
        return self.save_config()
    
    def save_config(self):
        """保存配置到文件"""
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
            return True, "配置保存成功"
        except Exception as e:
            return False, f"配置保存失败：{e}"
    
    def create_default_config(self):
        """创建默认配置文件"""
        default_config = {
            # 发件人配置
            "sender_email": "",
            "sender_password": "",
            "smtp_server": "",
            "smtp_port": 25,
            
            # 文件路径配置
            "email_excel_path": "",

            # 其他配置
            'sender_name': "",
            'email_subject': "支付明细",
            'email_body': "请查收"
        }
        
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=2)
            return True, "默认配置文件创建成功，请填写配置项"
        except Exception as e:
            return False, f"创建默认配置文件失败：{str(e)}"
    
    def get_email_sender_config(self):
        """获取EmailSender初始化所需的配置"""
        return {
            'sender_email': self.config['sender_email'],
            'sender_password': self.config['sender_password'],
            'sender_name': self.config['sender_name'],
            'smtp_server': self.config['smtp_server'],
            'smtp_port': self.config['smtp_port']
        }
    