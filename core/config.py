"""配置管理模块"""
from pathlib import Path
import configparser
from core.logger import Logger


class ConfigManager:
    def __init__(self, config_file="ppt_monitor.ini"):
        self.config_file = Path(config_file)
        self.config = configparser.ConfigParser()
        self.logger = Logger()
        self.load_config()
    
    def load_config(self):
        if self.config_file.exists():
            encodings = ['utf-8-sig', 'utf-8', 'gbk', 'gb2312']
            content = None
            for encoding in encodings:
                try:
                    with open(self.config_file, 'r', encoding=encoding) as f:
                        content = f.read()
                        break
                except UnicodeDecodeError:
                    continue
            if content is not None:
                try:
                    self.config.read_string(content)
                    self.logger.info(f"成功加载配置文件: {self.config_file}")
                except configparser.Error as e:
                    self.logger.error(f"解析配置文件失败: {e}")
                    self.create_default_config()
            else:
                self.logger.warning(f"无法读取配置文件，使用默认配置")
                self.create_default_config()
        else:
            self.logger.info(f"配置文件不存在，创建默认配置")
            self.create_default_config()
    
    def create_default_config(self):
        self.config['General'] = {
            'backup_dir': './',
            'max_retention_days': '30',
            'enable_fallback_monitor': 'false',
            'min_file_size_kb': '10',
            'scan_interval_seconds': '30',
            'log_non_removable_events': 'false'
        }
        self.config['Logging'] = {
            'max_log_files': '5',
            'max_log_size_mb': '10'
        }
        self.save_config()
        self.logger.info("已创建默认配置文件")
    
    def save_config(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8-sig') as configfile:
                self.config.write(configfile)
        except Exception as e:
            self.logger.error(f"保存配置文件失败: {e}")
    
    def get_backup_dir(self): return self.config.get('General', 'backup_dir', fallback='./')
    def get_max_retention_days(self): 
        try: return int(self.config.get('General', 'max_retention_days', fallback='30'))
        except ValueError: return 30
    def get_enable_fallback_monitor(self): return self.config.getboolean('General', 'enable_fallback_monitor', fallback=False)
    def get_min_file_size_kb(self):
        try: return int(self.config.get('General', 'min_file_size_kb', fallback='10'))
        except ValueError: return 10
    def get_scan_interval(self):
        try: return int(self.config.get('General', 'scan_interval_seconds', fallback='30'))
        except ValueError: return 30
    def get_log_non_removable_events(self): return self.config.getboolean('General', 'log_non_removable_events', fallback=False)
    
    def set_backup_dir(self, backup_dir): 
        self.config['General']['backup_dir'] = backup_dir
        self.save_config()
    def set_max_retention_days(self, days): 
        self.config['General']['max_retention_days'] = str(days)
        self.save_config()
    def set_min_file_size_kb(self, size_kb): 
        self.config['General']['min_file_size_kb'] = str(size_kb)
        self.save_config()
    def set_enable_fallback(self, enabled): 
        self.config['General']['enable_fallback_monitor'] = str(enabled)
        self.save_config()
    def set_scan_interval(self, seconds): 
        self.config['General']['scan_interval_seconds'] = str(seconds)
        self.save_config()