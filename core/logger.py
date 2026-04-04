"""日志管理模块"""
import sys
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path
from datetime import datetime


class Logger:
    """日志管理器 - 单例模式"""
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialize()
        return cls._instance
    
    def _initialize(self):
        if getattr(sys, 'frozen', False):
            base_dir = Path(sys.executable).parent
        else:
            base_dir = Path(__file__).parent.parent
        
        self.log_dir = base_dir / 'logs'
        self.log_dir.mkdir(exist_ok=True)
        
        log_file = self.log_dir / f'ppt_monitor_{datetime.now().strftime("%Y%m%d")}.log'
        
        self.logger = logging.getLogger('PPTMonitor')
        self.logger.setLevel(logging.DEBUG)
        
        if self.logger.handlers:
            return
        
        file_handler = RotatingFileHandler(str(log_file), maxBytes=10*1024*1024, backupCount=5, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
    
    def debug(self, message): self.logger.debug(message)
    def info(self, message): self.logger.info(message)
    def warning(self, message): self.logger.warning(message)
    def error(self, message): self.logger.error(message)
    def exception(self, message): self.logger.exception(message)