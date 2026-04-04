"""持久化文件管理模块"""
import json
import threading
import time
from pathlib import Path
from datetime import datetime
from core.logger import Logger
from utils.constants import STATE_SAVE_INTERVAL


class PersistentFileManager:
    """持久化文件管理器 - 延迟保存状态"""
    def __init__(self, base_backup_dir):
        self.base_backup_dir = Path(base_backup_dir)
        self.state_file = self.base_backup_dir / 'monitor_state.json'
        self.today = datetime.now().strftime("%Y-%m-%d")
        self.processed_files = {}
        self.lock = threading.Lock()
        self.logger = Logger()
        self._dirty = False
        
        self.load_state()
        self._start_auto_save_timer()
    
    def _start_auto_save_timer(self):
        def auto_save_loop():
            while True:
                time.sleep(STATE_SAVE_INTERVAL)
                if self._dirty:
                    self._do_save()
        threading.Thread(target=auto_save_loop, daemon=True).start()
    
    def _do_save(self):
        with self.lock:
            if not self._dirty:
                return
            data = {'date': self.today, 'processed_files': self.processed_files}
            try:
                with open(self.state_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self._dirty = False
            except Exception as e:
                self.logger.error(f"保存状态文件失败: {str(e)}")
    
    def load_state(self):
        if self.state_file.exists():
            try:
                with open(self.state_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if data.get('date') == self.today:
                        self.processed_files = data.get('processed_files', {})
                        self.logger.info(f"加载状态成功，今日已处理 {len(self.processed_files)} 个文件")
            except (json.JSONDecodeError, KeyError) as e:
                self.logger.error(f"加载状态文件失败: {e}")
    
    def save_state_immediately(self): self._do_save()
    
    def add_processed_file(self, file_path, mtime):
        with self.lock:
            self.processed_files[str(file_path)] = mtime
            self._dirty = True
    
    def is_already_processed(self, file_path):
        with self.lock:
            return str(file_path) in self.processed_files
    
    def get_file_mtime(self, file_path):
        with self.lock:
            return self.processed_files.get(str(file_path))
    
    def cleanup_old_state(self):
        current_date = datetime.now().strftime("%Y-%m-%d")
        if current_date != self.today:
            with self.lock:
                self.processed_files = {}
                self._dirty = True
            self.today = current_date
    
    def get_processed_count(self):
        with self.lock:
            return len(self.processed_files)