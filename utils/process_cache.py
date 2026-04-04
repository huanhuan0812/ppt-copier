"""进程缓存模块"""
import threading
import time
import psutil
from core.logger import Logger


class ProcessCache:
    """进程扫描缓存 - 避免锁竞争和长时间阻塞"""
    
    def __init__(self, ttl_seconds=2):
        self.cache = {}
        self.ttl = ttl_seconds
        self.logger = Logger()
        self._lock = threading.RLock()
        self._scan_lock = threading.Lock()
        
        self.pp_process_ids = []
        self.pp_process_ids_cache_time = 0
        self._is_scanning = False
        self.process_files_cache = {}
        self.invalid_pids = set()
        self.invalid_pids_cleanup_time = 0
        self._shutting_down = False
    
    def set_shutting_down(self, value): self._shutting_down = value
    
    def get_powerpoint_process_ids(self):
        """获取PowerPoint进程ID列表（带缓存，非阻塞）"""
        if self._shutting_down:
            return []
        
        with self._lock:
            now = time.time()
            if now - self.pp_process_ids_cache_time < self.ttl:
                return self.pp_process_ids.copy()
        
        if self._is_scanning or not self._scan_lock.acquire(blocking=False):
            with self._lock:
                return self.pp_process_ids.copy()
        
        try:
            self._is_scanning = True
            pids = []
            for proc in psutil.process_iter(['pid', 'name']):
                if self._shutting_down:
                    break
                try:
                    if proc.info['name'] and proc.info['name'].lower().startswith('powerpnt'):
                        pids.append(proc.info['pid'])
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
            
            with self._lock:
                self.pp_process_ids = pids
                self.pp_process_ids_cache_time = time.time()
                return pids.copy()
        finally:
            self._is_scanning = False
            self._scan_lock.release()
    
    def get_process_open_files_by_pid(self, pid):
        """通过进程ID获取打开的文件列表（带缓存）"""
        if self._shutting_down:
            return []
        
        with self._lock:
            now = time.time()
            if pid in self.invalid_pids:
                if now - self.invalid_pids_cleanup_time > 30:
                    self.invalid_pids.clear()
                    self.invalid_pids_cleanup_time = now
                return []
            if pid in self.process_files_cache:
                cached_time, cached_files = self.process_files_cache[pid]
                if now - cached_time < self.ttl:
                    return cached_files
        
        try:
            proc = psutil.Process(pid)
            if not proc.is_running():
                with self._lock:
                    self.invalid_pids.add(pid)
                return []
            
            open_files = []
            for file_info in proc.open_files():
                if self._shutting_down:
                    break
                open_files.append(file_info.path)
            
            with self._lock:
                self.process_files_cache[pid] = (time.time(), open_files)
                if len(self.process_files_cache) > 100:
                    oldest_keys = sorted(self.process_files_cache.keys(), 
                                       key=lambda k: self.process_files_cache[k][0])[:50]
                    for key in oldest_keys:
                        del self.process_files_cache[key]
            return open_files
        except (psutil.NoSuchProcess, psutil.AccessDenied, Exception):
            with self._lock:
                self.invalid_pids.add(pid)
            return []
    
    def invalidate_process(self, pid=None):
        """异步使缓存失效"""
        def async_invalidate():
            with self._lock:
                if pid is not None:
                    self.process_files_cache.pop(pid, None)
                    self.invalid_pids.discard(pid)
                else:
                    self.process_files_cache.clear()
                    self.invalid_pids.clear()
        threading.Thread(target=async_invalidate, daemon=True).start()
    
    def clear_all(self):
        def async_clear():
            with self._lock:
                self.process_files_cache.clear()
                self.invalid_pids.clear()
                self.pp_process_ids = []
                self.pp_process_ids_cache_time = 0
        threading.Thread(target=async_clear, daemon=True).start()