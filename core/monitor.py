"""PPT监控核心模块"""
import os
import time
import shutil
import threading
from pathlib import Path
from datetime import datetime, timedelta
import win32file
import win32con
import pythoncom
from core.logger import Logger
from core.config import ConfigManager
from core.file_manager import PersistentFileManager
from utils.process_cache import ProcessCache
from utils.constants import PPT_EXTENSIONS
from events.device_events import WindowsDeviceMonitor
from events.com_monitor import PowerPointEventMonitor


class PPTMonitor:
    def __init__(self):
        self.logger = Logger()
        self.logger.info("初始化PPT监控器（事件驱动模式）")
        
        self.config_manager = ConfigManager()
        self.base_backup_dir = Path(self.config_manager.get_backup_dir())
        self.base_backup_dir.mkdir(exist_ok=True)
        self.max_retention_days = self.config_manager.get_max_retention_days()
        self.enable_fallback = self.config_manager.get_enable_fallback_monitor()
        self.min_file_size_bytes = self.config_manager.get_min_file_size_kb() * 1024
        self.scan_interval = self.config_manager.get_scan_interval()
        
        self.file_manager = PersistentFileManager(self.base_backup_dir)
        self.process_cache = ProcessCache(ttl_seconds=2)
        
        self.device_monitor = WindowsDeviceMonitor(self.on_device_event)
        self.ppt_com_monitor = PowerPointEventMonitor(self)
        
        self.current_removable_drives_cache = set()
        self.processing_lock = threading.Lock()
        self.currently_processing = set()
        self.running = True
        self.fallback_thread = None
    
    def on_device_event(self, event_type, drives):
        if event_type == 'device_inserted':
            self.current_removable_drives_cache.update(drives)
            self.logger.info(f"设备插入: {drives}")
            self.process_cache.invalidate_process()
        else:
            for drive in drives:
                self.current_removable_drives_cache.discard(drive)
            self.logger.info(f"设备移除: {drives}")
            self.process_cache.invalidate_process()
    
    def on_powerpoint_quit(self):
        self.logger.info("PowerPoint已正常退出")
    
    def is_removable_drive(self, path):
        if not path:
            return False
        try:
            drive_letter = str(Path(path).drive).upper().rstrip(':')
            return drive_letter in self.current_removable_drives_cache
        except:
            return False
    
    def is_valid_ppt_file_for_backup(self, file_path):
        if not file_path:
            return False
        file_path = Path(file_path)
        if file_path.name.startswith('~$'):
            return False
        if file_path.suffix.lower() not in PPT_EXTENSIONS:
            return False
        try:
            if file_path.stat().st_size < self.min_file_size_bytes:
                return False
        except:
            return False
        try:
            with open(file_path, 'r+b') as f:
                return False
        except (IOError, OSError) as e:
            return e.errno == 13
        except:
            return True
    
    def has_file_changed(self, file_path):
        try:
            current_mtime = os.path.getmtime(str(file_path))
            self.file_manager.cleanup_old_state()
            if self.file_manager.is_already_processed(file_path):
                stored_mtime = self.file_manager.get_file_mtime(file_path)
                if stored_mtime is not None and abs(current_mtime - stored_mtime) < 1:
                    return False
            return True
        except:
            return True
    
    def copy_ppt_file(self, source_path):
        try:
            source_path = Path(source_path)
            if not source_path.exists() or not self.is_valid_ppt_file_for_backup(source_path):
                return False

            date_folder = self.base_backup_dir / datetime.now().strftime("%Y-%m-%d")
            date_folder.mkdir(exist_ok=True)
            
            dest_path = date_folder / source_path.name
            counter = 1
            original = dest_path
            while dest_path.exists():
                dest_path = original.parent / f"{original.stem}_{counter}{original.suffix}"
                counter += 1

            shutil.copy2(source_path, dest_path)
            self.logger.info(f"已备份: {source_path} -> {dest_path}")
            self.file_manager.add_processed_file(str(source_path), source_path.stat().st_mtime)
            return True
        except Exception as e:
            self.logger.error(f"备份失败: {e}")
            return False
    
    def process_ppt_file(self, file_path, source="事件"):
        with self.processing_lock:
            if str(file_path) in self.currently_processing:
                return False
            self.currently_processing.add(str(file_path))
        
        try:
            if self.is_removable_drive(file_path) and self.is_valid_ppt_file_for_backup(file_path) and self.has_file_changed(file_path):
                self.logger.info(f"[{source}] 检测到PPT文件: {file_path}")
                return self.copy_ppt_file(file_path)
            return False
        finally:
            with self.processing_lock:
                self.currently_processing.discard(str(file_path))
    
    def fallback_monitor_loop(self):
        """后备监控循环"""
        self.logger.info(f"后备监控已启动（间隔: {self.scan_interval}秒）")
        while self.running and self.enable_fallback:
            try:
                ppt_files = set()
                for pid in self.process_cache.get_powerpoint_process_ids():
                    for file_path in self.process_cache.get_process_open_files_by_pid(pid):
                        if file_path and self.is_removable_drive(file_path):
                            if Path(file_path).suffix.lower() in PPT_EXTENSIONS:
                                ppt_files.add(file_path)
                for file_path in ppt_files:
                    self.process_ppt_file(file_path, source="后备扫描")
                
                for _ in range(max(0, self.scan_interval)):
                    if not self.running or not self.enable_fallback:
                        break
                    time.sleep(1)
            except Exception as e:
                self.logger.error(f"后备监控出错: {e}")
                time.sleep(5)
    
    def start_fallback_monitor(self):
        if self.enable_fallback:
            self.fallback_thread = threading.Thread(target=self.fallback_monitor_loop, daemon=True)
            self.fallback_thread.start()
    
    def stop_fallback_monitor(self):
        self.enable_fallback = False
        if self.fallback_thread:
            self.fallback_thread.join(timeout=2)
    
    def set_fallback_enabled(self, enabled):
        if self.enable_fallback == enabled:
            return
        self.enable_fallback = enabled
        self.config_manager.set_enable_fallback(enabled)
        if enabled:
            self.start_fallback_monitor()
        else:
            self.stop_fallback_monitor()
    
    def get_connected_drives(self):
        return sorted(list(self.current_removable_drives_cache))
    
    def get_status_info(self):
        return {
            'connected_drives': list(self.current_removable_drives_cache),
            'processed_today': self.file_manager.get_processed_count(),
            'backup_dir': str(self.base_backup_dir),
            'fallback_enabled': self.enable_fallback,
            'scan_interval': self.scan_interval
        }
    
    def update_config(self, **kwargs):
        if 'backup_dir' in kwargs:
            new_dir = Path(kwargs['backup_dir'])
            new_dir.mkdir(exist_ok=True)
            self.config_manager.set_backup_dir(str(new_dir))
            self.base_backup_dir = new_dir
            self.file_manager = PersistentFileManager(self.base_backup_dir)
        if 'max_retention_days' in kwargs:
            self.config_manager.set_max_retention_days(kwargs['max_retention_days'])
        if 'min_file_size_kb' in kwargs:
            self.config_manager.set_min_file_size_kb(kwargs['min_file_size_kb'])
            self.min_file_size_bytes = kwargs['min_file_size_kb'] * 1024
        if 'enable_fallback' in kwargs:
            self.set_fallback_enabled(kwargs['enable_fallback'])
        if 'scan_interval' in kwargs:
            self.config_manager.set_scan_interval(kwargs['scan_interval'])
            self.scan_interval = kwargs['scan_interval']
    
    def cleanup_old_backups(self):
        try:
            retention_date = datetime.now() - timedelta(days=self.max_retention_days)
            for item in self.base_backup_dir.iterdir():
                if item.is_dir():
                    try:
                        if datetime.strptime(item.name, "%Y-%m-%d") < retention_date:
                            shutil.rmtree(item)
                            self.logger.info(f"删除过期备份: {item}")
                    except ValueError:
                        continue
        except Exception as e:
            self.logger.error(f"清理备份出错: {e}")
    
    def start_monitoring(self):
        self.logger.info("启动PPT监控...")
        self.device_monitor.start_listening()
        self.ppt_com_monitor.start_listening()
        self.start_fallback_monitor()
        
        last_cleanup = datetime.now()
        while self.running:
            try:
                if datetime.now().hour == 0 and (datetime.now() - last_cleanup).seconds > 3600:
                    self.cleanup_old_backups()
                    last_cleanup = datetime.now()
                time.sleep(60)
            except Exception as e:
                self.logger.error(f"监控循环出错: {e}")
                time.sleep(5)
    
    def stop_monitoring(self):
        self.running = False
        self.device_monitor.stop_listening()
        self.ppt_com_monitor.stop_listening()
        self.stop_fallback_monitor()
        self.file_manager.save_state_immediately()