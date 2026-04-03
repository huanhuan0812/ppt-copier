import os
import sys
import time
import shutil
import threading
import json
import win32file
import win32con
import pythoncom
import pywintypes
from win32com.client import GetObject, Dispatch, DispatchWithEvents
import psutil
from datetime import datetime, timedelta
import win32api
import win32gui
import win32event
import winerror
import struct
import configparser
import webbrowser
import logging
from logging.handlers import RotatingFileHandler
from ctypes import windll, c_ulong, byref, POINTER, Structure, WINFUNCTYPE, sizeof
from ctypes.wintypes import HWND, UINT, WPARAM, LPARAM, HICON, DWORD, HANDLE, BOOL
from win32con import PM_REMOVE, DEVICE_NOTIFY_WINDOW_HANDLE
from pathlib import Path
import hashlib
from collections import OrderedDict
from functools import lru_cache

# --- Windows API 常量和结构体 ---
WM_DEVICECHANGE = 0x0219
DBT_DEVICEARRIVAL = 0x8000
DBT_DEVICEREMOVECOMPLETE = 0x8004
DBT_DEVTYP_VOLUME = 0x00000002

class DEV_BROADCAST_HDR(Structure):
    _fields_ = [
        ("dbch_size", DWORD),
        ("dbch_devicetype", DWORD),
        ("dbch_reserved", DWORD),
    ]

class DEV_BROADCAST_VOLUME(Structure):
    _fields_ = [
        ("dbcv_size", DWORD),
        ("dbcv_devicetype", DWORD),
        ("dbcv_reserved", DWORD),
        ("dbcv_unitmask", DWORD),
        ("dbcv_flags", DWORD),
    ]

# --- Constants ---
PPT_EXTENSIONS = {'.ppt', '.pptx', '.pps', '.ppsx'}
APP_NAME = "PPTMonitor"
MUTEX_NAME = "Global\\{B8E2C5A1-9F4D-4E8E-9A2B-3C5D7E9F1A2B}"
STATE_SAVE_INTERVAL = 30  # 状态文件保存间隔（秒）


class SingleInstance:
    """确保只有一个程序实例运行"""
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialize()
        return cls._instance
    
    def _initialize(self):
        self.mutex = None
        self.is_first_instance = False
        try:
            self.mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
            if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
                self.is_first_instance = False
            else:
                self.is_first_instance = True
        except Exception:
            self.is_first_instance = True
    
    def is_first(self):
        return self.is_first_instance
    
    def bring_to_front(self):
        """尝试将已有实例窗口置前"""
        try:
            hwnd = win32gui.FindWindow("PPTMonitorTrayClass", None)
            if hwnd:
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
        except:
            pass


class Logger:
    """日志管理器"""
    _instance = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialize()
        return cls._instance
    
    def _initialize(self):
        """初始化日志系统"""
        if getattr(sys, 'frozen', False):
            base_dir = Path(sys.executable).parent
        else:
            base_dir = Path(__file__).parent
        
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
    
    def debug(self, message):
        self.logger.debug(message)
    
    def info(self, message):
        self.logger.info(message)
    
    def warning(self, message):
        self.logger.warning(message)
    
    def error(self, message):
        self.logger.error(message)
    
    def exception(self, message):
        self.logger.exception(message)


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
            'enable_fallback_monitor': 'false',  # 默认关闭轮询
            'min_file_size_kb': '10',
            'scan_interval_seconds': '30',
            'auto_start': 'false',
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
            self.logger.debug(f"配置文件已保存: {self.config_file}")
        except Exception as e:
            self.logger.error(f"保存配置文件失败: {e}")
    
    def get_backup_dir(self):
        return self.config.get('General', 'backup_dir', fallback='./')
    
    def get_max_retention_days(self):
        try:
            return int(self.config.get('General', 'max_retention_days', fallback='30'))
        except ValueError:
            return 30
    
    def get_enable_fallback_monitor(self):
        return self.config.getboolean('General', 'enable_fallback_monitor', fallback=False)
    
    def get_min_file_size_kb(self):
        try:
            return int(self.config.get('General', 'min_file_size_kb', fallback='10'))
        except ValueError:
            return 10
    
    def get_scan_interval(self):
        try:
            return int(self.config.get('General', 'scan_interval_seconds', fallback='30'))
        except ValueError:
            return 30
    
    def get_auto_start(self):
        return self.config.getboolean('General', 'auto_start', fallback=False)
    
    def get_log_non_removable_events(self):
        return self.config.getboolean('General', 'log_non_removable_events', fallback=False)
    
    def set_backup_dir(self, backup_dir):
        if not self.config.has_section('General'):
            self.config['General'] = {}
        self.config['General']['backup_dir'] = backup_dir
        self.save_config()
        self.logger.info(f"备份目录已更新: {backup_dir}")
    
    def set_max_retention_days(self, days):
        if not self.config.has_section('General'):
            self.config['General'] = {}
        self.config['General']['max_retention_days'] = str(days)
        self.save_config()
        self.logger.info(f"最大保留天数已更新: {days}")
    
    def set_min_file_size_kb(self, size_kb):
        if not self.config.has_section('General'):
            self.config['General'] = {}
        self.config['General']['min_file_size_kb'] = str(size_kb)
        self.save_config()
        self.logger.info(f"最小文件大小已更新: {size_kb}KB")
    
    def set_enable_fallback(self, enabled):
        if not self.config.has_section('General'):
            self.config['General'] = {}
        self.config['General']['enable_fallback_monitor'] = str(enabled)
        self.save_config()
        self.logger.info(f"后备监控已{'启用' if enabled else '禁用'}")
    
    def set_auto_start(self, enabled):
        if not self.config.has_section('General'):
            self.config['General'] = {}
        self.config['General']['auto_start'] = str(enabled)
        self.save_config()
        self.logger.info(f"开机自启已{'启用' if enabled else '禁用'}")
    
    def set_scan_interval(self, seconds):
        if not self.config.has_section('General'):
            self.config['General'] = {}
        self.config['General']['scan_interval_seconds'] = str(seconds)
        self.save_config()
        self.logger.info(f"扫描间隔已更新: {seconds}秒")
    
    def set_log_non_removable_events(self, enabled):
        if not self.config.has_section('General'):
            self.config['General'] = {}
        self.config['General']['log_non_removable_events'] = str(enabled)
        self.save_config()
        self.logger.info(f"非移动设备事件日志已{'启用' if enabled else '禁用'}")


class PersistentFileManager:
    """持久化文件管理器 - 优化版：延迟保存状态"""
    def __init__(self, base_backup_dir):
        self.base_backup_dir = Path(base_backup_dir)
        self.state_file = self.base_backup_dir / 'monitor_state.json'
        self.today = datetime.now().strftime("%Y-%m-%d")
        self.processed_files = {}
        self.lock = threading.Lock()
        self.logger = Logger()
        
        self._dirty = False
        self._last_save_time = 0
        
        self.load_state()
        self._start_auto_save_timer()
    
    def _start_auto_save_timer(self):
        """启动自动保存定时器"""
        def auto_save_loop():
            while True:
                time.sleep(STATE_SAVE_INTERVAL)
                if self._dirty:
                    self._do_save()
        
        save_thread = threading.Thread(target=auto_save_loop, daemon=True)
        save_thread.start()
    
    def _do_save(self):
        """实际执行保存操作"""
        with self.lock:
            if not self._dirty:
                return
            data = {'date': self.today, 'processed_files': self.processed_files}
            try:
                with open(self.state_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self._dirty = False
                self._last_save_time = time.time()
            except Exception as e:
                self.logger.error(f"保存状态文件失败: {str(e)}")
    
    def load_state(self):
        if self.state_file.exists():
            try:
                with open(self.state_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    stored_date = data.get('date', '')
                    if stored_date == self.today:
                        self.processed_files = data.get('processed_files', {})
                        self.logger.info(f"加载状态成功，今日已处理 {len(self.processed_files)} 个文件")
                    else:
                        self.processed_files = {}
                        self.logger.info(f"状态文件日期不匹配，已重置（存储日期: {stored_date}, 当前日期: {self.today}）")
            except (json.JSONDecodeError, KeyError) as e:
                self.processed_files = {}
                self.logger.error(f"加载状态文件失败: {e}")
        else:
            self.processed_files = {}
    
    def save_state_immediately(self):
        """立即保存状态（用于程序退出时）"""
        self._do_save()
    
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


class ProcessCache:
    """进程扫描缓存 - 优化版，避免锁竞争和长时间阻塞"""
    
    def __init__(self, ttl_seconds=2):
        self.cache = {}
        self.ttl = ttl_seconds
        self.logger = Logger()
        self._lock = threading.RLock()
        self._scan_lock = threading.Lock()
        
        # 缓存PowerPoint进程ID列表
        self.pp_process_ids = []
        self.pp_process_ids_cache_time = 0
        self._is_scanning = False
        
        # 缓存每个进程打开的文件
        self.process_files_cache = {}
        
        # 记录无效的进程ID
        self.invalid_pids = set()
        self.invalid_pids_cleanup_time = 0
        
        # 关闭标志
        self._shutting_down = False
    
    def set_shutting_down(self, value):
        """设置关闭标志，停止扫描"""
        self._shutting_down = value
    
    def get_powerpoint_process_ids(self):
        """获取PowerPoint进程ID列表（带缓存，非阻塞）"""
        if self._shutting_down:
            return []
        
        with self._lock:
            now = time.time()
            if now - self.pp_process_ids_cache_time < self.ttl:
                return self.pp_process_ids.copy()
        
        if self._is_scanning:
            with self._lock:
                return self.pp_process_ids.copy()
        
        if not self._scan_lock.acquire(blocking=False):
            with self._lock:
                return self.pp_process_ids.copy()
        
        try:
            self._is_scanning = True
            pids = []
            try:
                for proc in psutil.process_iter(['pid', 'name']):
                    if self._shutting_down:
                        break
                    try:
                        name = proc.info['name']
                        if name and name.lower().startswith('powerpnt'):
                            pids.append(proc.info['pid'])
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        continue
            except Exception as e:
                self.logger.error(f"扫描PowerPoint进程失败: {e}")
            
            with self._lock:
                self.pp_process_ids = pids
                self.pp_process_ids_cache_time = time.time()
                return pids.copy()
        finally:
            self._is_scanning = False
            self._scan_lock.release()
    
    def get_process_open_files_by_pid(self, pid):
        """通过进程ID获取打开的文件列表（带缓存，快速失败）"""
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
            try:
                for file_info in proc.open_files():
                    if self._shutting_down:
                        break
                    open_files.append(file_info.path)
            except psutil.AccessDenied:
                pass
            except psutil.NoSuchProcess:
                with self._lock:
                    self.invalid_pids.add(pid)
                return []
            except Exception:
                return []
            
            with self._lock:
                self.process_files_cache[pid] = (time.time(), open_files)
                
                if len(self.process_files_cache) > 100:
                    oldest_keys = sorted(self.process_files_cache.keys(), 
                                       key=lambda k: self.process_files_cache[k][0])[:50]
                    for key in oldest_keys:
                        del self.process_files_cache[key]
            
            return open_files
            
        except psutil.NoSuchProcess:
            with self._lock:
                self.invalid_pids.add(pid)
            return []
        except Exception:
            return []
    
    def invalidate_process(self, pid=None):
        """异步使缓存失效"""
        def async_invalidate():
            try:
                with self._lock:
                    if pid is not None:
                        if pid in self.process_files_cache:
                            del self.process_files_cache[pid]
                        self.invalid_pids.discard(pid)
                    else:
                        self.process_files_cache.clear()
                        self.invalid_pids.clear()
            except Exception:
                pass
        
        threading.Thread(target=async_invalidate, daemon=True).start()
    
    def clear_all(self):
        """清空所有缓存 - 异步执行"""
        def async_clear():
            try:
                with self._lock:
                    self.process_files_cache.clear()
                    self.invalid_pids.clear()
                    self.pp_process_ids = []
                    self.pp_process_ids_cache_time = 0
            except Exception:
                pass
        
        threading.Thread(target=async_clear, daemon=True).start()


class PowerPointEvents:
    """PowerPoint COM 事件处理器 - 优化版，避免关闭时卡顿"""
    def __init__(self, monitor_instance=None):
        self.monitor = monitor_instance
        self.logger = Logger()
    
    def set_monitor(self, monitor):
        self.monitor = monitor
    
    def _should_process_file(self, file_path):
        if not file_path or not self.monitor:
            return False
        return self.monitor.is_valid_ppt_file_for_backup(file_path)
    
    def _is_on_removable_drive(self, file_path):
        if not file_path or not self.monitor:
            return False
        return self.monitor.is_removable_drive(file_path)
    
    def OnPresentationOpen(self, Pres):
        """演示文稿打开事件"""
        try:
            file_path = None
            try:
                if hasattr(Pres, 'FullName'):
                    file_path = Pres.FullName
            except Exception:
                return
            
            if file_path and self._is_on_removable_drive(file_path):
                self.logger.info(f"COM事件 - PowerPoint打开文件: {file_path}")
                
                def delayed_check():
                    time.sleep(0.5)
                    if self._should_process_file(file_path):
                        if self.monitor:
                            self.monitor.process_ppt_file(file_path, source="COM事件(打开)")
                
                threading.Thread(target=delayed_check, daemon=True).start()
        except Exception as e:
            self.logger.error(f"处理OnPresentationOpen事件时出错: {e}")
    
    def OnPresentationClose(self, Pres):
        """演示文稿关闭事件 - 异步处理，避免阻塞"""
        try:
            if self.monitor and hasattr(self.monitor, 'process_cache'):
                def async_cleanup():
                    try:
                        time.sleep(0.1)
                        self.monitor.process_cache.invalidate_process()
                    except Exception:
                        pass
                
                threading.Thread(target=async_cleanup, daemon=True).start()
        except Exception:
            pass
    
    def OnPresentationSave(self, Pres):
        """演示文稿保存事件"""
        try:
            file_path = None
            try:
                if hasattr(Pres, 'FullName'):
                    file_path = Pres.FullName
            except Exception:
                return
            
            if file_path and self._is_on_removable_drive(file_path):
                self.logger.info(f"COM事件 - PowerPoint保存文件: {file_path}")
                
                if self._should_process_file(file_path):
                    if self.monitor:
                        self.monitor.process_ppt_file(file_path, source="COM事件(保存)")
        except Exception as e:
            self.logger.error(f"处理OnPresentationSave事件时出错: {e}")
    
    def OnQuit(self):
        """PowerPoint退出事件 - 异步处理"""
        self.logger.info("COM事件 - PowerPoint正在退出")
        if self.monitor:
            def async_quit():
                try:
                    time.sleep(0.1)
                    self.monitor.on_powerpoint_quit()
                    if hasattr(self.monitor, 'process_cache'):
                        self.monitor.process_cache.clear_all()
                except Exception:
                    pass
            
            threading.Thread(target=async_quit, daemon=True).start()


class PowerPointEventMonitor:
    """PowerPoint COM 事件监听器（优雅处理退出）"""
    def __init__(self, monitor):
        self.monitor = monitor
        self.running = False
        self.event_thread = None
        self.powerpoint = None
        self.logger = Logger()
        self.event_handler = None
        self.reconnect_delay = 3
        self.is_quitting = False
    
    def start_listening(self):
        if self.running:
            return
        self.running = True
        self.event_thread = threading.Thread(target=self._run_com_loop, daemon=True)
        self.event_thread.start()
        self.logger.info("PowerPoint COM事件监听器已启动")
    
    def stop_listening(self):
        self.running = False
        if self.monitor and hasattr(self.monitor, 'process_cache'):
            self.monitor.process_cache.set_shutting_down(True)
        
        if self.event_thread:
            self.event_thread.join(timeout=3)
        self.logger.info("PowerPoint COM事件监听器已停止")
    
    def _is_powerpoint_process_running(self):
        """快速检查PowerPoint进程是否在运行"""
        if self.monitor and hasattr(self.monitor, 'process_cache'):
            pids = self.monitor.process_cache.get_powerpoint_process_ids()
            return len(pids) > 0
        else:
            try:
                for proc in psutil.process_iter(['name']):
                    if proc.info['name'] and proc.info['name'].lower().startswith('powerpnt'):
                        return True
                return False
            except:
                return False
    
    def _is_quitting_error(self, e):
        quitting_hresults = [
            -2147417848,  # RPC_E_SERVER_DIED
            -2147023174,  # RPC_S_SERVER_UNAVAILABLE
            -2147418113,  # RPC_E_SERVER_DIED_DNE
        ]
        return e.hresult in quitting_hresults if hasattr(e, 'hresult') else False
    
    def _try_connect_powerpoint(self):
        try:
            if not self._is_powerpoint_process_running():
                return None
            
            try:
                powerpoint = Dispatch('PowerPoint.Application')
                if powerpoint:
                    try:
                        _ = powerpoint.Name
                        self.logger.info("通过 Dispatch 成功连接到 PowerPoint")
                        return powerpoint
                    except:
                        pass
            except Exception as e:
                self.logger.debug(f"Dispatch 连接失败: {e}")
            
            try:
                powerpoint = GetObject(Class='PowerPoint.Application')
                if powerpoint:
                    try:
                        _ = powerpoint.Name
                        self.logger.info("通过 GetObject 成功连接到 PowerPoint")
                        return powerpoint
                    except:
                        pass
            except Exception as e:
                self.logger.debug(f"GetObject 连接失败: {e}")
            
            return None
            
        except Exception as e:
            self.logger.error(f"连接PowerPoint时发生未知错误: {e}")
        return None
    
    def _setup_event_handler(self, powerpoint):
        try:
            event_handler = DispatchWithEvents(powerpoint, PowerPointEvents)
            if hasattr(event_handler, 'set_monitor'):
                event_handler.set_monitor(self.monitor)
            self.logger.info("PowerPoint事件处理器已设置")
            return event_handler
        except Exception as e:
            self.logger.error(f"设置事件处理器失败: {e}")
            return None
    
    def _process_existing_presentations_safe(self, powerpoint):
        try:
            presentations_count = 0
            try:
                presentations_count = powerpoint.Presentations.Count
            except pywintypes.com_error as e:
                if e.hresult == -2147352567:
                    self.logger.debug("无法获取演示文稿数量（权限不足）")
                else:
                    self.logger.debug(f"获取演示文稿数量失败: {e}")
                return
            
            if presentations_count > 0:
                self.logger.info(f"检测到 {presentations_count} 个已打开的演示文稿")
                
                for i in range(presentations_count):
                    try:
                        pres = powerpoint.Presentations.Item(i + 1)
                        file_path = None
                        
                        try:
                            if hasattr(pres, 'FullName'):
                                file_path = pres.FullName
                        except pywintypes.com_error:
                            continue
                        except Exception:
                            continue
                        
                        if file_path and self.monitor and self.monitor.is_removable_drive(file_path):
                            self.logger.info(f"检测到已打开的演示文稿: {file_path}")
                            
                            def delayed_process(fp=file_path):
                                time.sleep(1)
                                if self.monitor and self.monitor.is_valid_ppt_file_for_backup(fp):
                                    self.monitor.process_ppt_file(fp, source="已打开文件")
                            
                            threading.Thread(target=delayed_process, daemon=True).start()
                    except Exception:
                        continue
                        
        except Exception as e:
            self.logger.error(f"遍历已打开演示文稿时出错: {e}")
    
    def _run_com_loop(self):
        pythoncom.CoInitialize()
        
        self.logger.info("COM事件循环已启动")
        last_pp_check = 0
        self.is_quitting = False
        
        while self.running:
            try:
                current_time = time.time()
                
                if self.powerpoint is None and not self.is_quitting:
                    if self._is_powerpoint_process_running():
                        self.powerpoint = self._try_connect_powerpoint()
                        if self.powerpoint:
                            self.logger.info("已成功连接到PowerPoint")
                            self.event_handler = self._setup_event_handler(self.powerpoint)
                            self._process_existing_presentations_safe(self.powerpoint)
                        else:
                            time.sleep(self.reconnect_delay)
                    else:
                        time.sleep(5)
                
                elif self.powerpoint:
                    try:
                        if current_time - last_pp_check > 10:
                            last_pp_check = current_time
                            
                            if not self._is_powerpoint_process_running():
                                self.logger.info("PowerPoint进程已退出，正常断开连接")
                                self.powerpoint = None
                                self.event_handler = None
                                self.is_quitting = False
                                continue
                            
                            try:
                                _ = self.powerpoint.Name
                            except pywintypes.com_error as e:
                                if self._is_quitting_error(e):
                                    self.logger.info("PowerPoint正在正常退出，断开连接")
                                    self.powerpoint = None
                                    self.event_handler = None
                                    self.is_quitting = False
                                    continue
                                else:
                                    self.logger.warning(f"PowerPoint COM对象异常: {e}")
                                    self.powerpoint = None
                                    self.event_handler = None
                                    time.sleep(self.reconnect_delay)
                                    continue
                        
                        pythoncom.PumpWaitingMessages()
                        time.sleep(0.05)
                        
                    except pywintypes.com_error as e:
                        if self._is_quitting_error(e):
                            self.logger.info("PowerPoint正在退出，COM连接正常断开")
                            self.powerpoint = None
                            self.event_handler = None
                            self.is_quitting = False
                        else:
                            self.logger.warning(f"PowerPoint COM错误: {e}")
                            self.powerpoint = None
                            self.event_handler = None
                            time.sleep(self.reconnect_delay)
                    except Exception as e:
                        self.logger.error(f"处理COM消息时出错: {e}")
                        time.sleep(0.5)
                
            except Exception as e:
                self.logger.error(f"COM事件循环出错: {e}", exc_info=True)
                time.sleep(1)
        
        pythoncom.CoUninitialize()
        self.logger.info("COM事件循环已结束")


class WindowsDeviceMonitor:
    """Windows设备变化监听器"""
    def __init__(self, callback):
        self.callback = callback
        self.running = False
        self.removable_drives = set()
        self.logger = Logger()
        self.hwnd = None
        self.event_thread = None

        wc = win32gui.WNDCLASS()
        wc.lpszClassName = "PPTMonitorDeviceListener"
        wc.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wc.hbrBackground = win32con.COLOR_WINDOW
        wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpfnWndProc = self.wnd_proc
        self.class_atom = win32gui.RegisterClass(wc)

    def wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == WM_DEVICECHANGE:
            if wparam in (DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE):
                try:
                    hdr_ptr = LPARAM(lparam).value
                    if hdr_ptr:
                        hdr = DEV_BROADCAST_HDR.from_address(hdr_ptr)
                        if hdr.dbch_devicetype == DBT_DEVTYP_VOLUME:
                            vol_struct = DEV_BROADCAST_VOLUME.from_address(hdr_ptr)
                            unit_mask = vol_struct.dbcv_unitmask
                            
                            detected_drives = set()
                            for i in range(26):
                                if unit_mask & (1 << i):
                                    detected_drives.add(chr(ord('A') + i))

                            if wparam == DBT_DEVICEARRIVAL:
                                self.logger.info(f"检测到新插入的驱动器: {['{}:'.format(d) for d in detected_drives]}")
                                self.callback('device_inserted', list(detected_drives))
                            elif wparam == DBT_DEVICEREMOVECOMPLETE:
                                self.logger.info(f"检测到移除的驱动器: {['{}:'.format(d) for d in detected_drives]}")
                                self.callback('device_removed', list(detected_drives))
                except Exception as e:
                    self.logger.error(f"解析设备变更消息时出错: {e}")
            return True
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def start_listening(self):
        if self.running:
            return
        self.running = True
        self.event_thread = threading.Thread(target=self._run_message_loop, daemon=True)
        self.event_thread.start()
        self.logger.info("Windows设备事件监听器已启动")

    def stop_listening(self):
        self.running = False
        if self.hwnd:
            try:
                win32gui.PostMessage(self.hwnd, win32con.WM_QUIT, 0, 0)
            except:
                pass
        if self.event_thread:
            self.event_thread.join(timeout=2)
        self.logger.info("Windows设备事件监听器已停止")

    def _run_message_loop(self):
        self.hwnd = win32gui.CreateWindowEx(
            0,
            self.class_atom,
            "PPTMonitorDeviceListenerWindow",
            0,
            0, 0, 0, 0,
            0,
            0,
            win32api.GetModuleHandle(None),
            None
        )

        initial_drives = self.get_removable_drives()
        self.logger.info(f"初始可移动驱动器: {[f'{d}:' for d in initial_drives]}")
        if initial_drives:
            self.callback('device_inserted', list(initial_drives))

        dbv = DEV_BROADCAST_VOLUME()
        dbv.dbcv_size = sizeof(DEV_BROADCAST_VOLUME)
        dbv.dbcv_devicetype = DBT_DEVTYP_VOLUME
        dbv.dbcv_reserved = 0
        dbv.dbcv_unitmask = 0
        dbv.dbcv_flags = 0

        dev_broadcast_handle = windll.user32.RegisterDeviceNotificationW(
            self.hwnd,
            byref(dbv),
            DEVICE_NOTIFY_WINDOW_HANDLE
        )
        if not dev_broadcast_handle:
            self.logger.error(f"注册设备通知失败: {win32api.GetLastError()}")
        else:
            self.logger.debug("已注册设备通知")

        while self.running:
            try:
                ret, msg_tuple = win32gui.GetMessage(self.hwnd, 0, 0)
            
                if ret == 0:
                    self.logger.debug("收到 WM_QUIT，退出消息循环")
                    break
                elif ret > 0:
                    win32gui.DispatchMessage(msg_tuple)
                else:
                    self.logger.error(f"GetMessage 错误: {win32api.GetLastError()}")
                    time.sleep(0.1)
            except Exception as e:
                self.logger.error(f"消息循环中发生错误: {e}")
                time.sleep(0.1)

        if dev_broadcast_handle:
            windll.user32.UnregisterDeviceNotification(dev_broadcast_handle)
        try:
            win32gui.DestroyWindow(self.hwnd)
        except:
            pass
        self.hwnd = None

    def get_removable_drives(self):
        drives = set()
        drive_bitmask = win32file.GetLogicalDrives()
        for i in range(26):
            if drive_bitmask & (1 << i):
                drive_letter = chr(ord('A') + i) + ':\\'
                try:
                    drive_type = win32file.GetDriveType(drive_letter)
                    if drive_type in [win32file.DRIVE_REMOVABLE, win32file.DRIVE_CDROM]:
                        drives.add(drive_letter[0].upper())
                except:
                    continue
        return drives


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
        self.log_non_removable = self.config_manager.get_log_non_removable_events()
        
        self.file_manager = PersistentFileManager(self.base_backup_dir)
        self.ppt_extensions = PPT_EXTENSIONS
        
        self.process_cache = ProcessCache(ttl_seconds=2)
        
        self.device_monitor = WindowsDeviceMonitor(self.on_device_event)
        self.ppt_com_monitor = PowerPointEventMonitor(self)
        
        self.connected_drives = set()
        self.current_removable_drives_cache = set()
        
        self.processing_lock = threading.Lock()
        self.currently_processing = set()
        
        self.running = True
        self.fallback_thread = None
        
        self.logger.info(f"备份目录: {self.base_backup_dir}")
        self.logger.info(f"最大保留天数: {self.max_retention_days}")
        self.logger.info(f"最小文件大小: {self.config_manager.get_min_file_size_kb()}KB")
        self.logger.info(f"扫描间隔: {self.scan_interval}秒")
        self.logger.info(f"后备进程监控: {'启用' if self.enable_fallback else '禁用'}")
        self.logger.info(f"非移动设备日志: {'启用' if self.log_non_removable else '禁用'}")

    def on_powerpoint_quit(self):
        self.logger.info("PowerPoint已正常退出")

    def cleanup_old_backups(self):
        try:
            now = datetime.now()
            retention_date = now - timedelta(days=self.max_retention_days)
            
            base_path = Path(self.base_backup_dir)
            deleted_count = 0
            for item in base_path.iterdir():
                if item.is_dir():
                    try:
                        dir_date = datetime.strptime(item.name, "%Y-%m-%d")
                        if dir_date < retention_date:
                            shutil.rmtree(item)
                            deleted_count += 1
                            self.logger.info(f"已删除过期备份文件夹: {item}")
                    except ValueError:
                        continue
            
            if deleted_count > 0:
                self.logger.info(f"清理完成，共删除 {deleted_count} 个过期备份")
        except Exception as e:
            self.logger.error(f"清理旧备份时出错: {str(e)}")

    def get_date_folder_path(self):
        today = datetime.now().strftime("%Y-%m-%d")
        date_folder = self.base_backup_dir / today
        date_folder.mkdir(exist_ok=True)
        return date_folder

    def is_removable_drive(self, path):
        if not path or not isinstance(path, (str, Path)):
            return False
        try:
            path = Path(path)
            drive_letter = str(path.drive).upper().rstrip(':')
            return drive_letter in self.current_removable_drives_cache
        except:
            return False

    def is_valid_ppt_file_for_backup(self, file_path):
        if not file_path:
            return False
        
        file_path = Path(file_path)
        filename = file_path.name
        
        if filename.startswith('~$'):
            return False
        
        try:
            attrs = win32file.GetFileAttributes(str(file_path))
            if attrs & win32con.FILE_ATTRIBUTE_HIDDEN:
                return False
        except:
            pass
        
        if file_path.suffix.lower() not in self.ppt_extensions:
            return False
        
        try:
            file_size = file_path.stat().st_size
            if file_size < self.min_file_size_bytes:
                return False
        except Exception as e:
            self.logger.error(f"获取文件大小失败: {e}")
            return False
        
        try:
            with open(file_path, 'r+b') as f:
                return False
        except (IOError, OSError) as e:
            if e.errno == 13:
                return True
            else:
                return False
        except Exception as e:
            self.logger.error(f"检查文件锁定状态失败: {e}")
            return True

    def has_file_changed(self, file_path):
        try:
            file_path = str(file_path)
            current_mtime = os.path.getmtime(file_path)
            self.file_manager.cleanup_old_state()
            
            if self.file_manager.is_already_processed(file_path):
                stored_mtime = self.file_manager.get_file_mtime(file_path)
                if stored_mtime is not None:
                    if abs(current_mtime - stored_mtime) < 1:
                        return False
                    else:
                        return True
                else:
                    return True
            else:
                return True
        except Exception as e:
            self.logger.error(f"检查文件更改失败: {e}")
            return True

    def copy_ppt_file(self, source_path):
        try:
            source_path = Path(source_path)
            if not source_path.exists():
                self.logger.warning(f"源文件不存在: {source_path}")
                return False

            if not self.is_valid_ppt_file_for_backup(source_path):
                return False

            date_folder = self.get_date_folder_path()
            filename = source_path.name
            dest_path = date_folder / filename

            counter = 1
            original_dest_path = dest_path
            while dest_path.exists():
                dest_path = original_dest_path.parent / f"{original_dest_path.stem}_{counter}{original_dest_path.suffix}"
                counter += 1

            shutil.copy2(source_path, dest_path)
            self.logger.info(f"已备份PPT文件: {source_path} -> {dest_path}")
            
            current_mtime = source_path.stat().st_mtime
            self.file_manager.add_processed_file(str(source_path), current_mtime)
            
            self._release_com_resources()

            return True
        except Exception as e:
            self.logger.error(f"备份失败 {source_path}: {str(e)}", exc_info=True)
            return False
    
    def _release_com_resources(self):
        """释放COM资源，避免内存累积"""
        try:
            pythoncom.PumpWaitingMessages()
        except:
            pass

    def process_ppt_file(self, file_path, source="事件"):
        with self.processing_lock:
            if str(file_path) in self.currently_processing:
                return False
            self.currently_processing.add(str(file_path))
        
        try:
            if not self.is_removable_drive(file_path):
                return False
            
            if not self.is_valid_ppt_file_for_backup(file_path):
                return False
            
            if not self.has_file_changed(file_path):
                return False
            
            self.logger.info(f"[{source}] 检测到移动设备上的PPT文件: {file_path}")
            return self.copy_ppt_file(file_path)
        finally:
            with self.processing_lock:
                self.currently_processing.discard(str(file_path))

    def on_device_event(self, event_type, drives):
        if event_type == 'device_inserted':
            self.connected_drives.update(drives)
            self.current_removable_drives_cache.update(drives)
            self.logger.info(f"设备插入事件: {drives}")
            self.logger.info(f"当前连接的可移动驱动器: {sorted(list(self.connected_drives))}")
            self.process_cache.invalidate_process()
        elif event_type == 'device_removed':
            for drive in drives:
                self.connected_drives.discard(drive)
                self.current_removable_drives_cache.discard(drive)
            self.logger.info(f"设备移除事件: {drives}")
            self.logger.info(f"当前连接的可移动驱动器: {sorted(list(self.connected_drives))}")
            self.process_cache.invalidate_process()
            self._release_com_resources()

    def fallback_monitor_loop(self):
        """后备监控循环 - 通过缓存的进程ID扫描，避免COM调用导致的卡顿"""
        self.logger.info(f"后备进程监控已启动（扫描间隔: {self.scan_interval}秒）")
        
        while self.running and self.enable_fallback:
            try:
                scan_start = time.time()
                
                ppt_pids = self.process_cache.get_powerpoint_process_ids()
                
                if not ppt_pids:
                    time.sleep(self.scan_interval)
                    continue
                
                ppt_files_from_process = set()
                
                for pid in ppt_pids:
                    try:
                        open_files = self.process_cache.get_process_open_files_by_pid(pid)
                        for file_path in open_files:
                            if file_path and self.is_removable_drive(file_path):
                                file_ext = os.path.splitext(file_path)[1].lower()
                                if file_ext in self.ppt_extensions:
                                    ppt_files_from_process.add(file_path)
                    except Exception:
                        continue
                
                for file_path in ppt_files_from_process:
                    self.process_ppt_file(file_path, source="后备扫描")
                
                scan_duration = time.time() - scan_start
                wait_time = max(0, self.scan_interval - scan_duration)
                
                for _ in range(int(wait_time)):
                    if not self.running or not self.enable_fallback:
                        break
                    time.sleep(1)
                    
            except Exception as e:
                self.logger.error(f"后备监控循环出错: {str(e)}")
                time.sleep(5)
        
        self.logger.info("后备进程监控已停止")

    def start_fallback_monitor(self):
        if self.enable_fallback:
            self.fallback_thread = threading.Thread(target=self.fallback_monitor_loop, daemon=True)
            self.fallback_thread.start()
            self.logger.info("后备进程监控已启动")
        else:
            self.logger.info("后备进程监控已禁用（轮询开关关闭）")

    def stop_fallback_monitor(self):
        """停止后备监控"""
        self.enable_fallback = False
        if self.fallback_thread:
            self.fallback_thread.join(timeout=2)
    
    def set_fallback_enabled(self, enabled):
        """动态设置后备监控开关"""
        if self.enable_fallback == enabled:
            return
        
        self.enable_fallback = enabled
        self.config_manager.set_enable_fallback(enabled)
        
        if enabled:
            # 启动后备监控
            if not self.fallback_thread or not self.fallback_thread.is_alive():
                self.start_fallback_monitor()
            self.logger.info("后备监控已启用")
        else:
            # 停止后备监控
            self.stop_fallback_monitor()
            self.logger.info("后备监控已禁用")

    def get_connected_drives(self):
        return sorted(list(self.connected_drives))
    
    def get_status_info(self):
        return {
            'connected_drives': list(self.connected_drives),
            'processed_today': self.file_manager.get_processed_count(),
            'backup_dir': str(self.base_backup_dir),
            'max_retention_days': self.max_retention_days,
            'min_file_size_kb': self.min_file_size_bytes // 1024,
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
            self.max_retention_days = kwargs['max_retention_days']
        
        if 'min_file_size_kb' in kwargs:
            self.config_manager.set_min_file_size_kb(kwargs['min_file_size_kb'])
            self.min_file_size_bytes = kwargs['min_file_size_kb'] * 1024
        
        if 'enable_fallback' in kwargs:
            self.set_fallback_enabled(kwargs['enable_fallback'])
        
        if 'scan_interval' in kwargs:
            self.config_manager.set_scan_interval(kwargs['scan_interval'])
            self.scan_interval = kwargs['scan_interval']
        
        if 'log_non_removable' in kwargs:
            self.config_manager.set_log_non_removable_events(kwargs['log_non_removable'])
            self.log_non_removable = kwargs['log_non_removable']

    def start_monitoring(self):
        self.logger.info("=" * 50)
        self.logger.info("启动PPT监控（事件驱动模式）...")
        self.logger.info(f"基础备份目录: {self.base_backup_dir}")
        self.logger.info(f"最大保留天数: {self.max_retention_days}")
        self.logger.info(f"最小文件大小: {self.min_file_size_bytes} bytes")
        self.logger.info(f"今日已处理文件数量: {len(self.file_manager.processed_files)}")
        self.logger.info(f"状态保存间隔: {STATE_SAVE_INTERVAL}秒")
        self.logger.info(f"后备监控（轮询）: {'启用' if self.enable_fallback else '禁用'}")

        self.device_monitor.start_listening()
        self.ppt_com_monitor.start_listening()
        self.start_fallback_monitor()
        
        self.logger.info("所有事件监听器已启动，等待事件触发...")
        
        last_cleanup = datetime.now()
        while self.running:
            try:
                now = datetime.now()
                if now.hour == 0 and (now - last_cleanup).seconds > 3600:
                    self.cleanup_old_backups()
                    last_cleanup = now
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
        self.logger.info("监控已停止")


class SystemTrayApp:
    def __init__(self):
        self.logger = Logger()
        self.WM_NOTIFY_CALLBACK = win32gui.RegisterWindowMessage("NotifyCallback")
        self.hwnd = None
        self.monitor = None
        self.monitor_thread = None
        self.config_manager = ConfigManager()
        self.running = True
        
        self.logger.info("初始化系统托盘应用")
        
        wc = win32gui.WNDCLASS()
        wc.lpszClassName = "PPTMonitorTrayClass"
        wc.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wc.hbrBackground = win32con.COLOR_WINDOW
        wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpfnWndProc = self.wnd_proc
        self.class_atom = win32gui.RegisterClass(wc)
        
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = win32gui.CreateWindow(
            self.class_atom,
            "PPT Monitor",
            style,
            0, 0, 200, 200,
            0, 0,
            win32api.GetModuleHandle(None),
            None
        )
        
        self.create_tray_icon()
    
    def create_tray_icon(self):
        icon_path = Path(__file__).parent / "icon.ico"
        try:
            if icon_path.exists():
                hicon = win32gui.LoadImage(
                    0, str(icon_path), win32con.IMAGE_ICON, 0, 0, 
                    win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
                )
            else:
                hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
            self.logger.info("成功加载托盘图标")
        except Exception as e:
            self.logger.warning(f"加载图标失败，使用默认图标: {e}")
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
    
        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        nid = (self.hwnd, 0, flags, self.WM_NOTIFY_CALLBACK, hicon, "PPT文件备份监控")
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        self.logger.debug("托盘图标已添加")
    
    def update_tray_tooltip(self):
        if self.monitor:
            drives = self.monitor.get_connected_drives()
            count = self.monitor.file_manager.get_processed_count()
            fallback_status = "轮询:开" if self.monitor.enable_fallback else "轮询:关"
            if drives:
                tip = f"PPT备份监控\n已连接: {', '.join([f'{d}:' for d in drives])}\n今日备份: {count}个\n{fallback_status}"
            else:
                tip = f"PPT备份监控\n无连接设备\n今日备份: {count}个\n{fallback_status}"
        else:
            tip = "PPT备份监控\n未启动"
        
        try:
            nid = (self.hwnd, 0, win32gui.NIF_TIP, 0, 0, tip)
            win32gui.Shell_NotifyIcon(win32gui.NIM_MODIFY, nid)
        except:
            pass
    
    def open_config_file(self):
        config_file = Path(self.config_manager.config_file)
        if config_file.exists():
            try:
                os.startfile(str(config_file))
                self.logger.info(f"已打开配置文件: {config_file}")
            except Exception as e:
                self.logger.error(f"打开配置文件失败: {str(e)}")
        else:
            self.logger.warning(f"配置文件不存在: {config_file}")
    
    def open_backup_folder(self):
        if self.monitor:
            try:
                os.startfile(str(self.monitor.base_backup_dir))
                self.logger.info(f"已打开备份文件夹: {self.monitor.base_backup_dir}")
            except Exception as e:
                self.logger.error(f"打开备份文件夹失败: {str(e)}")
        else:
            backup_dir = self.config_manager.get_backup_dir()
            try:
                os.startfile(backup_dir)
            except Exception as e:
                self.logger.error(f"打开备份文件夹失败: {str(e)}")
    
    def open_log_folder(self):
        logger = Logger()
        try:
            os.startfile(str(logger.log_dir))
            self.logger.info("已打开日志文件夹")
        except Exception as e:
            self.logger.error(f"打开日志文件夹失败: {str(e)}")
    
    def toggle_fallback_monitor(self):
        """切换后备监控开关"""
        if self.monitor:
            new_state = not self.monitor.enable_fallback
            self.monitor.set_fallback_enabled(new_state)
            self.update_tray_tooltip()
            self.logger.info(f"后备监控已{'启用' if new_state else '禁用'}")
    
    def show_context_menu(self, x, y):
        menu = win32gui.CreatePopupMenu()
        
        status_text = "无连接设备"
        processed_count = 0
        fallback_status = ""
        if self.monitor:
            connected_drives = self.monitor.get_connected_drives()
            if connected_drives:
                status_text = f"连接设备: {', '.join([f'{d}:' for d in connected_drives])}"
            else:
                status_text = "无连接设备"
            processed_count = self.monitor.file_manager.get_processed_count()
            fallback_status = f"轮询监控: {'开启' if self.monitor.enable_fallback else '关闭'}"
        
        win32gui.AppendMenu(menu, win32con.MF_GRAYED, 0, f"状态: {status_text}")
        win32gui.AppendMenu(menu, win32con.MF_GRAYED, 0, f"今日已备份: {processed_count} 个文件")
        if fallback_status:
            win32gui.AppendMenu(menu, win32con.MF_GRAYED, 0, fallback_status)
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        
        # 设置子菜单
        settings_menu = win32gui.CreatePopupMenu()
        win32gui.AppendMenu(settings_menu, win32con.MF_STRING, 1000, "编辑配置文件")
        win32gui.AppendMenu(settings_menu, win32con.MF_STRING, 1002, "打开备份文件夹")
        win32gui.AppendMenu(settings_menu, win32con.MF_STRING, 1003, "打开日志文件夹")
        win32gui.AppendMenu(settings_menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(settings_menu, win32con.MF_STRING, 1004, "切换轮询监控")
        win32gui.AppendMenu(menu, win32con.MF_POPUP, settings_menu, "设置")
        
        # 帮助子菜单
        help_menu = win32gui.CreatePopupMenu()
        win32gui.AppendMenu(help_menu, win32con.MF_STRING, 2001, "关于")
        win32gui.AppendMenu(menu, win32con.MF_POPUP, help_menu, "帮助")
        
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING, 1001, "退出")
        
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, x, y, 0, self.hwnd, None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
    
    def show_about_dialog(self):
        about_text = f"""PPT文件备份监控程序
版本: 2.3
功能:
- 自动检测U盘等移动设备
- 实时监控PowerPoint文件编辑
- 自动备份PPT文件到本地
- 支持自定义备份目录和保留天数
- 进程ID缓存优化（避免COM调用卡顿）
- 可选的轮询监控（默认关闭）
- 延迟保存状态（{STATE_SAVE_INTERVAL}秒间隔）
- 智能日志过滤（只记录移动设备事件）

轮询监控说明:
- 开启后会定期扫描PowerPoint进程打开的文件
- 可能会略微增加系统开销
- 建议仅在COM事件监控不工作时开启

备份目录: {self.config_manager.get_backup_dir()}
保留天数: {self.config_manager.get_max_retention_days()}天

© 2026
"""
        win32api.MessageBox(self.hwnd, about_text, "关于", win32con.MB_OK | win32con.MB_ICONINFORMATION)
    
    def wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == self.WM_NOTIFY_CALLBACK:
            if lparam == win32con.WM_RBUTTONUP:
                pos = win32gui.GetCursorPos()
                self.show_context_menu(pos[0], pos[1])
            elif lparam == win32con.WM_LBUTTONDBLCLK:
                self.open_backup_folder()
        elif msg == win32con.WM_COMMAND:
            cmd = wparam & 0xFFFF
            if cmd == 1000:
                self.open_config_file()
            elif cmd == 1001:
                self.exit_app()
            elif cmd == 1002:
                self.open_backup_folder()
            elif cmd == 1003:
                self.open_log_folder()
            elif cmd == 1004:
                self.toggle_fallback_monitor()
            elif cmd == 2001:
                self.show_about_dialog()
        elif msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
            return 0
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
    
    def exit_app(self):
        self.logger.info("正在退出程序...")
        
        self.running = False
        
        if self.monitor:
            self.logger.info("正在停止监控...")
            self.monitor.stop_monitoring()
        
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (self.hwnd, 0))
            self.logger.debug("托盘图标已删除")
        except:
            pass
        
        win32gui.PostQuitMessage(0)
        
        try:
            win32gui.DestroyWindow(self.hwnd)
        except:
            pass
        
        self.logger.info("程序已退出")
    
    def start_monitoring(self):
        self.monitor = PPTMonitor()
        self.monitor_thread = threading.Thread(target=self.monitor.start_monitoring, daemon=True)
        self.monitor_thread.start()
        self.logger.info("监控线程已启动")
        
        def update_tooltip_loop():
            while self.running:
                time.sleep(5)
                self.update_tray_tooltip()
        
        threading.Thread(target=update_tooltip_loop, daemon=True).start()
    
    def run(self):
        self.start_monitoring()
        self.logger.info("PPT监控程序已启动，托盘图标已显示")
        
        self.update_tray_tooltip()
        
        while self.running:
            try:
                win32gui.PumpMessages()
            except KeyboardInterrupt:
                self.logger.info("收到键盘中断信号")
                self.exit_app()
                break
        
        self.logger.info("应用程序主循环已结束")


def main():
    single_instance = SingleInstance()
    if not single_instance.is_first():
        logger = Logger()
        logger.info("程序已在运行，尝试激活现有窗口...")
        single_instance.bring_to_front()
        return
    
    logger = Logger()
    logger.info("=" * 60)
    logger.info("PowerPoint PPT文件备份监控程序启动（事件驱动模式 - 进程ID缓存版）...")
    
    config_manager = ConfigManager()
    backup_dir = Path(config_manager.get_backup_dir())
    
    try:
        backup_dir.mkdir(exist_ok=True)
        logger.info(f"基础备份目录已准备: {backup_dir}")
        logger.info(f"最大保留天数: {config_manager.get_max_retention_days()}天")
        logger.info(f"最小文件大小: {config_manager.get_min_file_size_kb()}KB")
        logger.info(f"扫描间隔: {config_manager.get_scan_interval()}秒")
        logger.info(f"后备进程监控: {'启用' if config_manager.get_enable_fallback_monitor() else '禁用'}")
        logger.info(f"状态保存间隔: {STATE_SAVE_INTERVAL}秒")
        logger.info(f"非移动设备日志: {'启用' if config_manager.get_log_non_removable_events() else '禁用'}")
    except Exception as e:
        logger.error(f"无法创建备份目录 {backup_dir}: {str(e)}")
        return

    try:
        import psutil
    except ImportError:
        logger.critical("缺失关键依赖 'psutil'。请先运行 'pip install psutil'。")
        sys.exit(1)
    
    app = SystemTrayApp()
    
    try:
        app.run()
    except KeyboardInterrupt:
        logger.info("程序被用户中断")
    except Exception as e:
        logger.exception(f"程序运行出错: {str(e)}")
    
    logger.info("程序已完全退出")
    sys.exit(0)


if __name__ == "__main__":
    main()