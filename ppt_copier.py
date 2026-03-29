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
from win32com.client import GetObject
import subprocess
import psutil
from datetime import datetime, timedelta
import win32api
import win32gui
import win32event
import winerror
import pywintypes
import struct
import configparser
import webbrowser
import logging
from logging.handlers import RotatingFileHandler
from ctypes import windll, c_ulong, byref, POINTER, Structure, WINFUNCTYPE, sizeof
from ctypes.wintypes import HWND, UINT, WPARAM, LPARAM, HICON, DWORD, HANDLE, BOOL
from win32con import PM_REMOVE, DEVICE_NOTIFY_WINDOW_HANDLE
from pathlib import Path

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
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))
        
        self.log_dir = os.path.join(base_dir, 'logs')
        os.makedirs(self.log_dir, exist_ok=True)
        
        log_file = os.path.join(self.log_dir, f'ppt_monitor_{datetime.now().strftime("%Y%m%d")}.log')
        
        self.logger = logging.getLogger('PPTMonitor')
        self.logger.setLevel(logging.DEBUG)
        
        if self.logger.handlers:
            return
        
        file_handler = RotatingFileHandler(log_file, maxBytes=10*1024*1024, backupCount=5, encoding='utf-8')
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
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self.logger = Logger()
        self.load_config()
    
    def load_config(self):
        if os.path.exists(self.config_file):
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
            'scan_interval_seconds': '30'  # 新增：扫描间隔配置
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
    
    def get_scan_interval(self):
        """获取设备扫描间隔（秒）"""
        try:
            return int(self.config.get('General', 'scan_interval_seconds', fallback='30'))
        except ValueError:
            return 30
    
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
    
    def set_scan_interval(self, seconds):
        """设置扫描间隔"""
        if not self.config.has_section('General'):
            self.config['General'] = {}
        self.config['General']['scan_interval_seconds'] = str(seconds)
        self.save_config()
        self.logger.info(f"扫描间隔已更新: {seconds}秒")


class PersistentFileManager:
    def __init__(self, base_backup_dir):
        self.base_backup_dir = base_backup_dir
        self.state_file = os.path.join(base_backup_dir, 'monitor_state.json')
        self.today = datetime.now().strftime("%Y-%m-%d")
        self.processed_files = {}
        self.lock = threading.Lock()
        self.logger = Logger()
        self.load_state()
    
    def load_state(self):
        if os.path.exists(self.state_file):
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
            self.logger.debug("状态文件不存在，使用空状态")
    
    def save_state_immediately(self):
        with self.lock:
            data = {'date': self.today, 'processed_files': self.processed_files}
            try:
                with open(self.state_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self.logger.debug(f"状态已保存，今日已处理 {len(self.processed_files)} 个文件")
            except Exception as e:
                self.logger.error(f"保存状态文件失败: {str(e)}")
    
    def add_processed_file(self, file_path, mtime):
        with self.lock:
            self.processed_files[file_path] = mtime
        self.save_state_immediately()
        self.logger.debug(f"添加已处理文件: {file_path}")
    
    def is_already_processed(self, file_path):
        with self.lock:
            return file_path in self.processed_files
    
    def get_file_mtime(self, file_path):
        with self.lock:
            return self.processed_files.get(file_path)
    
    def cleanup_old_state(self):
        current_date = datetime.now().strftime("%Y-%m-%d")
        if current_date != self.today:
            with self.lock:
                self.processed_files = {}
            self.today = current_date
            self.save_state_immediately()
            self.logger.info(f"跨天重置状态，新日期: {self.today}")


# --- 重写 WindowsEventMonitor 类 ---
class WindowsEventMonitor:
    def __init__(self, callback):
        self.callback = callback
        self.running = False
        self.removable_drives = set()
        self.logger = Logger()
        self.hwnd = None
        self.event_thread = None

        # 注册窗口类
        wc = win32gui.WNDCLASS()
        wc.lpszClassName = "PPTMonitorDeviceListener"
        wc.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wc.hbrBackground = win32con.COLOR_WINDOW
        wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpfnWndProc = self.wnd_proc
        self.class_atom = win32gui.RegisterClass(wc)

    def wnd_proc(self, hwnd, msg, wparam, lparam):
        """窗口过程函数，处理 WM_DEVICECHANGE 消息"""
        if msg == WM_DEVICECHANGE:
            if wparam in (DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE):
                try:
                    # 解析 lParam 指向的 DEV_BROADCAST_HDR 结构
                    hdr_ptr = LPARAM(lparam).value
                    if hdr_ptr:
                        hdr = DEV_BROADCAST_HDR.from_address(hdr_ptr)
                        if hdr.dbch_devicetype == DBT_DEVTYP_VOLUME:
                            vol_struct = DEV_BROADCAST_VOLUME.from_address(hdr_ptr)
                            unit_mask = vol_struct.dbcv_unitmask
                            
                            # 将位掩码转换为驱动器字母集合
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
        """启动事件监听线程"""
        if self.running:
            self.logger.warning("事件监听器已在运行中")
            return
        self.running = True
        self.event_thread = threading.Thread(target=self._run_message_loop, daemon=True)
        self.event_thread.start()
        self.logger.info("设备事件监听器已启动")

    def stop_listening(self):
        """停止事件监听"""
        self.running = False
        if self.hwnd:
            try:
                win32gui.PostMessage(self.hwnd, win32con.WM_QUIT, 0, 0)
            except:
                pass
        if self.event_thread:
            self.event_thread.join(timeout=2)
        self.logger.info("设备事件监听器已停止")

    def _run_message_loop(self):
        """在后台线程中运行的消息循环"""
        # 创建隐藏窗口
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

        # 初始获取当前连接的驱动器
        initial_drives = self.get_removable_drives()
        self.logger.info(f"初始可移动驱动器: {[f'{d}:' for d in initial_drives]}")
        if initial_drives:
            self.callback('device_inserted', list(initial_drives))

        # 注册设备通知
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

        # 定义 MSG 结构体
        from ctypes import wintypes
        class MSG(Structure):
            _fields_ = [
                ('hwnd', wintypes.HWND),
                ('message', wintypes.UINT),
                ('wParam', wintypes.WPARAM),
                ('lParam', wintypes.LPARAM),
                ('time', wintypes.DWORD),
                ('pt', wintypes.POINT),
                ('lPrivate', wintypes.DWORD),
            ]

        # 消息循环 - 使用 GetMessage 实现真正的阻塞等待
        msg = MSG()
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

        # 清理
        if dev_broadcast_handle:
            windll.user32.UnregisterDeviceNotification(dev_broadcast_handle)
        try:
            win32gui.DestroyWindow(self.hwnd)
        except:
            pass
        self.hwnd = None

    def get_removable_drives(self):
        """获取当前可移动驱动器 (用于初始化)"""
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


# --- PPTMonitor 类 ---
class PPTMonitor:
    def __init__(self):
        self.logger = Logger()
        self.logger.info("初始化PPT监控器")
        
        self.config_manager = ConfigManager()
        self.base_backup_dir = self.config_manager.get_backup_dir()
        os.makedirs(self.base_backup_dir, exist_ok=True)
        self.max_retention_days = self.config_manager.get_max_retention_days()
        self.scan_interval = self.config_manager.get_scan_interval()  # 获取扫描间隔
        self.file_manager = PersistentFileManager(self.base_backup_dir)
        self.ppt_extensions = PPT_EXTENSIONS
        # 设备事件监控器现在是事件驱动的
        self.device_monitor = WindowsEventMonitor(self.on_device_event)
        self.connected_drives = set()
        self.current_removable_drives_cache = set()  # 缓存优化
        
        # 新增：用于定期扫描的线程和标志
        self.scan_thread = None
        self.scan_thread_running = False
        
        self.logger.info(f"备份目录: {self.base_backup_dir}")
        self.logger.info(f"最大保留天数: {self.max_retention_days}")
        self.logger.info(f"设备扫描间隔: {self.scan_interval}秒")

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
        date_folder = os.path.join(self.base_backup_dir, today)
        os.makedirs(date_folder, exist_ok=True)
        return date_folder

    def is_removable_drive(self, path):
        """检查路径是否在可移动驱动器上 (使用缓存)"""
        if not path or not isinstance(path, str):
            return False
        drive_letter = os.path.splitdrive(path)[0].upper().rstrip(':')
        return drive_letter in self.current_removable_drives_cache

    def is_valid_ppt_file(self, file_path):
        if not file_path:
            return False
        filename = os.path.basename(file_path)
        if filename.startswith('~$'):
            return False
        try:
            attrs = win32file.GetFileAttributes(file_path)
            if attrs & win32con.FILE_ATTRIBUTE_HIDDEN:
                return False
        except:
            pass
        if Path(file_path).suffix.lower() not in self.ppt_extensions:
            return False
        return True

    def has_file_changed(self, file_path):
        try:
            current_mtime = os.path.getmtime(file_path)
            self.file_manager.cleanup_old_state()
            
            if self.file_manager.is_already_processed(file_path):
                stored_mtime = self.file_manager.get_file_mtime(file_path)
                if stored_mtime is not None:
                    if abs(current_mtime - stored_mtime) < 1:
                        return False
                    else:
                        self.file_manager.add_processed_file(file_path, current_mtime)
                        self.logger.debug(f"文件已更改: {file_path}")
                        return True
                else:
                    return True
            else:
                self.file_manager.add_processed_file(file_path, current_mtime)
                self.logger.debug(f"新文件: {file_path}")
                return True
        except Exception as e:
            self.logger.error(f"检查文件更改失败 {file_path}: {e}")
            return True

    def copy_ppt_file(self, source_path):
        try:
            if not source_path or not os.path.exists(source_path):
                self.logger.warning(f"源文件不存在或路径无效: {source_path}")
                return False

            date_folder = self.get_date_folder_path()
            filename = os.path.basename(source_path)
            dest_path = os.path.join(date_folder, filename)

            counter = 1
            original_dest_path = dest_path
            while os.path.exists(dest_path):
                name, ext = os.path.splitext(original_dest_path)
                dest_path = f"{name}_{counter}{ext}"
                counter += 1

            try:
                shutil.copy2(source_path, dest_path)
                self.logger.info(f"已备份PPT文件: {source_path} -> {dest_path}")
            except (IOError, OSError) as copy_err:
                self.logger.error(f"备份文件时I/O错误 {source_path}: {copy_err}")
                return False

            try:
                current_mtime = os.path.getmtime(source_path)
                self.file_manager.add_processed_file(source_path, current_mtime)
            except (OSError, IOError) as mtime_err:
                self.logger.error(f"无法获取源文件修改时间或更新记录失败 {source_path}: {mtime_err}")

            return True
        except Exception as e:
            self.logger.error(f"备份失败 {source_path} (未知错误): {str(e)}", exc_info=True)
            return False

    def get_powerpoint_processes_with_files(self):
        ppt_files = []
        try:
            for proc in psutil.process_iter(['pid', 'name', 'open_files']):
                if proc.info['name'].lower().startswith('powerpnt'):
                    try:
                        open_files = proc.info['open_files']
                        if open_files:
                            for file_info in open_files:
                                file_path = file_info.path
                                if self.is_valid_ppt_file(file_path):
                                    ppt_files.append((proc.pid, file_path))
                    except (psutil.AccessDenied, psutil.NoSuchProcess):
                        continue
        except Exception as e:
            self.logger.error(f"获取PowerPoint进程信息时出错: {str(e)}")
        return ppt_files

    def _process_potential_ppt_file(self, file_path):
        """统一处理可能的PPT文件的逻辑"""
        if (self.is_valid_ppt_file(file_path) and 
            self.is_removable_drive(file_path) and 
            self.has_file_changed(file_path)):
            
            self.logger.info(f"检测到PowerPoint打开移动设备上的PPT文件: {file_path}")
            self.copy_ppt_file(file_path)

    def monitor_powerpoint_via_com(self):
        try:
            pythoncom.CoInitialize()
            try:
                powerpoint = GetObject(None, 'PowerPoint.Application')
                if powerpoint and hasattr(powerpoint, 'Presentations'):
                    presentations_count = 0
                    try:
                        presentations_count = powerpoint.Presentations.Count
                    except:
                        pass
                    
                    if presentations_count > 0:
                        for i in range(presentations_count):
                            try:
                                presentation = powerpoint.Presentations.Item(i + 1)
                                if hasattr(presentation, 'FullName') and presentation.FullName:
                                    full_path = presentation.FullName
                                    self._process_potential_ppt_file(full_path)
                            except pywintypes.com_error as com_err:
                                if com_err.hresult == -2147188160:
                                    continue
                                elif com_err.hresult == -2147352567:
                                    if "Automation rights are not granted" in str(com_err):
                                        continue
                                    else:
                                        self.logger.debug(f"COM错误: {str(com_err)}")
                                else:
                                    self.logger.debug(f"COM错误: {str(com_err)}")
                                continue
                            except Exception as e:
                                self.logger.error(f"访问Presentation项时出错: {str(e)}")
                                continue
            except pywintypes.com_error:
                pass
            except Exception as e:
                self.logger.error(f"访问PowerPoint COM对象时出错: {str(e)}")
            finally:
                pythoncom.CoUninitialize()
        except Exception as e:
            self.logger.error(f"初始化COM时出错: {str(e)}")

    def monitor_powerpoint_via_process(self):
        try:
            ppt_processes = self.get_powerpoint_processes_with_files()
            for pid, file_path in ppt_processes:
                self._process_potential_ppt_file(file_path)
        except Exception as e:
            self.logger.error(f"通过进程监控PowerPoint时出错: {str(e)}")

    def on_device_event(self, event_type, drives):
        """设备事件回调函数"""
        if event_type == 'device_inserted':
            self.connected_drives.update(drives)
            self.current_removable_drives_cache.update(drives)
            self.logger.info(f"设备插入事件: {drives}")
            self.logger.info(f"当前连接的可移动驱动器: {sorted(list(self.connected_drives))}")
        elif event_type == 'device_removed':
            for drive in drives:
                if drive in self.connected_drives:
                    self.connected_drives.remove(drive)
                if drive in self.current_removable_drives_cache:
                    self.current_removable_drives_cache.remove(drive)
            self.logger.info(f"设备移除事件: {drives}")
            self.logger.info(f"当前连接的可移动驱动器: {sorted(list(self.connected_drives))}")

    def get_connected_drives(self):
        return sorted(list(self.connected_drives))
    
    def _scan_devices_periodically(self):
        """定期扫描设备列表，防止事件遗漏"""
        self.logger.info("设备定期扫描线程已启动")
        
        while self.scan_thread_running:
            try:
                time.sleep(self.scan_interval)
                
                if not self.scan_thread_running:
                    break
                
                # 获取当前实际连接的可移动驱动器
                current_drives = self.device_monitor.get_removable_drives()
                current_drives_set = set(current_drives)
                
                # 与缓存的驱动器列表进行比较
                cached_drives_set = self.connected_drives.copy()
                
                # 找出新增的驱动器（事件遗漏）
                new_drives = current_drives_set - cached_drives_set
                if new_drives:
                    self.logger.warning(f"定期扫描发现遗漏的设备插入事件: {list(new_drives)}")
                    self.on_device_event('device_inserted', list(new_drives))
                
                # 找出移除的驱动器（事件遗漏）
                removed_drives = cached_drives_set - current_drives_set
                if removed_drives:
                    self.logger.warning(f"定期扫描发现遗漏的设备移除事件: {list(removed_drives)}")
                    self.on_device_event('device_removed', list(removed_drives))
                
                # 如果没有变化，记录调试信息（可选，避免日志过多）
                # if not new_drives and not removed_drives and self.logger.logger.isEnabledFor(logging.DEBUG):
                #     self.logger.debug(f"定期扫描完成，当前设备: {sorted(list(current_drives_set))}")
                    
            except Exception as e:
                self.logger.error(f"定期扫描设备时出错: {str(e)}")
        
        self.logger.info("设备定期扫描线程已停止")
    
    def start_scan_thread(self):
        """启动定期扫描线程"""
        if self.scan_thread_running:
            self.logger.warning("定期扫描线程已在运行中")
            return
        
        self.scan_thread_running = True
        self.scan_thread = threading.Thread(target=self._scan_devices_periodically, daemon=True)
        self.scan_thread.start()
        self.logger.info(f"设备定期扫描线程已启动，扫描间隔: {self.scan_interval}秒")
    
    def stop_scan_thread(self):
        """停止定期扫描线程"""
        self.scan_thread_running = False
        if self.scan_thread:
            self.scan_thread.join(timeout=3)
        self.logger.info("设备定期扫描线程已停止")
    
    def start_monitoring(self):
        self.logger.info("=" * 50)
        self.logger.info("开始监控PowerPoint打开的移动存储设备上的PPT文件...")
        self.logger.info(f"基础备份目录: {self.base_backup_dir}")
        self.logger.info(f"最大保留天数: {self.max_retention_days}")
        self.logger.info(f"设备扫描间隔: {self.scan_interval}秒")
        self.logger.info(f"今日已处理文件数量: {len(self.file_manager.processed_files)}")

        # 启动事件监听器
        self.device_monitor.start_listening()
        
        # 启动定期扫描线程，防止事件遗漏
        self.start_scan_thread()
        
        while self.running:
            try:
                current_hour = datetime.now().hour
                if current_hour == 0:
                    self.cleanup_old_backups()
                
                self.file_manager.cleanup_old_state()
                self.monitor_powerpoint_via_com()
                self.monitor_powerpoint_via_process()
                time.sleep(2)
                
            except KeyboardInterrupt:
                self.logger.info("监控已停止")
                self.file_manager.save_state_immediately()
                self.device_monitor.stop_listening()
                self.stop_scan_thread()
                break
            except Exception as e:
                self.logger.error(f"监控过程中出现错误: {str(e)}", exc_info=True)
                time.sleep(2)

    def stop_monitoring(self):
        """停止监控"""
        self.running = False
        self.stop_scan_thread()


# --- 修改 SystemTrayApp 类 ---
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
        icon_path = "./icon.ico"
        try:
            hicon = win32gui.LoadImage(
                0, icon_path, win32con.IMAGE_ICON, 0, 0, 
                win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            )
            self.logger.info("成功加载托盘图标")
        except Exception as e:
            self.logger.warning(f"加载图标失败，使用默认图标: {e}")
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
    
        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        nid = (self.hwnd, 0, flags, self.WM_NOTIFY_CALLBACK, hicon, "PPT文件备份监控")
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
        self.logger.debug("托盘图标已添加")
    
    def open_config_file(self):
        config_file = self.config_manager.config_file
        if os.path.exists(config_file):
            try:
                os.startfile(config_file)
                self.logger.info(f"已打开配置文件: {config_file}")
            except Exception as e:
                self.logger.error(f"打开配置文件失败: {str(e)}")
        else:
            self.logger.warning(f"配置文件不存在: {config_file}")
    
    def show_context_menu(self, x, y):
        menu = win32gui.CreatePopupMenu()
        
        status_text = "无连接设备"
        retention_text = f"保留天数: {self.config_manager.get_max_retention_days()}天"
        scan_interval_text = f"扫描间隔: {self.config_manager.get_scan_interval()}秒"
        processed_count = 0
        if self.monitor:
            connected_drives = self.monitor.get_connected_drives()
            if connected_drives:
                status_text = f"连接设备: {', '.join([f'{d}:' for d in connected_drives])}"
            else:
                status_text = "无连接设备"
            if self.monitor.file_manager:
                processed_count = len(self.monitor.file_manager.processed_files)
        
        win32gui.AppendMenu(menu, win32con.MF_GRAYED, 0, f"状态: {status_text}")
        win32gui.AppendMenu(menu, win32con.MF_GRAYED, 0, f"配置: {retention_text}")
        win32gui.AppendMenu(menu, win32con.MF_GRAYED, 0, f"扫描: {scan_interval_text}")
        win32gui.AppendMenu(menu, win32con.MF_GRAYED, 0, f"今日处理: {processed_count}个文件")
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING, 1000, "打开配置文件")
        win32gui.AppendMenu(menu, win32con.MF_STRING, 1001, "退出")
        
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, x, y, 0, self.hwnd, None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
    
    def wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == self.WM_NOTIFY_CALLBACK:
            if lparam == win32con.WM_RBUTTONUP:
                pos = win32gui.GetCursorPos()
                self.show_context_menu(pos[0], pos[1])
            elif lparam == win32con.WM_LBUTTONDBLCLK:
                pass
        elif msg == win32con.WM_COMMAND:
            if wparam == 1000:
                self.open_config_file()
            elif wparam == 1001:
                self.exit_app()
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
            self.logger.info("正在保存状态...")
            self.monitor.file_manager.save_state_immediately()
            self.monitor.device_monitor.stop_listening()
        
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
        self.monitor.running = True
        self.monitor_thread = threading.Thread(target=self.monitor.start_monitoring, daemon=True)
        self.monitor_thread.start()
        self.logger.info("监控线程已启动")
    
    def run(self):
        self.start_monitoring()
        self.logger.info("PPT监控程序已启动，托盘图标已显示")
        
        while self.running:
            try:
                win32gui.PumpMessages()
            except KeyboardInterrupt:
                self.logger.info("收到键盘中断信号")
                self.exit_app()
                break
        
        self.logger.info("应用程序主循环已结束")


def main():
    logger = Logger()
    logger.info("=" * 60)
    logger.info("PowerPoint PPT文件备份监控程序启动...")
    logger.info("仅监控PowerPoint打开的移动存储设备上的PPT文件")
    
    config_manager = ConfigManager()
    backup_dir = config_manager.get_backup_dir()
    
    try:
        os.makedirs(backup_dir, exist_ok=True)
        logger.info(f"基础备份目录已准备: {backup_dir}")
        logger.info(f"最大保留天数: {config_manager.get_max_retention_days()}天")
        logger.info(f"设备扫描间隔: {config_manager.get_scan_interval()}秒")
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
