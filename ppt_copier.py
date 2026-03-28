import os
import sys
import time
import shutil
import threading
from pathlib import Path
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

# 导入Win32 API常量和函数
from ctypes import windll, c_ulong, byref
from ctypes.wintypes import HWND, UINT, WPARAM, LPARAM, HICON

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
        # 获取程序所在目录
        if getattr(sys, 'frozen', False):
            # 如果是打包后的exe
            base_dir = os.path.dirname(sys.executable)
        else:
            # 如果是脚本运行
            base_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 创建logs目录
        self.log_dir = os.path.join(base_dir, 'logs')
        os.makedirs(self.log_dir, exist_ok=True)
        
        # 日志文件路径
        log_file = os.path.join(self.log_dir, f'ppt_monitor_{datetime.now().strftime("%Y%m%d")}.log')
        
        # 创建logger
        self.logger = logging.getLogger('PPTMonitor')
        self.logger.setLevel(logging.DEBUG)
        
        # 避免重复添加handler
        if self.logger.handlers:
            return
        
        # 文件处理器（按大小滚动，最大10MB，保留5个备份）
        file_handler = RotatingFileHandler(
            log_file, 
            maxBytes=10*1024*1024,  # 10MB
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setLevel(logging.DEBUG)
        
        # 控制台处理器（只输出INFO及以上级别）
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        
        # 设置格式
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        file_handler.setFormatter(formatter)
        console_handler.setFormatter(formatter)
        
        # 添加处理器
        self.logger.addHandler(file_handler)
        self.logger.addHandler(console_handler)
    
    def debug(self, message):
        """调试日志"""
        self.logger.debug(message)
    
    def info(self, message):
        """信息日志"""
        self.logger.info(message)
    
    def warning(self, message):
        """警告日志"""
        self.logger.warning(message)
    
    def error(self, message):
        """错误日志"""
        self.logger.error(message)
    
    def exception(self, message):
        """异常日志（包含堆栈信息）"""
        self.logger.exception(message)


# 修改 ConfigManager 类，使用日志
class ConfigManager:
    def __init__(self, config_file="ppt_monitor.ini"):
        self.config_file = config_file
        self.config = configparser.ConfigParser()
        self.logger = Logger()  # 添加日志
        self.load_config()
    
    def load_config(self):
        """加载配置文件"""
        if os.path.exists(self.config_file):
            # 尝试多种编码方式读取配置文件
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
        """创建默认配置"""
        self.config['General'] = {
            'backup_dir': 'D:/Backup/Desktop',
            'max_retention_days': '30'
        }
        self.save_config()
        self.logger.info("已创建默认配置文件")
    
    def save_config(self):
        """保存配置文件"""
        try:
            with open(self.config_file, 'w', encoding='utf-8-sig') as configfile:
                self.config.write(configfile)
            self.logger.debug(f"配置文件已保存: {self.config_file}")
        except Exception as e:
            self.logger.error(f"保存配置文件失败: {e}")
    
    def get_backup_dir(self):
        """获取备份目录"""
        return self.config.get('General', 'backup_dir', fallback='D:/Backup/Desktop')
    
    def get_max_retention_days(self):
        """获取最大保留天数"""
        try:
            return int(self.config.get('General', 'max_retention_days', fallback='30'))
        except ValueError:
            return 30
    
    def set_backup_dir(self, backup_dir):
        """设置备份目录"""
        if not self.config.has_section('General'):
            self.config['General'] = {}
        self.config['General']['backup_dir'] = backup_dir
        self.save_config()
        self.logger.info(f"备份目录已更新: {backup_dir}")
    
    def set_max_retention_days(self, days):
        """设置最大保留天数"""
        if not self.config.has_section('General'):
            self.config['General'] = {}
        self.config['General']['max_retention_days'] = str(days)
        self.save_config()
        self.logger.info(f"最大保留天数已更新: {days}")


# 修改 PersistentFileManager 类
class PersistentFileManager:
    def __init__(self, base_backup_dir):
        self.base_backup_dir = base_backup_dir
        self.state_file = os.path.join(base_backup_dir, 'monitor_state.json')
        self.today = datetime.now().strftime("%Y-%m-%d")
        self.processed_files = {}
        self.lock = threading.Lock()
        self.logger = Logger()  # 添加日志
        self.load_state()
    
    def load_state(self):
        """从文件加载状态"""
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
        """立即保存状态到文件（线程安全）"""
        with self.lock:
            data = {
                'date': self.today,
                'processed_files': self.processed_files
            }
            try:
                with open(self.state_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self.logger.debug(f"状态已保存，今日已处理 {len(self.processed_files)} 个文件")
            except Exception as e:
                self.logger.error(f"保存状态文件失败: {str(e)}")
    
    def add_processed_file(self, file_path, mtime):
        """添加已处理的文件并立即保存"""
        with self.lock:
            self.processed_files[file_path] = mtime
        self.save_state_immediately()
        self.logger.debug(f"添加已处理文件: {file_path}")
    
    def is_already_processed(self, file_path):
        """检查文件是否已处理（线程安全）"""
        with self.lock:
            return file_path in self.processed_files
    
    def get_file_mtime(self, file_path):
        """获取文件的修改时间（线程安全）"""
        with self.lock:
            return self.processed_files.get(file_path)
    
    def cleanup_old_state(self):
        """清理旧的状态数据"""
        current_date = datetime.now().strftime("%Y-%m-%d")
        if current_date != self.today:
            with self.lock:
                self.processed_files = {}
            self.today = current_date
            self.save_state_immediately()
            self.logger.info(f"跨天重置状态，新日期: {self.today}")


# 修改 WindowsEventMonitor 类
class WindowsEventMonitor:
    def __init__(self, callback):
        self.callback = callback
        self.running = True
        self.removable_drives = set()
        self.logger = Logger()  # 添加日志
    
    def get_removable_drives(self):
        """获取当前可移动驱动器"""
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
    
    def monitor_device_events(self):
        """监控设备插入/拔出事件"""
        old_drives = self.get_removable_drives()
        self.logger.info(f"初始可移动驱动器: {[f'{d}:' for d in old_drives]}")
        
        while self.running:
            time.sleep(1)
            new_drives = self.get_removable_drives()
            
            if new_drives != old_drives:
                added = new_drives - old_drives
                removed = old_drives - new_drives
                
                if added:
                    self.logger.info(f"检测到新插入的驱动器: {['{}:'.format(d) for d in added]}")
                    self.callback('device_inserted', list(added))
                
                if removed:
                    self.logger.info(f"检测到移除的驱动器: {['{}:'.format(d) for d in removed]}")
                    self.callback('device_removed', list(removed))
                
                old_drives = new_drives


# 修改 PPTMonitor 类
class PPTMonitor:
    def __init__(self):
        self.logger = Logger()  # 添加日志
        self.logger.info("初始化PPT监控器")
        
        # 加载配置
        self.config_manager = ConfigManager()
        
        # 目标备份目录
        self.base_backup_dir = self.config_manager.get_backup_dir()
        # 创建基础备份目录
        os.makedirs(self.base_backup_dir, exist_ok=True)
        
        # 最大保留天数
        self.max_retention_days = self.config_manager.get_max_retention_days()
        
        # 持久化文件管理器
        self.file_manager = PersistentFileManager(self.base_backup_dir)
        
        # PPT文件扩展名
        self.ppt_extensions = {'.ppt', '.pptx', '.pps', '.ppsx'}
        
        # 设备事件监控器
        self.device_monitor = WindowsEventMonitor(self.on_device_event)
        
        # 当前连接的可移动驱动器
        self.connected_drives = set()
        
        self.logger.info(f"备份目录: {self.base_backup_dir}")
        self.logger.info(f"最大保留天数: {self.max_retention_days}")
    
    def cleanup_old_backups(self):
        """清理超过保留天数的旧备份"""
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
        """获取按日期组织的备份目录路径"""
        today = datetime.now().strftime("%Y-%m-%d")
        date_folder = os.path.join(self.base_backup_dir, today)
        os.makedirs(date_folder, exist_ok=True)
        return date_folder
    
    def is_removable_drive(self, path):
        """检查路径是否在可移动驱动器上"""
        if not path or not isinstance(path, str):
            return False
        drive_letter = os.path.splitdrive(path)[0].upper()
        if not drive_letter.endswith(':'):
            return False
        
        try:
            drive_type = win32file.GetDriveType(f"{drive_letter}\\")
            return drive_type in [win32file.DRIVE_REMOVABLE, win32file.DRIVE_CDROM]
        except:
            return False
    
    def is_valid_ppt_file(self, file_path):
        """检查是否为有效的PPT文件（排除临时文件）"""
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
        """检查文件是否已更改"""
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
        """复制PPT文件到按日期组织的备份目录，保留原始文件名"""
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
            
            shutil.copy2(source_path, dest_path)
            self.logger.info(f"已备份PPT文件: {source_path} -> {dest_path}")
            
            try:
                current_mtime = os.path.getmtime(source_path)
                self.file_manager.add_processed_file(source_path, current_mtime)
            except Exception as e:
                self.logger.error(f"更新文件记录失败: {str(e)}")
            
            return True
        except Exception as e:
            self.logger.error(f"备份失败 {source_path}: {str(e)}")
            return False
    
    def get_powerpoint_processes_with_files(self):
        """获取PowerPoint进程及其打开的文件"""
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
    
    def monitor_powerpoint_via_com(self):
        """通过COM接口监控PowerPoint"""
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
                                    
                                    if (self.is_valid_ppt_file(full_path) and 
                                        self.is_removable_drive(full_path) and 
                                        self.has_file_changed(full_path)):
                                        
                                        self.logger.info(f"通过COM检测到PowerPoint打开移动设备上的PPT文件: {full_path}")
                                        self.copy_ppt_file(full_path)
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
                
            pythoncom.CoUninitialize()
            
        except Exception as e:
            self.logger.error(f"初始化COM时出错: {str(e)}")
    
    def monitor_powerpoint_via_process(self):
        """通过进程监控PowerPoint"""
        try:
            ppt_processes = self.get_powerpoint_processes_with_files()
            
            for pid, file_path in ppt_processes:
                if (self.is_valid_ppt_file(file_path) and 
                    self.is_removable_drive(file_path) and 
                    self.has_file_changed(file_path)):
                    
                    self.logger.info(f"通过进程监控检测到PowerPoint打开移动设备上的PPT文件: {file_path} (PID: {pid})")
                    self.copy_ppt_file(file_path)
                    
        except Exception as e:
            self.logger.error(f"通过进程监控PowerPoint时出错: {str(e)}")
    
    def on_device_event(self, event_type, drives):
        """设备事件回调函数"""
        if event_type == 'device_inserted':
            self.connected_drives.update(drives)
            self.logger.info(f"设备插入事件: {drives}")
            self.logger.info(f"当前连接的可移动驱动器: {sorted(list(self.connected_drives))}")
        elif event_type == 'device_removed':
            for drive in drives:
                if drive in self.connected_drives:
                    self.connected_drives.remove(drive)
            self.logger.info(f"设备移除事件: {drives}")
            self.logger.info(f"当前连接的可移动驱动器: {sorted(list(self.connected_drives))}")
    
    def get_connected_drives(self):
        """获取当前连接的驱动器列表"""
        return sorted(list(self.connected_drives))
    
    def start_monitoring(self):
        """开始监控"""
        self.logger.info("=" * 50)
        self.logger.info("开始监控PowerPoint打开的移动存储设备上的PPT文件...")
        self.logger.info(f"基础备份目录: {self.base_backup_dir}")
        self.logger.info(f"最大保留天数: {self.max_retention_days}")
        self.logger.info(f"今日已处理文件数量: {len(self.file_manager.processed_files)}")
        
        initial_drives = self.device_monitor.get_removable_drives()
        self.connected_drives = initial_drives
        self.logger.info(f"初始可移动驱动器: {sorted(list(initial_drives))}")
        
        device_thread = threading.Thread(target=self.device_monitor.monitor_device_events, daemon=True)
        device_thread.start()
        
        while True:
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
                self.device_monitor.running = False
                break
            except Exception as e:
                self.logger.error(f"监控过程中出现错误: {str(e)}", exc_info=True)
                time.sleep(2)


# 修改 SystemTrayApp 类
class SystemTrayApp:
    def __init__(self):
        self.logger = Logger()  # 添加日志
        self.WM_NOTIFY_CALLBACK = win32gui.RegisterWindowMessage("NotifyCallback")
        self.hwnd = None
        self.monitor = None
        self.monitor_thread = None
        self.config_manager = ConfigManager()
        self.running = True
        
        self.logger.info("初始化系统托盘应用")
        
        # 注册窗口类
        wc = win32gui.WNDCLASS()
        wc.lpszClassName = "PPTMonitorTrayClass"
        wc.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wc.hbrBackground = win32con.COLOR_WINDOW
        wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpfnWndProc = self.wnd_proc
        self.class_atom = win32gui.RegisterClass(wc)
        
        # 创建窗口
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
        
        # 创建托盘图标
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
        """打开配置文件"""
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
        """显示右键菜单"""
        menu = win32gui.CreatePopupMenu()
        
        status_text = "无连接设备"
        retention_text = f"保留天数: {self.config_manager.get_max_retention_days()}天"
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
        win32gui.AppendMenu(menu, win32con.MF_GRAYED, 0, f"今日处理: {processed_count}个文件")
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(menu, win32con.MF_STRING, 1000, "打开配置文件")
        win32gui.AppendMenu(menu, win32con.MF_STRING, 1001, "退出")
        
        win32gui.SetForegroundWindow(self.hwnd)
        win32gui.TrackPopupMenu(menu, win32con.TPM_LEFTALIGN, x, y, 0, self.hwnd, None)
        win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)
    
    def wnd_proc(self, hwnd, msg, wparam, lparam):
        """窗口过程函数"""
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
        """退出应用程序"""
        self.logger.info("正在退出程序...")
        
        self.running = False
        
        if self.monitor:
            self.logger.info("正在保存状态...")
            self.monitor.file_manager.save_state_immediately()
            self.monitor.device_monitor.running = False
        
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
        """启动监控"""
        self.monitor = PPTMonitor()
        self.monitor_thread = threading.Thread(target=self.monitor.start_monitoring, daemon=True)
        self.monitor_thread.start()
        self.logger.info("监控线程已启动")
    
    def run(self):
        """运行应用程序"""
        self.start_monitoring()
        self.logger.info("PPT监控程序已启动，托盘图标已显示")
        
        while self.running:
            try:
                win32gui.PumpMessages()
                time.sleep(0.1)
            except KeyboardInterrupt:
                self.logger.info("收到键盘中断信号")
                self.exit_app()
                break
        
        self.logger.info("应用程序主循环已结束")


def main():
    """主函数"""
    # 初始化日志
    logger = Logger()
    logger.info("=" * 60)
    logger.info("PowerPoint PPT文件备份监控程序启动...")
    logger.info("仅监控PowerPoint打开的移动存储设备上的PPT文件")
    
    # 加载配置
    config_manager = ConfigManager()
    backup_dir = config_manager.get_backup_dir()
    
    # 检查并创建备份目录
    try:
        os.makedirs(backup_dir, exist_ok=True)
        logger.info(f"基础备份目录已准备: {backup_dir}")
        logger.info(f"最大保留天数: {config_manager.get_max_retention_days()}天")
    except Exception as e:
        logger.error(f"无法创建备份目录 {backup_dir}: {str(e)}")
        return
    
    # 安装psutil依赖
    try:
        import psutil
        logger.info("psutil模块已加载")
    except ImportError:
        logger.warning("psutil未安装，正在安装...")
        try:
            import subprocess
            subprocess.check_call([sys.executable, "-m", "pip", "install", "psutil"])
            import psutil
            logger.info("psutil安装成功")
        except Exception as e:
            logger.error(f"无法安装psutil: {str(e)}")
            return
    
    # 创建托盘应用程序
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
