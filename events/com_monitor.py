"""PowerPoint COM 事件监听器"""
import threading
import time
import pythoncom
import pywintypes
from win32com.client import Dispatch, GetObject, DispatchWithEvents
import psutil
from events.com_events import PowerPointEvents
from core.logger import Logger
from utils.constants import QUITTING_HRESULTS


class PowerPointEventMonitor:
    """PowerPoint COM 事件监听器"""
    def __init__(self, monitor):
        self.monitor = monitor
        self.running = False
        self.event_thread = None
        self.powerpoint = None
        self.logger = Logger()
        self.event_handler = None
        self.reconnect_delay = 3
    
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
    
    def _is_powerpoint_process_running(self):
        if self.monitor and hasattr(self.monitor, 'process_cache'):
            return len(self.monitor.process_cache.get_powerpoint_process_ids()) > 0
        for proc in psutil.process_iter(['name']):
            if proc.info['name'] and proc.info['name'].lower().startswith('powerpnt'):
                return True
        return False
    
    def _is_quitting_error(self, e):
        return hasattr(e, 'hresult') and e.hresult in QUITTING_HRESULTS
    
    def _try_connect_powerpoint(self):
        if not self._is_powerpoint_process_running():
            return None
        try:
            powerpoint = Dispatch('PowerPoint.Application')
            if powerpoint:
                _ = powerpoint.Name
                return powerpoint
        except:
            pass
        try:
            powerpoint = GetObject(Class='PowerPoint.Application')
            if powerpoint:
                _ = powerpoint.Name
                return powerpoint
        except:
            pass
        return None
    
    def _process_existing_presentations(self, powerpoint):
        try:
            presentations_count = powerpoint.Presentations.Count
            for i in range(presentations_count):
                try:
                    pres = powerpoint.Presentations.Item(i + 1)
                    file_path = pres.FullName if hasattr(pres, 'FullName') else None
                    if file_path and self.monitor and self.monitor.is_removable_drive(file_path):
                        self.logger.info(f"检测到已打开的演示文稿: {file_path}")
                        def delayed_process(fp=file_path):
                            time.sleep(1)
                            if self.monitor.is_valid_ppt_file_for_backup(fp):
                                self.monitor.process_ppt_file(fp, source="已打开文件")
                        threading.Thread(target=delayed_process, daemon=True).start()
                except:
                    continue
        except:
            pass
    
    def _run_com_loop(self):
        pythoncom.CoInitialize()
        self.logger.info("COM事件循环已启动")
        last_pp_check = 0
        
        while self.running:
            try:
                if self.powerpoint is None:
                    if self._is_powerpoint_process_running():
                        self.powerpoint = self._try_connect_powerpoint()
                        if self.powerpoint:
                            self.logger.info("已成功连接到PowerPoint")
                            self.event_handler = DispatchWithEvents(self.powerpoint, PowerPointEvents)
                            self.event_handler.set_monitor(self.monitor)
                            self._process_existing_presentations(self.powerpoint)
                        else:
                            time.sleep(self.reconnect_delay)
                    else:
                        time.sleep(5)
                elif self.powerpoint:
                    try:
                        if time.time() - last_pp_check > 10:
                            last_pp_check = time.time()
                            if not self._is_powerpoint_process_running():
                                self.powerpoint = None
                                continue
                            _ = self.powerpoint.Name
                        
                        pythoncom.PumpWaitingMessages()
                        time.sleep(0.05)
                    except pywintypes.com_error as e:
                        if self._is_quitting_error(e):
                            self.powerpoint = None
                        else:
                            self.powerpoint = None
                            time.sleep(self.reconnect_delay)
            except Exception as e:
                self.logger.error(f"COM事件循环出错: {e}")
                time.sleep(1)
        
        pythoncom.CoUninitialize()
        self.logger.info("COM事件循环已结束")