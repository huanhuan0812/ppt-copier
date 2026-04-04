"""单实例管理模块"""
import win32event
import win32api
import win32gui
import win32con
import winerror
from core.logger import Logger
from utils.constants import MUTEX_NAME


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