"""Windows设备变化监听器"""
import threading
import time
import win32gui
import win32api
import win32file
import win32con
from ctypes import windll, byref, sizeof
from utils.constants import (
    WM_DEVICECHANGE, DBT_DEVICEARRIVAL, DBT_DEVICEREMOVECOMPLETE,
    DBT_DEVTYP_VOLUME, DEV_BROADCAST_HDR, DEV_BROADCAST_VOLUME,
    DEVICE_NOTIFY_WINDOW_HANDLE
)
from core.logger import Logger


class WindowsDeviceMonitor:
    """Windows设备变化监听器"""
    def __init__(self, callback):
        self.callback = callback
        self.running = False
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
                    # 直接使用 lparam 作为地址
                    if lparam:
                        hdr = DEV_BROADCAST_HDR.from_address(lparam)
                        if hdr.dbch_devicetype == DBT_DEVTYP_VOLUME:
                            vol_struct = DEV_BROADCAST_VOLUME.from_address(lparam)
                            unit_mask = vol_struct.dbcv_unitmask
                            
                            detected_drives = set()
                            for i in range(26):
                                if unit_mask & (1 << i):
                                    detected_drives.add(chr(ord('A') + i))

                            if wparam == DBT_DEVICEARRIVAL:
                                self.logger.info(f"检测到新插入的驱动器: {['{}:'.format(d) for d in detected_drives]}")
                                self.callback('device_inserted', list(detected_drives))
                            else:
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
            0, self.class_atom, "PPTMonitorDeviceListenerWindow", 0,
            0, 0, 0, 0, 0, 0, win32api.GetModuleHandle(None), None
        )

        if not self.hwnd:
            self.logger.error("无法创建监听窗口")
            return

        # 获取初始的可移动驱动器
        initial_drives = self.get_removable_drives()
        if initial_drives:
            self.logger.info(f"初始可移动驱动器: {[f'{d}:' for d in initial_drives]}")
            self.callback('device_inserted', list(initial_drives))

        # 注册设备通知
        dbv = DEV_BROADCAST_VOLUME()
        dbv.dbcv_size = sizeof(DEV_BROADCAST_VOLUME)
        dbv.dbcv_devicetype = DBT_DEVTYP_VOLUME
        dbv.dbcv_reserved = 0
        dbv.dbcv_unitmask = 0
        dbv.dbcv_flags = 0
        
        dev_broadcast_handle = windll.user32.RegisterDeviceNotificationW(
            self.hwnd, byref(dbv), DEVICE_NOTIFY_WINDOW_HANDLE
        )
        
        if not dev_broadcast_handle:
            self.logger.warning(f"注册设备通知失败，将使用轮询方式检测设备")
        else:
            self.logger.debug("已注册设备通知")

        # 消息循环
        while self.running:
            try:
                # 使用 PeekMessage 避免阻塞
                win32gui.PumpWaitingMessages()
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