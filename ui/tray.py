"""系统托盘界面模块"""
import threading
import os
import time
from pathlib import Path
import win32gui
import win32api
import win32con
from core.logger import Logger
from core.config import ConfigManager
from core.monitor import PPTMonitor


class SystemTrayApp:
    def __init__(self):
        self.logger = Logger()
        self.WM_NOTIFY_CALLBACK = win32gui.RegisterWindowMessage("NotifyCallback")
        self.hwnd = None
        self.monitor = None
        self.monitor_thread = None
        self.config_manager = ConfigManager()
        self.running = True
        
        wc = win32gui.WNDCLASS()
        wc.lpszClassName = "PPTMonitorTrayClass"
        wc.style = win32con.CS_HREDRAW | win32con.CS_VREDRAW
        wc.hbrBackground = win32con.COLOR_WINDOW
        wc.hInstance = win32api.GetModuleHandle(None)
        wc.lpfnWndProc = self.wnd_proc
        self.class_atom = win32gui.RegisterClass(wc)
        
        self.hwnd = win32gui.CreateWindow(
            self.class_atom, "PPT Monitor",
            win32con.WS_OVERLAPPED | win32con.WS_SYSMENU,
            0, 0, 200, 200, 0, 0,
            win32api.GetModuleHandle(None), None
        )
        self.create_tray_icon()
    
    def create_tray_icon(self):
        icon_path = Path(__file__).parent.parent / "icon.ico"
        try:
            if icon_path.exists():
                hicon = win32gui.LoadImage(0, str(icon_path), win32con.IMAGE_ICON, 0, 0, 
                                          win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE)
            else:
                hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
        except:
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
        
        nid = (self.hwnd, 0, win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
               self.WM_NOTIFY_CALLBACK, hicon, "PPT文件备份监控")
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
    
    def update_tray_tooltip(self):
        if self.monitor:
            drives = self.monitor.get_connected_drives()
            count = self.monitor.file_manager.get_processed_count()
            fallback = "轮询:开" if self.monitor.enable_fallback else "轮询:关"
            tip = f"PPT备份监控\n已连接: {', '.join([f'{d}:' for d in drives]) or '无'}\n今日备份: {count}个\n{fallback}"
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
            os.startfile(str(config_file))
    
    def open_backup_folder(self):
        if self.monitor:
            os.startfile(str(self.monitor.base_backup_dir))
        else:
            os.startfile(self.config_manager.get_backup_dir())
    
    def open_log_folder(self):
        from core.logger import Logger
        os.startfile(str(Logger().log_dir))
    
    def toggle_fallback_monitor(self):
        if self.monitor:
            self.monitor.set_fallback_enabled(not self.monitor.enable_fallback)
            self.update_tray_tooltip()
    
    def show_about_dialog(self):
        about_text = f"""PPT文件备份监控程序 v2.3

功能:
- 自动检测U盘等移动设备
- 实时监控PowerPoint文件
- 自动备份到本地
- 进程缓存优化性能

备份目录: {self.config_manager.get_backup_dir()}
保留天数: {self.config_manager.get_max_retention_days()}天
"""
        win32api.MessageBox(self.hwnd, about_text, "关于", win32con.MB_OK | win32con.MB_ICONINFORMATION)
    
    def show_context_menu(self, x, y):
        menu = win32gui.CreatePopupMenu()
        status = f"连接设备: {', '.join([f'{d}:' for d in self.monitor.get_connected_drives()]) if self.monitor and self.monitor.get_connected_drives() else '无连接设备'}"
        win32gui.AppendMenu(menu, win32con.MF_GRAYED, 0, status)
        win32gui.AppendMenu(menu, win32con.MF_GRAYED, 0, f"今日备份: {self.monitor.file_manager.get_processed_count() if self.monitor else 0}个")
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
        
        settings_menu = win32gui.CreatePopupMenu()
        win32gui.AppendMenu(settings_menu, win32con.MF_STRING, 1000, "编辑配置文件")
        win32gui.AppendMenu(settings_menu, win32con.MF_STRING, 1002, "打开备份文件夹")
        win32gui.AppendMenu(settings_menu, win32con.MF_STRING, 1003, "打开日志文件夹")
        win32gui.AppendMenu(settings_menu, win32con.MF_SEPARATOR, 0, "")
        win32gui.AppendMenu(settings_menu, win32con.MF_STRING, 1004, "切换轮询监控")
        win32gui.AppendMenu(menu, win32con.MF_POPUP, settings_menu, "设置")
        
        help_menu = win32gui.CreatePopupMenu()
        win32gui.AppendMenu(help_menu, win32con.MF_STRING, 2001, "关于")
        win32gui.AppendMenu(menu, win32con.MF_POPUP, help_menu, "帮助")
        
        win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, "")
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
                self.open_backup_folder()
        elif msg == win32con.WM_COMMAND:
            cmd = wparam & 0xFFFF
            if cmd == 1000: self.open_config_file()
            elif cmd == 1001: self.exit_app()
            elif cmd == 1002: self.open_backup_folder()
            elif cmd == 1003: self.open_log_folder()
            elif cmd == 1004: self.toggle_fallback_monitor()
            elif cmd == 2001: self.show_about_dialog()
        elif msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
            return 0
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
    
    def exit_app(self):
        self.logger.info("正在退出程序...")
        self.running = False
        if self.monitor:
            self.monitor.stop_monitoring()
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, (self.hwnd, 0))
        except:
            pass
        win32gui.PostQuitMessage(0)
    
    def start_monitoring(self):
        self.monitor = PPTMonitor()
        self.monitor_thread = threading.Thread(target=self.monitor.start_monitoring, daemon=True)
        self.monitor_thread.start()
        
        def update_tooltip():
            while self.running:
                time.sleep(5)
                self.update_tray_tooltip()
        threading.Thread(target=update_tooltip, daemon=True).start()
    
    def run(self):
        self.start_monitoring()
        self.update_tray_tooltip()
        while self.running:
            try:
                win32gui.PumpMessages()
            except KeyboardInterrupt:
                self.exit_app()
                break