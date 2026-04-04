"""PowerPoint COM 事件处理器"""
import threading
import time
from core.logger import Logger


class PowerPointEvents:
    """PowerPoint COM 事件处理器"""
    def __init__(self, monitor_instance=None):
        self.monitor = monitor_instance
        self.logger = Logger()
    
    def set_monitor(self, monitor): self.monitor = monitor
    
    def _should_process_file(self, file_path):
        return file_path and self.monitor and self.monitor.is_valid_ppt_file_for_backup(file_path)
    
    def _is_on_removable_drive(self, file_path):
        return file_path and self.monitor and self.monitor.is_removable_drive(file_path)
    
    def OnPresentationOpen(self, Pres):
        """演示文稿打开事件"""
        try:
            file_path = Pres.FullName if hasattr(Pres, 'FullName') else None
            if file_path and self._is_on_removable_drive(file_path):
                self.logger.info(f"COM事件 - PowerPoint打开文件: {file_path}")
                
                def delayed_check():
                    time.sleep(0.5)
                    if self._should_process_file(file_path):
                        self.monitor and self.monitor.process_ppt_file(file_path, source="COM事件(打开)")
                threading.Thread(target=delayed_check, daemon=True).start()
        except Exception as e:
            self.logger.error(f"处理OnPresentationOpen事件时出错: {e}")
    
    def OnPresentationClose(self, Pres):
        """演示文稿关闭事件"""
        if self.monitor and hasattr(self.monitor, 'process_cache'):
            def async_cleanup():
                time.sleep(0.1)
                self.monitor.process_cache.invalidate_process()
            threading.Thread(target=async_cleanup, daemon=True).start()
    
    def OnPresentationSave(self, Pres):
        """演示文稿保存事件"""
        try:
            file_path = Pres.FullName if hasattr(Pres, 'FullName') else None
            if file_path and self._is_on_removable_drive(file_path):
                self.logger.info(f"COM事件 - PowerPoint保存文件: {file_path}")
                if self._should_process_file(file_path):
                    self.monitor and self.monitor.process_ppt_file(file_path, source="COM事件(保存)")
        except Exception as e:
            self.logger.error(f"处理OnPresentationSave事件时出错: {e}")
    
    def OnQuit(self):
        """PowerPoint退出事件"""
        self.logger.info("COM事件 - PowerPoint正在退出")
        if self.monitor:
            def async_quit():
                time.sleep(0.1)
                self.monitor.on_powerpoint_quit()
                self.monitor.process_cache.clear_all()
            threading.Thread(target=async_quit, daemon=True).start()