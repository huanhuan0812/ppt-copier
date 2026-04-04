"""程序入口"""
import sys
from core.logger import Logger
from core.single_instance import SingleInstance
from core.config import ConfigManager
from ui.tray import SystemTrayApp


def main():
    single_instance = SingleInstance()
    if not single_instance.is_first():
        logger = Logger()
        logger.info("程序已在运行，尝试激活现有窗口...")
        single_instance.bring_to_front()
        return
    
    logger = Logger()
    logger.info("=" * 60)
    logger.info("PowerPoint PPT文件备份监控程序启动 v2.3")
    
    config_manager = ConfigManager()
    backup_dir = config_manager.get_backup_dir()
    
    try:
        from pathlib import Path
        Path(backup_dir).mkdir(exist_ok=True)
        logger.info(f"备份目录: {backup_dir}")
        logger.info(f"保留天数: {config_manager.get_max_retention_days()}天")
        logger.info(f"最小文件大小: {config_manager.get_min_file_size_kb()}KB")
        logger.info(f"后备监控: {'启用' if config_manager.get_enable_fallback_monitor() else '禁用'}")
    except Exception as e:
        logger.error(f"初始化失败: {e}")
        return
    
    try:
        import psutil
    except ImportError:
        logger.critical("缺失依赖 psutil，请运行: pip install psutil")
        sys.exit(1)
    
    app = SystemTrayApp()
    try:
        app.run()
    except KeyboardInterrupt:
        logger.info("程序被用户中断")
    except Exception as e:
        logger.exception(f"程序运行出错: {e}")
    
    logger.info("程序已退出")
    sys.exit(0)


if __name__ == "__main__":
    main()