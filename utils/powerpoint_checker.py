"""
PowerPoint COM可用性检查工具
用于检测系统是否安装了Microsoft PowerPoint以及COM接口是否可用
"""

import sys
import os
from typing import Tuple, Dict, Optional


class PowerPointChecker:
    """PowerPoint安装和COM可用性检查器"""
    
    def __init__(self):
        """初始化检查器"""
        self.results = {
            'powerpoint_installed': False,
            'com_available': False,
            'powerpoint_version': None,
            'com_error': None,
            'installation_path': None
        }
    
    def check_powerpoint_installation(self) -> Tuple[bool, Optional[str], Optional[str]]:
        """
        检查PowerPoint是否安装
        
        Returns:
            tuple: (是否安装, 安装路径, 版本号)
        """
        if sys.platform != 'win32':
            self.results['com_error'] = "PowerPoint检查仅支持Windows系统"
            return False, None, None
        
        # 通过Windows注册表检查
        try:
            import winreg
            reg_paths = [
                r"SOFTWARE\Microsoft\Office\16.0\PowerPoint\Application",
                r"SOFTWARE\Microsoft\Office\15.0\PowerPoint\Application",
                r"SOFTWARE\Microsoft\Office\14.0\PowerPoint\Application",
                r"SOFTWARE\Microsoft\Office\12.0\PowerPoint\Application",
                r"SOFTWARE\Microsoft\Office\11.0\PowerPoint\Application",
                r"SOFTWARE\WOW6432Node\Microsoft\Office\16.0\PowerPoint\Application",
                r"SOFTWARE\WOW6432Node\Microsoft\Office\15.0\PowerPoint\Application",
                r"SOFTWARE\WOW6432Node\Microsoft\Office\14.0\PowerPoint\Application",
                r"SOFTWARE\WOW6432Node\Microsoft\Office\12.0\PowerPoint\Application",
            ]
            
            for reg_path in reg_paths:
                try:
                    with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
                        # 获取安装路径
                        try:
                            install_path, _ = winreg.QueryValueEx(key, "Path")
                        except:
                            install_path = "未知"
                        
                        # 获取版本信息
                        version = reg_path.split('\\')[3] if 'Office' in reg_path else "未知"
                        
                        self.results['powerpoint_installed'] = True
                        self.results['installation_path'] = install_path
                        self.results['powerpoint_version'] = version
                        return True, install_path, version
                except:
                    continue
                    
        except ImportError:
            pass
        except Exception as e:
            self.results['com_error'] = f"注册表访问错误: {str(e)}"
        
        # 检查常见安装路径
        common_paths = [
            r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\POWERPNT.EXE",
            r"C:\Program Files\Microsoft Office\Office16\POWERPNT.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE",
            r"C:\Program Files\Microsoft Office\Office15\POWERPNT.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office15\POWERPNT.EXE",
            r"C:\Program Files\Microsoft Office\Office14\POWERPNT.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office14\POWERPNT.EXE",
        ]
        
        for path in common_paths:
            if os.path.exists(path):
                self.results['powerpoint_installed'] = True
                self.results['installation_path'] = path
                return True, path, None
        
        return False, None, None
    
    def check_com_availability(self) -> Tuple[bool, Optional[str]]:
        """
        检查PowerPoint COM对象是否可用
        
        Returns:
            tuple: (COM是否可用, 错误信息)
        """
        if sys.platform != 'win32':
            return False, "COM检查仅支持Windows系统"
        
        try:
            import win32com.client
            from win32com.client import Dispatch
        except ImportError:
            error_msg = "pywin32未安装。请运行: pip install pywin32"
            self.results['com_error'] = error_msg
            return False, error_msg
        
        try:
            ppt = None
            try:
                ppt = Dispatch("PowerPoint.Application")
                if ppt:
                    # 获取版本信息
                    if hasattr(ppt, 'Version'):
                        version = ppt.Version
                        self.results['powerpoint_version'] = version
                    
                    self.results['com_available'] = True
                    ppt.Quit()
                    ppt = None
                    return True, None
            except Exception as e:
                error_msg = f"创建PowerPoint COM对象失败: {str(e)}"
                self.results['com_error'] = error_msg
                return False, error_msg
            finally:
                if ppt:
                    try:
                        ppt.Quit()
                    except:
                        pass
                        
        except Exception as e:
            error_msg = f"COM调用失败: {str(e)}"
            self.results['com_error'] = error_msg
            return False, error_msg
        
        return False, "未知错误"
    
    def check_all(self) -> Dict:
        """
        执行完整检查
        
        Returns:
            dict: 包含所有检查结果的字典
        """
        # 检查PowerPoint安装
        self.check_powerpoint_installation()
        
        # 检查COM可用性
        self.check_com_availability()
        
        return self.results
    
    def is_ready_for_automation(self) -> bool:
        """
        检查是否可以进行PowerPoint自动化操作
        
        Returns:
            bool: 是否可以安全使用PowerPoint COM
        """
        return self.results['powerpoint_installed'] and self.results['com_available']
    
    def get_results(self) -> Dict:
        """
        获取检查结果
        
        Returns:
            dict: 检查结果字典
        """
        return self.results


def quick_check() -> bool:
    """
    快速检查函数
    
    Returns:
        bool: PowerPoint COM是否可用
    """
    checker = PowerPointChecker()
    results = checker.check_all()
    return checker.is_ready_for_automation()