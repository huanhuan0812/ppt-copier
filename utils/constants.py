"""常量定义模块"""
import struct
from ctypes import Structure, c_ulong, byref, POINTER, WINFUNCTYPE, sizeof
from ctypes.wintypes import HWND, UINT, WPARAM, LPARAM, HICON, DWORD, HANDLE, BOOL

# 文件扩展名
PPT_EXTENSIONS = {'.ppt', '.pptx', '.pps', '.ppsx'}

# 应用名称和互斥锁
APP_NAME = "PPTMonitor"
MUTEX_NAME = "Global\\{B8E2C5A1-9F4D-4E8E-9A2B-3C5D7E9F1A2B}"

# 状态保存间隔（秒）
STATE_SAVE_INTERVAL = 30

# Windows消息常量
WM_DEVICECHANGE = 0x0219
DBT_DEVICEARRIVAL = 0x8000
DBT_DEVICEREMOVECOMPLETE = 0x8004
DBT_DEVTYP_VOLUME = 0x00000002

# 设备通知标志
DEVICE_NOTIFY_WINDOW_HANDLE = 0x00000000
DEVICE_NOTIFY_SERVICE_HANDLE = 0x00000001
DEVICE_NOTIFY_ALL_INTERFACE_CLASSES = 0x00000004

# COM 退出错误码
QUITTING_HRESULTS = [
    -2147417848,  # RPC_E_SERVER_DIED
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
    -2147418113,  # RPC_E_SERVER_DIED_DNE
]


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