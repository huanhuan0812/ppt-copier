# ppt-copier
a script that copy ppt file from  portable storage device to other place

### 实现思路
- Windows 移动设备插拔事件监听
- powerpoint com组件文件打开事件
- explorer预览窗格过滤（ppt预览时同样存在打开事件，且此时无权限获取文件路径导致报错）
- 托盘图标状态监测
- 缓存
