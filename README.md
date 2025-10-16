# [网络适配器管理工具](https://github.com/CurtisYan/NetAdapterTool)

Windows系统网络适配器速度和双工模式管理工具，支持图形化界面操作。为NA（广软网协）而做

## 下载
- GitHub 发布页：[Releases](https://github.com/CurtisYan/NetAdapterTool/releases)
- 123云盘镜像：[123云盘免登录下载不限速](https://www.123912.com/s/S20tjv-yXVsd)

## 示例截图
 
![示例1](img/example1.png) ![示例2](img/example2.png)
 
## 功能特性

本工具提供现代化的图形界面，支持查看和管理Windows系统中的网络适配器设置，包括实时显示网络速度、双工模式和IP地址信息，具备自动权限检测和多线程处理能力，确保操作流畅不卡顿。

### 🔧 核心功能
- **网络适配器管理** - 查看和修改网络适配器速度和双工模式
- **实时状态显示** - 显示IP地址、连接状态和网络速度
- **智能过滤** - 支持仅显示有线网卡，过滤无线适配器
- **多线程处理** - 后台操作，界面响应流畅

### 🌐 系统兼容性
- **Windows版本支持** - Windows 7/8/10/11 (32位/64位)
- **PowerShell兼容** - 支持PowerShell 5.x 和 PowerShell 7.x
- **多路径检测** - 自动检测系统中可用的PowerShell版本
- **WMI兼容性** - 智能WMI连接管理，支持多线程环境
- **权限管理** - 多种管理员权限检测方法，自动提权

### 🛡️ 健壮性设计
- **系统诊断** - 内置兼容性检查和诊断工具
- **错误恢复** - 智能错误处理和降级方案
- **资源管理** - 支持多种部署方式（源码/打包exe）
- **日志系统** - 详细的操作日志和错误追踪

## 快速开始

### 源码运行

#### 安装依赖
```bash
pip install -r requirements.txt
```

#### 运行程序
```bash
python gui.py
```

**注意**：需要以管理员身份运行命令行或IDE

## 打包（Nuitka）

在项目根目录执行以下命令完成打包（示例）：

```bat
nuitka ^
  --standalone --onefile ^
  --enable-plugin=pyqt5 ^
  --include-qt-plugins=platforms,imageformats,styles ^
  --python-flag=-OO ^
  --nofollow-import-to=PyQt5.QtQml,PyQt5.QtQuick,PyQt5.QtWebEngineWidgets,PyQt5.QtWebKit,PyQt5.QtMultimedia,PyQt5.QtNetwork,PyQt5.QtSql,PyQt5.QtPrintSupport,tkinter,matplotlib,numpy,pandas ^
  --enable-plugin=upx ^
  --upx-binary="D:\upx\upx.exe" ^
  --onefile-no-compression ^
  --windows-uac-admin ^
  --windows-console-mode=disable ^
  --windows-icon-from-ico=img/NA.ico ^
  --include-data-dir=img=img ^
  --product-name="NetAdapterTool" ^
  --file-version=1.0.0 ^
  --product-version=1.0.0 ^
  --output-filename="网络适配器修改器.exe" ^
  gui.py
```

## 注意事项

### 开发环境
❗ 以管理员身份运行Python脚本或IDE

### 生产环境  
✅ 使用打包后的exe文件，会自动请求管理员权限

## 项目结构

```text
NetAdapterTool/
├─ gui.py
├─ network_adapter.py
├─ network_settings.py
├─ system_compatibility.py
├─ app.manifest
├─ requirements.txt
├─ img/
│  ├─ NA.ico
│  ├─ example1.png
│  ├─ example2.png
│  └─ NA (蓝透明).jpg
  └─ README.md
```
 
## 故障排除

### 常见问题
1. **WMI连接失败** → 确保以管理员身份运行
2. **PowerShell超时** → 检查系统负载和PowerShell版本
3. **找不到适配器** → 检查网络适配器驱动是否正常

### 技术支持
- 查看程序日志（点击"显示日志"按钮）
- 使用系统诊断功能（帮助菜单）
- 检查Windows事件查看器
- 确认WMI服务正常运行

## 许可证

MIT License
