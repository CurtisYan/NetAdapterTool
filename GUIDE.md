# 网络适配器管理工具 - 开发指南

## 项目结构

```
NetAdapterTool/
├── main.py              # 主程序入口，命令行接口
├── network_adapter.py   # 网络适配器信息获取模块
├── network_settings.py  # 网络设置修改模块
├── requirements.txt     # Python依赖包列表
└── GUIDE.md            # 本开发指南
```

## 核心模块详解

### 1. network_adapter.py - 信息获取模块

**主要作用**: 获取系统中网络适配器的状态信息

#### 核心类: NetworkAdapter
- **初始化**: 创建WMI连接，用于查询Windows系统信息
- **主要功能**: 枚举适配器、获取速度/双工状态、查询支持的选项

#### 重要函数说明:

**get_all_adapters()** 
- 功能: 获取所有物理网络适配器列表
- 返回: 包含适配器信息的字典列表
- 筛选条件: 只返回物理适配器且已启用的

**get_adapter_by_name(name)**
- 功能: 根据名称模糊匹配查找适配器
- 参数: name - 适配器名称关键字
- 返回: 匹配的适配器信息字典

**_get_adapter_speed(device_id)**
- 功能: 通过PowerShell查询适配器当前速度
- 技术: 调用Get-NetAdapterAdvancedProperty命令
- 返回: 速度字符串或"Unknown"

**_get_adapter_duplex(device_id)**
- 功能: 通过PowerShell查询适配器当前双工模式
- 技术: 调用Get-NetAdapterAdvancedProperty命令
- 返回: 双工模式字符串或"Unknown"

**list_available_speeds(adapter_name)**
- 功能: 获取适配器支持的所有速度选项
- 用途: 为用户提供可选的速度设置列表

**list_available_duplex_modes(adapter_name)**
- 功能: 获取适配器支持的所有双工模式选项
- 用途: 为用户提供可选的双工模式列表

### 2. network_settings.py - 设置修改模块

**主要作用**: 修改网络适配器的速度和双工模式设置

#### 核心类: NetworkSettings
- **权限检查**: 自动检测是否有管理员权限
- **PowerShell执行**: 封装PowerShell命令执行逻辑

#### 重要函数说明:

**_check_admin_rights()**
- 功能: 检查当前程序是否以管理员身份运行
- 技术: 使用Windows API IsUserAnAdmin()
- 重要性: 修改网络设置需要管理员权限

**_run_powershell_command(command)**
- 功能: 执行PowerShell命令的通用函数
- 参数: command - 要执行的PowerShell命令字符串
- 返回: (成功状态, 输出信息)的元组
- 用途: 所有网络设置修改都通过此函数执行

**set_adapter_speed(adapter_name, speed)**
- 功能: 设置指定适配器的网络速度
- 技术: 使用Set-NetAdapterAdvancedProperty命令
- 权限: 需要管理员权限

**set_adapter_duplex(adapter_name, duplex_mode)**
- 功能: 设置指定适配器的双工模式
- 技术: 使用Set-NetAdapterAdvancedProperty命令
- 权限: 需要管理员权限

**set_adapter_both(adapter_name, speed, duplex_mode)**
- 功能: 同时设置速度和双工模式
- 逻辑: 先设置速度，再设置双工模式
- 优势: 减少用户操作步骤

**restart_adapter(adapter_name)**
- 功能: 重启网络适配器使设置生效
- 步骤: 先禁用适配器，再启用适配器
- 用途: 某些设置需要重启适配器才能生效

**request_admin_rights()**
- 功能: 请求管理员权限重新启动程序
- 技术: 使用ShellExecuteW API的"runas"动词
- 用途: 当检测到权限不足时自动提权

### 3. main.py - 主程序入口

**主要作用**: 提供命令行接口，整合各个模块功能

#### 重要函数说明:

**print_adapter_info(adapter_info)**
- 功能: 格式化打印适配器信息
- 用途: 统一的信息显示格式

**list_adapters()**
- 功能: 列出系统中所有可用的网络适配器
- 调用: NetworkAdapter.get_all_adapters()
- 输出: 编号列表形式显示所有适配器

**show_adapter_details(adapter_name)**
- 功能: 显示指定适配器的详细信息
- 包含: 当前状态、支持的速度选项、支持的双工模式
- 用途: 帮助用户了解适配器能力

**set_adapter_settings(adapter_name, speed, duplex, restart)**
- 功能: 设置适配器参数的核心函数
- 验证: 检查管理员权限、验证适配器存在性
- 逻辑: 根据参数调用相应的设置函数

**main()**
- 功能: 程序入口点，解析命令行参数
- 技术: 使用argparse库处理命令行参数
- 逻辑: 根据不同参数调用对应功能函数

## 命令行使用方法

### 基本命令格式
```bash
python main.py [选项]
```

### 常用命令示例

**查看所有适配器**
```bash
python main.py --list
python main.py -l
```

**查看特定适配器详情**
```bash
python main.py --info "以太网"
python main.py -i "Ethernet"
```

**设置适配器速度**
```bash
python main.py --adapter "以太网" --speed "1000 Mbps Full Duplex"
python main.py -a "Ethernet" -s "100 Mbps Full Duplex"
```

**设置适配器双工模式**
```bash
python main.py --adapter "以太网" --duplex "Full Duplex"
python main.py -a "Ethernet" -d "Half Duplex"
```

**同时设置速度和双工模式**
```bash
python main.py -a "以太网" -s "1000 Mbps Full Duplex" -d "Full Duplex"
```

**设置后重启适配器**
```bash
python main.py -a "以太网" -s "100 Mbps Full Duplex" --restart
```

## 技术要点

### Windows API调用
- **WMI**: 用于查询系统硬件信息
- **PowerShell**: 用于执行网络管理命令
- **ctypes**: 用于调用Windows系统API

### 权限管理
- 程序自动检测管理员权限
- 可以自动请求提权重启
- 所有修改操作都需要管理员权限

### 错误处理
- 每个关键操作都有异常捕获
- 提供清晰的错误信息反馈
- 权限不足时给出明确提示

### PowerShell命令
- `Get-NetAdapterAdvancedProperty`: 查询适配器高级属性
- `Set-NetAdapterAdvancedProperty`: 设置适配器高级属性
- `Disable-NetAdapter/Enable-NetAdapter`: 禁用/启用适配器

## 开发注意事项

1. **权限要求**: 修改网络设置必须以管理员身份运行
2. **适配器名称**: 支持部分匹配，便于用户使用
3. **设置值格式**: 必须使用适配器支持的确切字符串值
4. **异常处理**: 所有外部调用都有异常保护
5. **跨版本兼容**: 代码适用于Windows 10/11系统

## 下一步开发计划

1. **图形界面**: 使用PyQt5开发GUI版本
2. **配置保存**: 支持保存和恢复网络配置
3. **批量操作**: 支持同时操作多个适配器
4. **自动检测**: 自动检测最优网络设置
5. **打包分发**: 使用PyInstaller打包为可执行文件
