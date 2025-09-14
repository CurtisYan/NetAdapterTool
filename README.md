# Windows网络适配器管理工具

一个用于管理Windows系统网络适配器速度和双工模式的Python工具。

## 功能特性

- 查看所有网络适配器信息
- 显示当前网络速度和双工模式
- 一键切换网络适配器设置
- 支持命令行操作
- 自动权限管理

## 安装依赖

```bash
pip install -r requirements.txt
```

## 快速开始

### 1. 查看所有适配器
```bash
python main.py --list
```

### 2. 查看特定适配器详情
```bash
python main.py --info "以太网"
```

### 3. 设置网络速度
```bash
python main.py --adapter "以太网" --speed "1000 Mbps Full Duplex"
```

### 4. 设置双工模式
```bash
python main.py --adapter "以太网" --duplex "Full Duplex"
```

### 5. 同时设置速度和双工模式
```bash
python main.py -a "以太网" -s "100 Mbps Full Duplex" -d "Full Duplex" --restart
```

## 注意事项

- 修改网络设置需要管理员权限
- 程序会自动检测权限并提示
- 某些设置可能需要重启适配器才能生效
- 支持适配器名称的模糊匹配

## 文件说明

- `main.py` - 主程序入口
- `network_adapter.py` - 网络适配器信息获取
- `network_settings.py` - 网络设置修改
- `GUIDE.md` - 详细开发指南

## 系统要求

- Windows 10/11
- Python 3.8+
- 管理员权限（修改设置时）
