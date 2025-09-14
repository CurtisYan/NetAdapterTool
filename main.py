"""
网络适配器管理工具主程序
提供命令行接口来管理网络适配器的速度和双工模式
"""

import argparse
import sys
from network_adapter import NetworkAdapter
from network_settings import NetworkSettings


def print_adapter_info(adapter_info):
    """打印适配器信息"""
    print(f"适配器名称: {adapter_info['name']}")
    print(f"MAC地址: {adapter_info['mac_address']}")
    print(f"IP地址: {adapter_info['ip_address']}")
    print(f"当前速度: {adapter_info['speed']}")
    print(f"当前双工模式: {adapter_info['duplex']}")
    print(f"连接状态: {'已连接' if adapter_info['status'] == 2 else '未连接'}")
    print("-" * 50)


def list_adapters():
    """列出所有网络适配器"""
    adapter = NetworkAdapter()
    adapters = adapter.get_all_adapters()
    
    if not adapters:
        print("未找到可用的网络适配器")
        return
    
    print("系统中的网络适配器:")
    print("=" * 50)
    for i, adapter_info in enumerate(adapters, 1):
        print(f"[{i}] ", end="")
        print_adapter_info(adapter_info)


def show_adapter_details(adapter_name):
    """显示特定适配器的详细信息"""
    adapter = NetworkAdapter()
    adapter_info = adapter.get_adapter_by_name(adapter_name)
    
    if not adapter_info:
        print(f"未找到名称包含 '{adapter_name}' 的适配器")
        return
    
    print("适配器详细信息:")
    print("=" * 50)
    print_adapter_info(adapter_info)
    
    # 显示支持的速度选项
    speeds = adapter.list_available_speeds(adapter_info['name'])
    if speeds:
        print("支持的速度选项:")
        for speed in speeds:
            print(f"  - {speed}")
    
    # 显示支持的双工模式选项
    duplex_modes = adapter.list_available_duplex_modes(adapter_info['name'])
    if duplex_modes:
        print("支持的双工模式:")
        for mode in duplex_modes:
            print(f"  - {mode}")


def set_adapter_settings(adapter_name, speed=None, duplex=None, restart=False):
    """设置适配器参数"""
    settings = NetworkSettings()
    
    if not settings.is_admin:
        print("错误: 需要管理员权限才能修改网络设置")
        print("请以管理员身份运行此程序")
        return False
    
    # 验证适配器是否存在
    adapter = NetworkAdapter()
    adapter_info = adapter.get_adapter_by_name(adapter_name)
    if not adapter_info:
        print(f"错误: 未找到名称包含 '{adapter_name}' 的适配器")
        return False
    
    actual_name = adapter_info['name']
    
    # 设置速度和双工模式
    if speed and duplex:
        success, message = settings.set_adapter_both(actual_name, speed, duplex)
    elif speed:
        success, message = settings.set_adapter_speed(actual_name, speed)
    elif duplex:
        success, message = settings.set_adapter_duplex(actual_name, duplex)
    else:
        print("错误: 必须指定速度或双工模式")
        return False
    
    print(message)
    
    # 如果设置成功且需要重启适配器
    if success and restart:
        print("正在重启适配器...")
        restart_success, restart_message = settings.restart_adapter(actual_name)
        print(restart_message)
        return restart_success
    
    return success


def main():
    parser = argparse.ArgumentParser(description='Windows网络适配器管理工具')
    parser.add_argument('--list', '-l', action='store_true', help='列出所有网络适配器')
    parser.add_argument('--info', '-i', type=str, help='显示指定适配器的详细信息')
    parser.add_argument('--adapter', '-a', type=str, help='要操作的适配器名称(部分匹配)')
    parser.add_argument('--speed', '-s', type=str, help='设置网络速度 (如: 1000 Mbps Full Duplex)')
    parser.add_argument('--duplex', '-d', type=str, help='设置双工模式 (如: Full Duplex)')
    parser.add_argument('--restart', '-r', action='store_true', help='设置后重启适配器')
    
    args = parser.parse_args()
    
    # 如果没有参数，显示帮助
    if len(sys.argv) == 1:
        parser.print_help()
        return
    
    # 列出所有适配器
    if args.list:
        list_adapters()
        return
    
    # 显示适配器详细信息
    if args.info:
        show_adapter_details(args.info)
        return
    
    # 设置适配器参数
    if args.adapter and (args.speed or args.duplex):
        success = set_adapter_settings(args.adapter, args.speed, args.duplex, args.restart)
        if success:
            print("设置完成!")
        else:
            print("设置失败!")
        return
    
    # 参数不完整
    if args.adapter and not (args.speed or args.duplex):
        print("错误: 指定适配器时必须同时指定速度或双工模式")
        parser.print_help()
    elif (args.speed or args.duplex) and not args.adapter:
        print("错误: 设置参数时必须指定适配器")
        parser.print_help()


if __name__ == "__main__":
    main()
