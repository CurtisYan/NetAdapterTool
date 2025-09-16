"""
网络适配器信息获取和管理模块
用于获取Windows系统中网络适配器的状态信息
"""

import wmi
import subprocess
import json
import pythoncom
from typing import List, Dict, Optional


class NetworkAdapter:
    def __init__(self):
        self.wmi_conn = None
        self._init_wmi_connection()
    
    def _init_wmi_connection(self):
        """初始化WMI连接，带重试机制"""
        max_retries = 3
        for attempt in range(max_retries):
            try:
                # 在多线程环境中初始化COM组件
                try:
                    pythoncom.CoInitialize()
                except:
                    pass  # 如果已经初始化过，忽略错误
                
                self.wmi_conn = wmi.WMI()
                return
            except Exception as e:
                if attempt == max_retries - 1:
                    raise Exception(f"WMI连接失败，已重试{max_retries}次: {str(e)}")
                import time
                time.sleep(1)  # 等待1秒后重试
    
    def reconnect_wmi(self):
        """重新连接WMI"""
        self.wmi_conn = None
        # 在重连前也初始化COM组件
        try:
            pythoncom.CoInitialize()
        except:
            pass  # 如果已经初始化过，忽略错误
        self._init_wmi_connection()
    
    def get_all_adapters(self) -> List[Dict]:
        """获取所有网络适配器信息"""
        adapters = []
        
        try:
            # 检查WMI连接
            if self.wmi_conn is None:
                self._init_wmi_connection()
            
            # 获取物理网络适配器，包括已连接和未连接的
            for adapter in self.wmi_conn.Win32_NetworkAdapter():
                # 更宽松的筛选条件：物理适配器且不是软件适配器
                if (adapter.PhysicalAdapter and 
                    adapter.Name and 
                    adapter.MACAddress and
                    'Virtual' not in adapter.Name and
                    'Loopback' not in adapter.Name and
                    'Wireless' not in adapter.Name):
                    
                    adapter_info = {
                        'name': adapter.Name,
                        'device_id': adapter.DeviceID,
                        'mac_address': adapter.MACAddress,
                        'alias': adapter.NetConnectionID,
                        'ip_address': self._get_adapter_ip(adapter.NetConnectionID),
                        'status': adapter.NetConnectionStatus,
                        'speed': self._get_adapter_speed(adapter.NetConnectionID),
                        'duplex': self._get_adapter_duplex(adapter.NetConnectionID)
                    }
                    adapters.append(adapter_info)
        
        except Exception as e:
            # WMI连接失败时，尝试重新连接
            try:
                self.reconnect_wmi()
                # 重试一次，使用更宽松的条件
                for adapter in self.wmi_conn.Win32_NetworkAdapter():
                    if (adapter.PhysicalAdapter and 
                        adapter.Name and 
                        adapter.MACAddress and
                        'Virtual' not in adapter.Name and
                        'Loopback' not in adapter.Name and
                        'Wireless' not in adapter.Name):
                        
                        adapter_info = {
                            'name': adapter.Name,
                            'device_id': adapter.DeviceID,
                            'mac_address': adapter.MACAddress,
                            'alias': adapter.NetConnectionID,
                            'ip_address': self._get_adapter_ip(adapter.NetConnectionID),
                            'status': adapter.NetConnectionStatus,
                            'speed': self._get_adapter_speed(adapter.NetConnectionID),
                            'duplex': self._get_adapter_duplex(adapter.NetConnectionID)
                        }
                        adapters.append(adapter_info)
            except Exception as retry_error:
                raise Exception(f"获取网络适配器失败: {str(retry_error)}")
        
        return adapters
    
    def _get_adapter_ip(self, connection_id: str) -> Optional[str]:
        """获取适配器IP地址"""
        if not connection_id:
            return 'Unknown'
        try:
            cmd = f'Get-NetIPAddress -InterfaceAlias "{connection_id}" -AddressFamily IPv4 | Select-Object -First 1 IPAddress | ConvertTo-Json'
            result = subprocess.run(['powershell', '-Command', cmd], 
                                  capture_output=True, text=True, shell=True)
            
            if result.returncode == 0 and result.stdout.strip():
                try:
                    data = json.loads(result.stdout)
                    return data.get('IPAddress', 'Unknown')
                except (json.JSONDecodeError, ValueError):
                    pass
        except:
            pass
        return 'Unknown'
    
    def _get_adapter_speed(self, connection_id: str) -> Optional[str]:
        """获取适配器当前速度设置"""
        if not connection_id:
            return 'Unknown'
        try:
            # 尝试多种方法获取速度信息
            methods = [
                f'Get-NetAdapter -Name "{connection_id}" | Select-Object LinkSpeed | ConvertTo-Json',
                f'Get-NetAdapterAdvancedProperty -Name "{connection_id}" -DisplayName "*Speed*" | Select-Object -First 1 DisplayValue | ConvertTo-Json',
                f'Get-NetAdapterAdvancedProperty -Name "{connection_id}" -DisplayName "*Link Speed*" | Select-Object -First 1 DisplayValue | ConvertTo-Json'
            ]
            
            for cmd in methods:
                result = subprocess.run(['powershell', '-Command', cmd], 
                                      capture_output=True, text=True, shell=True)
                
                if result.returncode == 0 and result.stdout.strip():
                    try:
                        data = json.loads(result.stdout)
                        speed_value = data.get('LinkSpeed') or data.get('DisplayValue')
                        if speed_value and speed_value != 'Unknown':
                            # 格式化速度显示
                            if isinstance(speed_value, int):
                                return f"{speed_value // 1000000} Mbps"
                            return str(speed_value)
                    except (json.JSONDecodeError, ValueError):
                        continue
        except:
            pass
        return 'Unknown'
    
    def _get_adapter_duplex(self, connection_id: str) -> Optional[str]:
        """获取适配器当前双工模式"""
        if not connection_id:
            return 'Unknown'
        try:
            # 尝试多种方法获取双工模式信息
            methods = [
                f'Get-NetAdapter -Name "{connection_id}" | Select-Object FullDuplex | ConvertTo-Json',
                f'Get-NetAdapterAdvancedProperty -Name "{connection_id}" -DisplayName "*Duplex*" | Select-Object -First 1 DisplayValue | ConvertTo-Json',
                f'Get-NetAdapterAdvancedProperty -Name "{connection_id}" -DisplayName "*Flow Control*" | Select-Object -First 1 DisplayValue | ConvertTo-Json'
            ]
            
            for cmd in methods:
                result = subprocess.run(['powershell', '-Command', cmd], 
                                      capture_output=True, text=True, shell=True)
                
                if result.returncode == 0 and result.stdout.strip():
                    try:
                        data = json.loads(result.stdout)
                        duplex_value = data.get('FullDuplex') or data.get('DisplayValue')
                        if duplex_value is not None and duplex_value != 'Unknown':
                            # 格式化双工模式显示为中文
                            if isinstance(duplex_value, bool):
                                return "全双工" if duplex_value else "半双工"
                            duplex_str = str(duplex_value).lower()
                            if 'full' in duplex_str:
                                return "全双工"
                            elif 'half' in duplex_str:
                                return "半双工"
                            return str(duplex_value)
                    except (json.JSONDecodeError, ValueError):
                        continue
        except:
            pass
        return 'Unknown'
    
    def get_adapter_by_name(self, name: str) -> Optional[Dict]:
        """根据名称获取特定适配器信息"""
        adapters = self.get_all_adapters()
        for adapter in adapters:
            if name.lower() in adapter['name'].lower():
                return adapter
        return None
    
    def get_speed_duplex_options(self, adapter_name: str = None) -> List[str]:
        """动态获取适配器支持的速度双工选项"""
        if adapter_name:
            try:
                # 转义适配器名称中的特殊字符
                safe_name = adapter_name.replace('"', '`"').replace("'", "`'")
                cmd = f'Get-NetAdapterAdvancedProperty -Name "{safe_name}" -RegistryKeyword "*SpeedDuplex" | Select-Object -ExpandProperty ValidDisplayValues'
                result = subprocess.run(['powershell', '-Command', cmd], 
                                      capture_output=True, text=True, shell=True, timeout=10)
                
                if result.returncode == 0 and result.stdout.strip():
                    options = [line.strip() for line in result.stdout.strip().split('\n') if line.strip()]
                    return options
            except (subprocess.TimeoutExpired, subprocess.SubprocessError):
                pass
        
        # 如果无法获取实际选项，返回空列表
        return []
