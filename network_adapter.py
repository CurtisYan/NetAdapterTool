"""
优化版网络适配器信息获取和管理模块
解决卡死问题，提升性能和兼容性
"""

import wmi
import subprocess
import json
import pythoncom
import threading
import time
import logging
from typing import List, Dict, Optional
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FutureTimeoutError

# 默认速度双工选项（当无法从系统获取时使用）
DEFAULT_SPEED_DUPLEX_OPTIONS = [
    "自动侦测", 
    "10 Mbps 半双工", 
    "10 Mbps 全双工", 
    "100 Mbps 半双工", 
    "100 Mbps 全双工", 
    "1.0 Gbps 全双工"
]

# 全局配置
CONFIG = {
    'WMI_TIMEOUT': 15,          # WMI连接超时时间
    'POWERSHELL_TIMEOUT': 8,    # PowerShell命令超时时间
    'MAX_RETRIES': 2,           # 最大重试次数
    'THREAD_POOL_SIZE': 4,      # 线程池大小
}


class NetworkAdapter:
    def __init__(self, lazy_init=True):
        """
        优化版网络适配器类
        
        Args:
            lazy_init: 是否延迟初始化WMI连接（推荐True）
        """
        self.wmi_conn = None
        self._wmi_lock = threading.Lock()
        self._initialized = False
        
        if not lazy_init:
            self._init_wmi_connection()
    
    def _init_wmi_connection(self):
        """初始化WMI连接，带超时和重试机制"""
        if self._initialized:
            return True
            
        with self._wmi_lock:
            if self._initialized:  # 双重检查
                return True
                
            max_retries = CONFIG['MAX_RETRIES']
            for attempt in range(max_retries):
                try:
                    # 在多线程环境中初始化COM组件
                    try:
                        pythoncom.CoInitialize()
                    except:
                        pass  # 如果已经初始化过，忽略错误
                    
                    # 使用线程池执行WMI初始化，带超时
                    with ThreadPoolExecutor(max_workers=1) as executor:
                        future = executor.submit(self._create_wmi_connection)
                        self.wmi_conn = future.result(timeout=CONFIG['WMI_TIMEOUT'])
                    
                    self._initialized = True
                    return True
                    
                except FutureTimeoutError:
                    if attempt == max_retries - 1:
                        raise Exception(f"WMI连接超时，已重试{max_retries}次")
                    time.sleep(1)
                except Exception as e:
                    if attempt == max_retries - 1:
                        raise Exception(f"WMI连接失败，已重试{max_retries}次: {str(e)}")
                    time.sleep(1)
            
            return False
    
    def _create_wmi_connection(self):
        """创建WMI连接的实际方法"""
        # 在新线程中必须初始化COM组件
        try:
            pythoncom.CoInitialize()
        except:
            pass  # 如果已经初始化过，忽略错误
        
        return wmi.WMI()
    
    def reconnect_wmi(self):
        """重新连接WMI"""
        with self._wmi_lock:
            self.wmi_conn = None
            self._initialized = False
            
        # 在重连前也初始化COM组件
        try:
            pythoncom.CoInitialize()
        except:
            pass
            
        return self._init_wmi_connection()
    
    def cleanup(self):
        """清理资源"""
        with self._wmi_lock:
            if self.wmi_conn:
                try:
                    self.wmi_conn = None
                except:
                    pass
            self._initialized = False
            
        # 清理COM组件
        try:
            pythoncom.CoUninitialize()
        except:
            pass
    
    def _run_powershell_safe(self, command: str, timeout: int = None) -> tuple:
        """安全执行PowerShell命令，带超时和错误处理，支持多种Windows版本"""
        if timeout is None:
            timeout = CONFIG['POWERSHELL_TIMEOUT']
        
        # 尝试多种PowerShell路径，提高兼容性
        powershell_paths = [
            'powershell',  # 系统PATH中的PowerShell
            r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe',  # Windows PowerShell 5.x
            r'C:\Program Files\PowerShell\7\pwsh.exe',  # PowerShell 7.x
            r'C:\Program Files (x86)\PowerShell\7\pwsh.exe',  # PowerShell 7.x (x86)
        ]
        
        for ps_path in powershell_paths:
            try:
                # 根据路径类型选择调用方式
                if ps_path == 'powershell':
                    # 使用shell=True调用系统PATH中的PowerShell
                    result = subprocess.run(
                        ['powershell', '-ExecutionPolicy', 'Bypass', '-Command', command], 
                        capture_output=True, 
                        text=True, 
                        shell=True, 
                        timeout=timeout,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                else:
                    # 使用完整路径，不需要shell=True
                    import os
                    if not os.path.exists(ps_path):
                        continue
                    result = subprocess.run(
                        [ps_path, '-ExecutionPolicy', 'Bypass', '-Command', command], 
                        capture_output=True, 
                        text=True, 
                        timeout=timeout,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                
                if result.returncode == 0:
                    return True, result.stdout.strip()
                else:
                    # 如果是第一个路径失败，尝试下一个
                    if ps_path == powershell_paths[0]:
                        continue
                    return False, result.stderr.strip() or "PowerShell命令执行失败"
                    
            except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
                # 尝试下一个PowerShell路径
                continue
            except Exception as e:
                # 如果是最后一个路径，返回错误
                if ps_path == powershell_paths[-1]:
                    return False, f"PowerShell命令执行异常: {str(e)}"
                continue
        
        return False, "未找到可用的PowerShell"
    
    def get_all_adapters(self) -> List[Dict]:
        """获取所有网络适配器信息（优化版）"""
        adapters: List[Dict] = []
        
        # 1) 首选 PowerShell：Get-NetAdapter（更稳定、无COM依赖）
        try:
            ps_cmd = 'Get-NetAdapter -Physical | Select-Object Name,InterfaceDescription,MacAddress,Status,LinkSpeed | ConvertTo-Json -Depth 2'
            success, output = self._run_powershell_safe(ps_cmd, timeout=6)
            if success and output:
                try:
                    data = json.loads(output)
                    items = data if isinstance(data, list) else [data]
                    for item in items:
                        name = (item.get('InterfaceDescription') or item.get('Name') or '').strip()
                        alias = (item.get('Name') or '').strip()
                        mac = (item.get('MacAddress') or '').strip()
                        status = item.get('Status')
                        speed = (item.get('LinkSpeed') or 'Unknown').strip()
                        
                        # 跳过无效项（仍排除虚拟/回环，保留无线）
                        if not alias or not name or not mac:
                            continue
                        if 'Virtual' in name or 'Loopback' in name:
                            continue
                        
                        # 并行获取 IP/双工（基于 alias 作为 InterfaceAlias）
                        adapters.append({
                            'name': name,
                            'device_id': alias,  # 无WMI DeviceID，这里用alias占位
                            'mac_address': mac,
                            'alias': alias,
                            'ip_address': self._get_adapter_ip_fast(alias),
                            'status': status,
                            'speed': speed if speed else self._get_adapter_speed_fast(alias),
                            'duplex': self._get_adapter_duplex_fast(alias)
                        })
                    
                    if adapters:
                        logging.info(f"适配器枚举使用: PowerShell(Get-NetAdapter)，找到 {len(adapters)} 个")
                        return adapters
                except Exception:
                    # JSON解析失败则尝试WMI兜底
                    pass
        except Exception:
            # PowerShell 失败，进入WMI兜底
            pass
        
        # 2) 兜底：WMI（为避免跨线程复用同一连接，这里使用本地连接）
        try:
            local_wmi = self._create_wmi_connection()
        except Exception as e:
            raise Exception(f"WMI连接失败，已重试{CONFIG['MAX_RETRIES']}次: {str(e)}")
        
        try:
            wmi_adapters = []
            for adapter in local_wmi.Win32_NetworkAdapter():
                if (adapter.PhysicalAdapter and 
                    adapter.Name and 
                    adapter.MACAddress and
                    'Virtual' not in adapter.Name and
                    'Loopback' not in adapter.Name):
                    wmi_adapters.append(adapter)
            
            with ThreadPoolExecutor(max_workers=CONFIG['THREAD_POOL_SIZE']) as executor:
                futures = []
                for adapter in wmi_adapters:
                    future = executor.submit(self._get_adapter_details, adapter)
                    futures.append(future)
                for future in futures:
                    try:
                        adapter_info = future.result(timeout=10)
                        if adapter_info:
                            adapters.append(adapter_info)
                    except Exception as e:
                        print(f"获取适配器信息失败: {e}")
                        continue
            logging.info(f"适配器枚举使用: WMI 兜底，找到 {len(adapters)} 个")
        except Exception as e:
            raise Exception(f"获取网络适配器失败: {str(e)}")
        
        return adapters
    
    def _get_adapter_details(self, adapter) -> Optional[Dict]:
        """获取单个适配器的详细信息"""
        # 确保在线程池中也初始化COM组件
        try:
            pythoncom.CoInitialize()
        except:
            pass  # 如果已经初始化过，忽略错误
        
        try:
            connection_id = adapter.NetConnectionID or adapter.Name
            
            adapter_info = {
                'name': adapter.Name,
                'device_id': adapter.DeviceID,
                'mac_address': adapter.MACAddress,
                'alias': connection_id,
                'ip_address': self._get_adapter_ip_fast(connection_id),
                'status': adapter.NetConnectionStatus,
                'speed': self._get_adapter_speed_fast(connection_id),
                'duplex': self._get_adapter_duplex_fast(connection_id)
            }
            return adapter_info
        except Exception as e:
            print(f"获取适配器 {adapter.Name} 详细信息失败: {e}")
            return None
    
    def _get_adapter_ip_fast(self, connection_id: str) -> str:
        """快速获取适配器IP地址"""
        if not connection_id:
            return 'Unknown'
        
        try:
            # 使用更简单的命令
            cmd = f'(Get-NetIPAddress -InterfaceAlias "{connection_id}" -AddressFamily IPv4 -ErrorAction SilentlyContinue | Select-Object -First 1).IPAddress'
            success, result = self._run_powershell_safe(cmd, timeout=5)
            
            if success and result and result != 'Unknown':
                return result
        except:
            pass
        return 'Unknown'
    
    def _get_adapter_speed_fast(self, connection_id: str) -> str:
        """快速获取适配器速度"""
        if not connection_id:
            return 'Unknown'
        
        try:
            # 简化命令，只尝试一种方法
            cmd = f'(Get-NetAdapter -Name "{connection_id}" -ErrorAction SilentlyContinue).LinkSpeed'
            success, result = self._run_powershell_safe(cmd, timeout=5)
            
            # 某些系统返回如"1 Gbps"或"100 Mbps"等字符串，直接返回更直观
            if success and result:
                return result.strip()
        except:
            pass
        return 'Unknown'
    
    def _get_adapter_duplex_fast(self, connection_id: str) -> str:
        """快速获取适配器双工模式"""
        if not connection_id:
            return 'Unknown'
        
        try:
            # 简化命令
            cmd = f'(Get-NetAdapter -Name "{connection_id}" -ErrorAction SilentlyContinue).FullDuplex'
            success, result = self._run_powershell_safe(cmd, timeout=5)
            
            if success and result:
                if result.lower() == 'true':
                    return "全双工"
                elif result.lower() == 'false':
                    return "半双工"
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
    
    def get_speed_duplex_options(self, adapter_name: str = None, use_fallback: bool = True) -> List[str]:
        """动态获取适配器支持的速度双工选项（优化版）"""
        if adapter_name:
            try:
                # 转义适配器名称中的特殊字符
                safe_name = adapter_name.replace('"', '`"').replace("'", "`'")
                
                # 简化命令，减少超时时间
                cmd = f'Get-NetAdapterAdvancedProperty -Name "{safe_name}" -RegistryKeyword "*SpeedDuplex" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty ValidDisplayValues'
                success, result = self._run_powershell_safe(cmd, timeout=6)
                
                if success and result:
                    options = [line.strip() for line in result.split('\n') if line.strip()]
                    if options:
                        return options
                        
            except Exception:
                pass
        
        # 如果无法获取实际选项，根据use_fallback参数决定返回值
        if use_fallback:
            return DEFAULT_SPEED_DUPLEX_OPTIONS.copy()
        else:
            return []
    
    def health_check(self) -> Dict[str, bool]:
        """系统健康检查"""
        results = {
            'wmi_available': False,
            'powershell_available': False,
            'admin_rights': False
        }
        
        # 检查WMI
        try:
            results['wmi_available'] = self._init_wmi_connection()
        except:
            pass
        
        # 检查PowerShell
        try:
            success, _ = self._run_powershell_safe('echo "test"', timeout=3)
            results['powershell_available'] = success
        except:
            pass
        
        # 检查管理员权限
        try:
            import ctypes
            results['admin_rights'] = ctypes.windll.shell32.IsUserAnAdmin()
        except:
            pass
        
        return results
