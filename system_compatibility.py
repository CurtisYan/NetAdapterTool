"""
系统兼容性检查模块
用于检查不同Windows版本和环境的兼容性
"""

import os
import sys
import platform
import subprocess
import logging
from typing import Dict, List, Tuple


class SystemCompatibility:
    """系统兼容性检查器"""
    
    def __init__(self):
        self.system_info = self._get_system_info()
        
    def _get_system_info(self) -> Dict:
        """获取系统信息"""
        try:
            return {
                'platform': platform.platform(),
                'system': platform.system(),
                'release': platform.release(),
                'version': platform.version(),
                'machine': platform.machine(),
                'processor': platform.processor(),
                'python_version': sys.version,
                'is_64bit': sys.maxsize > 2**32,
                'is_admin': self._check_admin_simple()
            }
        except Exception as e:
            logging.warning(f"获取系统信息失败: {e}")
            return {}
    
    def _check_admin_simple(self) -> bool:
        """简单的管理员权限检查"""
        try:
            import ctypes
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    def check_powershell_compatibility(self) -> Dict:
        """检查PowerShell兼容性"""
        result = {
            'available': False,
            'version': 'Unknown',
            'path': None,
            'execution_policy': 'Unknown'
        }
        
        # 检查多种PowerShell路径
        powershell_paths = [
            'powershell',
            r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe',
            r'C:\Program Files\PowerShell\7\pwsh.exe',
            r'C:\Program Files (x86)\PowerShell\7\pwsh.exe'
        ]
        
        for ps_path in powershell_paths:
            try:
                if ps_path != 'powershell' and not os.path.exists(ps_path):
                    continue
                    
                # 测试PowerShell是否可用
                cmd = [ps_path, '-Command', 'echo "test"'] if ps_path != 'powershell' else ['powershell', '-Command', 'echo "test"']
                
                if ps_path == 'powershell':
                    proc = subprocess.run(cmd, capture_output=True, text=True, shell=True, timeout=5)
                else:
                    proc = subprocess.run(cmd, capture_output=True, text=True, timeout=5)
                
                if proc.returncode == 0:
                    result['available'] = True
                    result['path'] = ps_path
                    
                    # 获取版本信息
                    try:
                        version_cmd = [ps_path, '-Command', '$PSVersionTable.PSVersion.ToString()'] if ps_path != 'powershell' else ['powershell', '-Command', '$PSVersionTable.PSVersion.ToString()']
                        if ps_path == 'powershell':
                            version_proc = subprocess.run(version_cmd, capture_output=True, text=True, shell=True, timeout=5)
                        else:
                            version_proc = subprocess.run(version_cmd, capture_output=True, text=True, timeout=5)
                        
                        if version_proc.returncode == 0:
                            result['version'] = version_proc.stdout.strip()
                    except:
                        pass
                    
                    # 获取执行策略
                    try:
                        policy_cmd = [ps_path, '-Command', 'Get-ExecutionPolicy'] if ps_path != 'powershell' else ['powershell', '-Command', 'Get-ExecutionPolicy']
                        if ps_path == 'powershell':
                            policy_proc = subprocess.run(policy_cmd, capture_output=True, text=True, shell=True, timeout=5)
                        else:
                            policy_proc = subprocess.run(policy_cmd, capture_output=True, text=True, timeout=5)
                        
                        if policy_proc.returncode == 0:
                            result['execution_policy'] = policy_proc.stdout.strip()
                    except:
                        pass
                    
                    break
                    
            except Exception as e:
                logging.debug(f"PowerShell路径 {ps_path} 检查失败: {e}")
                continue
        
        return result
    
    def check_wmi_compatibility(self) -> Dict:
        """检查WMI兼容性"""
        result = {
            'available': False,
            'service_running': False,
            'error': None
        }
        
        try:
            # 检查WMI服务状态
            wmi_service_cmd = ['sc', 'query', 'winmgmt']
            proc = subprocess.run(wmi_service_cmd, capture_output=True, text=True, timeout=10)
            
            if proc.returncode == 0 and 'RUNNING' in proc.stdout:
                result['service_running'] = True
            
            # 尝试导入wmi模块
            try:
                import wmi
                # 尝试创建WMI连接
                wmi_conn = wmi.WMI()
                result['available'] = True
            except ImportError:
                result['error'] = 'WMI模块未安装'
            except Exception as e:
                result['error'] = f'WMI连接失败: {str(e)}'
                
        except Exception as e:
            result['error'] = f'WMI检查异常: {str(e)}'
        
        return result
    
    def check_network_commands_compatibility(self) -> Dict:
        """检查网络命令兼容性"""
        result = {
            'netsh_available': False,
            'get_netadapter_available': False,
            'wmic_available': False
        }
        
        # 检查netsh
        try:
            proc = subprocess.run(['netsh', 'interface', 'show', 'interface'], 
                                capture_output=True, text=True, timeout=10)
            result['netsh_available'] = proc.returncode == 0
        except:
            pass
        
        # 检查Get-NetAdapter (PowerShell)
        try:
            proc = subprocess.run(['powershell', '-Command', 'Get-NetAdapter | Select-Object -First 1'], 
                                capture_output=True, text=True, shell=True, timeout=10)
            result['get_netadapter_available'] = proc.returncode == 0
        except:
            pass
        
        # 检查wmic
        try:
            proc = subprocess.run(['wmic', 'path', 'win32_networkadapter', 'get', 'name', '/format:list'], 
                                capture_output=True, text=True, timeout=10)
            result['wmic_available'] = proc.returncode == 0
        except:
            pass
        
        return result
    
    def get_compatibility_report(self) -> Dict:
        """获取完整的兼容性报告"""
        report = {
            'system_info': self.system_info,
            'powershell': self.check_powershell_compatibility(),
            'wmi': self.check_wmi_compatibility(),
            'network_commands': self.check_network_commands_compatibility(),
            'recommendations': []
        }
        
        # 生成建议
        recommendations = []
        
        if not report['powershell']['available']:
            recommendations.append("PowerShell不可用，请检查系统配置或安装PowerShell")
        elif report['powershell']['execution_policy'] == 'Restricted':
            recommendations.append("PowerShell执行策略受限，建议设置为RemoteSigned或Bypass")
        
        if not report['wmi']['available']:
            if not report['wmi']['service_running']:
                recommendations.append("WMI服务未运行，请启动Windows Management Instrumentation服务")
            else:
                recommendations.append("WMI不可用，可能需要安装pywin32或wmi模块")
        
        if not report['network_commands']['get_netadapter_available']:
            recommendations.append("Get-NetAdapter命令不可用，可能需要更新PowerShell或Windows版本")
        
        if not self.system_info.get('is_admin', False):
            recommendations.append("当前未以管理员身份运行，某些功能可能受限")
        
        report['recommendations'] = recommendations
        return report
    
    def print_compatibility_report(self):
        """打印兼容性报告"""
        report = self.get_compatibility_report()
        
        print("=" * 60)
        print("系统兼容性检查报告")
        print("=" * 60)
        
        print(f"\n系统信息:")
        print(f"  平台: {report['system_info'].get('platform', 'Unknown')}")
        print(f"  Python版本: {report['system_info'].get('python_version', 'Unknown').split()[0]}")
        print(f"  管理员权限: {'是' if report['system_info'].get('is_admin', False) else '否'}")
        
        print(f"\nPowerShell兼容性:")
        ps_info = report['powershell']
        print(f"  可用性: {'是' if ps_info['available'] else '否'}")
        if ps_info['available']:
            print(f"  路径: {ps_info['path']}")
            print(f"  版本: {ps_info['version']}")
            print(f"  执行策略: {ps_info['execution_policy']}")
        
        print(f"\nWMI兼容性:")
        wmi_info = report['wmi']
        print(f"  可用性: {'是' if wmi_info['available'] else '否'}")
        print(f"  服务运行: {'是' if wmi_info['service_running'] else '否'}")
        if wmi_info['error']:
            print(f"  错误: {wmi_info['error']}")
        
        print(f"\n网络命令兼容性:")
        net_info = report['network_commands']
        print(f"  netsh: {'可用' if net_info['netsh_available'] else '不可用'}")
        print(f"  Get-NetAdapter: {'可用' if net_info['get_netadapter_available'] else '不可用'}")
        print(f"  wmic: {'可用' if net_info['wmic_available'] else '不可用'}")
        
        if report['recommendations']:
            print(f"\n建议:")
            for i, rec in enumerate(report['recommendations'], 1):
                print(f"  {i}. {rec}")
        
        print("=" * 60)


if __name__ == "__main__":
    # 运行兼容性检查
    checker = SystemCompatibility()
    checker.print_compatibility_report()
