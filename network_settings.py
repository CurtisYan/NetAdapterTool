"""
网络适配器设置修改模块
用于修改Windows系统中网络适配器的速度和双工模式
"""

import subprocess
import ctypes
import sys
from typing import Optional, Tuple, List


class NetworkSettings:
    def __init__(self):
        self.is_admin = self._check_admin_rights()
    
    def _check_admin_rights(self) -> bool:
        """检查当前是否有管理员权限，支持多种检测方法"""
        try:
            # 方法1：使用ctypes检查
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            try:
                # 方法2：尝试访问需要管理员权限的注册表项
                import winreg
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                   "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\System",
                                   0, winreg.KEY_READ)
                winreg.CloseKey(key)
                return True
            except:
                try:
                    # 方法3：尝试创建临时文件到系统目录
                    import tempfile
                    temp_file = os.path.join(os.environ.get('SYSTEMROOT', 'C:\\Windows'), 'temp_admin_test.tmp')
                    with open(temp_file, 'w') as f:
                        f.write('test')
                    os.remove(temp_file)
                    return True
                except:
                    return False
    
    def _run_powershell_command(self, command: str) -> Tuple[bool, str]:
        """执行PowerShell命令并返回结果，支持多种Windows版本"""
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
                        timeout=10,
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
                        timeout=10,
                        creationflags=subprocess.CREATE_NO_WINDOW
                    )
                
                if result.returncode == 0:
                    return True, result.stdout.strip()
                else:
                    # 如果是第一个路径失败，尝试下一个
                    if ps_path == powershell_paths[0]:
                        continue
                    error_msg = result.stderr.strip() or result.stdout.strip() or "PowerShell命令执行失败"
                    return False, error_msg
                    
            except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
                # 尝试下一个PowerShell路径
                continue
            except Exception as e:
                # 如果是最后一个路径，返回错误
                if ps_path == powershell_paths[-1]:
                    return False, f"PowerShell命令执行异常: {str(e)}"
                continue
        
        return False, "未找到可用的PowerShell"
    
    def set_adapter_speed_duplex(self, adapter_name: str, speed_duplex: str) -> Tuple[bool, str]:
        """设置适配器的速度和双工模式"""
        if not self.is_admin:
            return False, "需要管理员权限才能修改网络设置"
        
        # 验证和清理输入参数
        if not adapter_name or not adapter_name.strip():
            return False, "适配器名称不能为空"
        
        if not speed_duplex or not speed_duplex.strip():
            return False, "速度双工设置不能为空"
        
        try:
            # 转义特殊字符
            safe_adapter_name = adapter_name.replace('"', '`"').replace("'", "`'")
            safe_speed_duplex = speed_duplex.replace('"', '`"').replace("'", "`'")
            
            # 先尝试使用 RegistryKeyword，再回退到 DisplayName 匹配
            commands = [
                f'Set-NetAdapterAdvancedProperty -Name "{safe_adapter_name}" -RegistryKeyword "*SpeedDuplex" -DisplayValue "{safe_speed_duplex}"',
                f'Set-NetAdapterAdvancedProperty -Name "{safe_adapter_name}" -DisplayName "*Speed*Duplex*" -DisplayValue "{safe_speed_duplex}"'
            ]
            last_err = ''
            for command in commands:
                success, message = self._run_powershell_command(command)
                if success:
                    return True, f"成功设置 {adapter_name} 的网络设置为 {speed_duplex}"
                last_err = message
            # 两种方式都失败，返回更清晰的错误并提示可用值
            tips_cmd = (
                f'Get-NetAdapterAdvancedProperty -Name "{safe_adapter_name}" | '
                f'Where-Object {{$_.RegistryKeyword -like "*Speed*" -or $_.DisplayName -like "*Duplex*" -or $_.DisplayName -like "*Speed*"}} | '
                'Select-Object -Property DisplayName, RegistryKeyword, DisplayValue | Format-Table -AutoSize'
            )
            _ok, tips = self._run_powershell_command(tips_cmd)
            hint = f"\n\n可用的相关高级属性如下(供排查):\n{tips}" if tips else ''
            return False, f"设置失败: {last_err}{hint}"
                
        except Exception as e:
            return False, f"设置失败: {str(e)}"
    
    def get_valid_speed_duplex_options(self, adapter_name: str) -> List[str]:
        """获取适配器支持的速度双工选项
        
        注意：此方法已弃用，建议直接使用 NetworkAdapter.get_speed_duplex_options()
        为了避免循环导入，这里使用延迟导入
        """
        try:
            # 延迟导入避免循环依赖
            from network_adapter import DEFAULT_SPEED_DUPLEX_OPTIONS
            
            # 转义适配器名称中的特殊字符
            safe_name = adapter_name.replace('"', '`"').replace("'", "`'")
            
            # 尝试多种方法获取选项
            commands = [
                f'Get-NetAdapterAdvancedProperty -Name "{safe_name}" -RegistryKeyword "*SpeedDuplex" | Select-Object -ExpandProperty ValidDisplayValues',
                f'Get-NetAdapterAdvancedProperty -Name "{safe_name}" -DisplayName "*Speed*Duplex*" | Select-Object -ExpandProperty ValidDisplayValues'
            ]
            
            for command in commands:
                success, result = self._run_powershell_command(command)
                if success and result.strip():
                    options = [line.strip() for line in result.strip().split('\n') if line.strip()]
                    if options:
                        return options
        except Exception:
            pass
        
        # 使用统一的默认选项
        try:
            from network_adapter import DEFAULT_SPEED_DUPLEX_OPTIONS
            return DEFAULT_SPEED_DUPLEX_OPTIONS.copy()
        except ImportError:
            # 如果导入失败，使用本地默认值
            return ["自动侦测", "10 Mbps 半双工", "10 Mbps 全双工", "100 Mbps 半双工", "100 Mbps 全双工", "1.0 Gbps 全双工"]
    
    def get_current_speed_duplex(self, adapter_name: str) -> str:
        """获取当前的速度双工设置"""
        if not adapter_name or not adapter_name.strip():
            return "Unknown"
            
        try:
            # 转义特殊字符
            safe_adapter_name = adapter_name.replace('"', '`"').replace("'", "`'")
            
            # 使用 Get-NetAdapterAdvancedProperty 获取当前设置
            commands = [
                f'Get-NetAdapterAdvancedProperty -Name "{safe_adapter_name}" -RegistryKeyword "*SpeedDuplex" | Select-Object -ExpandProperty DisplayValue',
                f'Get-NetAdapterAdvancedProperty -Name "{safe_adapter_name}" -DisplayName "*Speed*Duplex*" | Select-Object -ExpandProperty DisplayValue'
            ]
            for command in commands:
                success, result = self._run_powershell_command(command)
                if success and result.strip():
                    return result.strip()
        except (subprocess.TimeoutExpired, subprocess.SubprocessError):
            pass
        return "Unknown"
    
    def restart_adapter(self, adapter_name: str) -> Tuple[bool, str]:
        """重启网络适配器使设置生效"""
        if not self.is_admin:
            return False, "需要管理员权限才能重启网络适配器"
        
        # 禁用适配器
        disable_cmd = f'Disable-NetAdapter -Name "{adapter_name}" -Confirm:$false'
        success, message = self._run_powershell_command(disable_cmd)
        
        if not success:
            return False, f"禁用适配器失败: {message}"
        
        # 启用适配器
        enable_cmd = f'Enable-NetAdapter -Name "{adapter_name}" -Confirm:$false'
        success, message = self._run_powershell_command(enable_cmd)
        
        if success:
            return True, f"成功重启适配器 {adapter_name}"
        else:
            return False, f"启用适配器失败: {message}"
    
    def request_admin_rights(self):
        """请求管理员权限重新启动程序"""
        if not self.is_admin:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
