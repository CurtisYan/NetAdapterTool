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
        """检查当前是否有管理员权限"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin()
        except:
            return False
    
    def _run_powershell_command(self, command: str) -> Tuple[bool, str]:
        """执行PowerShell命令并返回结果"""
        try:
            result = subprocess.run(['powershell', '-Command', command], 
                                  capture_output=True, text=True, shell=True)
            
            if result.returncode == 0:
                return True, result.stdout.strip()
            else:
                return False, result.stderr.strip()
        except Exception as e:
            return False, str(e)
    
    def set_adapter_speed_duplex(self, adapter_name: str, speed_duplex: str) -> Tuple[bool, str]:
        """设置适配器的速度和双工模式"""
        if not self.is_admin:
            return False, "需要管理员权限才能修改网络设置"
        
        try:
            # 先尝试使用 RegistryKeyword，再回退到 DisplayName 匹配
            commands = [
                f'Set-NetAdapterAdvancedProperty -Name "{adapter_name}" -RegistryKeyword "*SpeedDuplex" -DisplayValue "{speed_duplex}"',
                f'Set-NetAdapterAdvancedProperty -Name "{adapter_name}" -DisplayName "*Speed*Duplex*" -DisplayValue "{speed_duplex}"'
            ]
            last_err = ''
            for command in commands:
                success, message = self._run_powershell_command(command)
                if success:
                    return True, f"成功设置 {adapter_name} 的网络设置为 {speed_duplex}"
                last_err = message
            # 两种方式都失败，返回更清晰的错误并提示可用值
            tips_cmd = (
                f'Get-NetAdapterAdvancedProperty -Name "{adapter_name}" | '
                f'Where-Object {{$_.RegistryKeyword -like "*Speed*" -or $_.DisplayName -like "*Duplex*" -or $_.DisplayName -like "*Speed*"}} | '
                'Select-Object -Property DisplayName, RegistryKeyword, DisplayValue | Format-Table -AutoSize'
            )
            _ok, tips = self._run_powershell_command(tips_cmd)
            hint = f"\n\n可用的相关高级属性如下(供排查):\n{tips}" if tips else ''
            return False, f"设置失败: {last_err}{hint}"
                
        except Exception as e:
            return False, f"设置失败: {str(e)}"
    
    def _parse_speed_duplex(self, speed_duplex: str) -> Tuple[int, str]:
        """解析速度和双工设置"""
        if "半双工" in speed_duplex:
            duplex = "Half"
        elif "全双工" in speed_duplex:
            duplex = "Full"
        else:
            duplex = "Auto"
        
        if "10 Mbps" in speed_duplex:
            speed = 10
        elif "100 Mbps" in speed_duplex:
            speed = 100
        elif "1.0 Gbps" in speed_duplex or "1000 Mbps" in speed_duplex:
            speed = 1000
        else:
            speed = 0  # 自动协商
        
        return speed, duplex
    
    def get_valid_speed_duplex_options(self, adapter_name: str) -> List[str]:
        """获取适配器支持的速度双工选项"""
        try:
            commands = [
                f'Get-NetAdapterAdvancedProperty -Name "{adapter_name}" -RegistryKeyword "*SpeedDuplex" | Select-Object -ExpandProperty ValidDisplayValues',
                f'Get-NetAdapterAdvancedProperty -Name "{adapter_name}" -DisplayName "*Speed*Duplex*" | Select-Object -ExpandProperty ValidDisplayValues'
            ]
            for command in commands:
                result = subprocess.run(['powershell', '-Command', command], 
                                      capture_output=True, text=True, shell=True)
                if result.returncode == 0 and result.stdout.strip():
                    options = [line.strip() for line in result.stdout.strip().split('\n') if line.strip()]
                    if options:
                        return options
        except:
            pass
        
        # 默认选项
        return ["自动侦测", "10 Mbps 半双工", "10 Mbps 全双工", "100 Mbps 半双工", "100 Mbps 全双工", "1.0 Gbps 全双工"]
    
    def get_current_speed_duplex(self, adapter_name: str) -> str:
        """获取当前的速度双工设置"""
        try:
            # 使用 Get-NetAdapterAdvancedProperty 获取当前设置
            commands = [
                f'Get-NetAdapterAdvancedProperty -Name "{adapter_name}" -RegistryKeyword "*SpeedDuplex" | Select-Object -ExpandProperty DisplayValue',
                f'Get-NetAdapterAdvancedProperty -Name "{adapter_name}" -DisplayName "*Speed*Duplex*" | Select-Object -ExpandProperty DisplayValue'
            ]
            for command in commands:
                result = subprocess.run(['powershell', '-Command', command], 
                                      capture_output=True, text=True, shell=True)
                if result.returncode == 0 and result.stdout.strip():
                    return result.stdout.strip()
        except:
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
