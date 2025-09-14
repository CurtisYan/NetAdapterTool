"""
网络适配器管理工具图形界面
使用PyQt5实现简洁的GUI界面
"""

import sys
import os
import ctypes
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QComboBox, QPushButton, 
                             QGroupBox, QMessageBox, QProgressBar)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont
from network_adapter import NetworkAdapter
from network_settings import NetworkSettings


class WorkerThread(QThread):
    """后台工作线程，避免界面卡死"""
    finished = pyqtSignal(bool, str, list)  # 添加适配器列表参数
    progress_update = pyqtSignal(str)  # 进度更新信号
    
    def __init__(self, settings, adapter, adapter_name, speed_duplex):
        super().__init__()
        self.settings = settings
        self.adapter = adapter
        self.adapter_name = adapter_name
        self.speed_duplex = speed_duplex
    
    def run(self):
        try:
            # 第一步：应用设置
            self.progress_update.emit("正在应用网络设置...")
            success, message = self.settings.set_adapter_speed_duplex(
                self.adapter_name, self.speed_duplex)
            
            if success:
                # 第二步：刷新适配器列表
                self.progress_update.emit("正在刷新适配器状态...")
                try:
                    updated_adapters = self.adapter.get_all_adapters()
                    self.finished.emit(True, message, updated_adapters)
                except Exception as refresh_error:
                    # 即使刷新失败，设置应用也是成功的，但不在成功消息中显示错误
                    print(f"刷新适配器状态失败: {str(refresh_error)}")
                    self.finished.emit(True, message, [])
            else:
                self.finished.emit(False, message, [])
                
        except Exception as e:
            self.finished.emit(False, f"操作失败: {str(e)}", [])


class NetworkAdapterGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.adapter = NetworkAdapter()
        self.settings = NetworkSettings()
        self.current_adapters = []
        self.worker_thread = None
        
        self.init_ui()
        self.refresh_adapters()
    
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("网络适配器管理工具")
        self.setFixedSize(500, 400)
        
        # 创建中央widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        
        # 标题
        title_label = QLabel("网络适配器管理工具")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 适配器选择组
        adapter_group = QGroupBox("适配器选择")
        adapter_layout = QVBoxLayout(adapter_group)
        
        self.adapter_combo = QComboBox()
        self.adapter_combo.currentTextChanged.connect(self.on_adapter_changed)
        self.adapter_combo.setMinimumHeight(35)  # 增大高度
        self.adapter_combo.setStyleSheet("QComboBox { font-size: 12px; padding: 5px; } QComboBox::drop-down { width: 25px; } QComboBox QAbstractItemView { min-width: 400px; }")
        adapter_layout.addWidget(self.adapter_combo)
        
        # 当前状态显示
        self.status_label = QLabel("当前状态: 未选择适配器")
        adapter_layout.addWidget(self.status_label)
        
        main_layout.addWidget(adapter_group)
        
        # 设置组
        settings_group = QGroupBox("网络设置")
        settings_layout = QVBoxLayout(settings_group)
        
        # 速度双工设置（合并为一个选项）
        speed_duplex_layout = QHBoxLayout()
        speed_duplex_layout.addWidget(QLabel("速度和双工:"))
        self.speed_duplex_combo = QComboBox()
        # 初始化时使用默认选项
        self.speed_duplex_combo.addItems(self.adapter.get_speed_duplex_options())
        self.speed_duplex_combo.setCurrentText("自动侦测")  # 设置默认值
        self.speed_duplex_combo.setMinimumHeight(35)  # 增大高度
        self.speed_duplex_combo.setStyleSheet("QComboBox { font-size: 12px; padding: 5px; } QComboBox::drop-down { width: 25px; } QComboBox QAbstractItemView { min-width: 250px; }")
        speed_duplex_layout.addWidget(self.speed_duplex_combo)
        settings_layout.addLayout(speed_duplex_layout)
        
        main_layout.addWidget(settings_group)
        
        # 按钮组
        button_layout = QHBoxLayout()
        
        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.clicked.connect(self.refresh_adapters)
        button_layout.addWidget(self.refresh_btn)
        
        self.apply_btn = QPushButton("应用设置")
        self.apply_btn.clicked.connect(self.apply_settings)
        self.apply_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; }")
        button_layout.addWidget(self.apply_btn)
        
        main_layout.addLayout(button_layout)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # 状态栏
        if self.settings.is_admin:
            self.statusBar().showMessage("就绪 - 管理员模式")
        else:
            self.statusBar().showMessage("就绪 - 普通用户模式（修改设置需要提权）")
    
    def refresh_adapters(self):
        """刷新适配器列表"""
        try:
            self.statusBar().showMessage("正在刷新适配器列表...")
            self.current_adapters = self.adapter.get_all_adapters()
            self.update_adapter_list(self.current_adapters)
        except Exception as e:
            QMessageBox.critical(self, "错误", f"刷新适配器失败: {str(e)}")
            self.statusBar().showMessage("刷新失败")
    
    def update_adapter_list(self, adapters):
        """更新适配器列表显示"""
        # 保存当前选中的适配器
        current_selection = self.adapter_combo.currentText()
        
        self.current_adapters = adapters
        self.adapter_combo.clear()
        
        if self.current_adapters:
            for adapter in self.current_adapters:
                # 显示设备名，内部保存接口别名（InterfaceAlias）
                self.adapter_combo.addItem(adapter['name'], userData=adapter.get('alias'))
            
            # 尝试恢复之前的选择
            if current_selection:
                index = self.adapter_combo.findText(current_selection)
                if index >= 0:
                    self.adapter_combo.setCurrentIndex(index)
            
            self.statusBar().showMessage(f"找到 {len(self.current_adapters)} 个适配器")
        else:
            self.statusBar().showMessage("未找到可用的网络适配器")
            self.status_label.setText("当前状态: 未找到适配器")
    
    def on_adapter_changed(self):
        """适配器选择改变时的处理"""
        current_text = self.adapter_combo.currentText()
        current_alias = self.adapter_combo.currentData()
        if not current_text:
            return
        
        # 找到对应的适配器信息并获取实际的速度双工设置
        for adapter in self.current_adapters:
            if adapter['name'] == current_text:
                # 获取实际的速度双工设置
                alias = current_alias or adapter.get('alias') or adapter['name']
                actual_speed_duplex = self.settings.get_current_speed_duplex(alias)
                status_text = (f"当前状态: {actual_speed_duplex} | "
                             f"IP: {adapter['ip_address']}")
                self.status_label.setText(status_text)
                
                # 更新下拉菜单选项为当前适配器支持的选项
                self.update_speed_duplex_options(alias)
                break
    
    def apply_settings(self):
        """应用网络设置"""
        if not self.adapter_combo.currentText():
            QMessageBox.warning(self, "警告", "请先选择一个网络适配器")
            return
        
        if not self.settings.is_admin:
            reply = QMessageBox.question(self, "需要管理员权限", 
                                       "修改网络设置需要管理员权限\n\n"
                                       "是否要以管理员身份重新启动程序？",
                                       QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.restart_as_admin()
            return
        
        # 获取设置值
        adapter_name = self.adapter_combo.currentText()  # 显示用途
        adapter_alias = self.adapter_combo.currentData() or adapter_name
        speed_duplex = self.speed_duplex_combo.currentText()
        
        # 确认对话框
        reply = QMessageBox.question(self, "确认操作", 
                                   f"确定要修改适配器 '{adapter_name}' 的设置吗?\n\n"
                                   f"速度和双工模式: {speed_duplex}",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            # 后台操作用别名，以确保 PowerShell -Name 匹配接口别名
            self.start_operation(adapter_alias, speed_duplex)
    
    def start_operation(self, adapter_name, speed_duplex):
        """启动后台操作"""
        self.apply_btn.setEnabled(False)
        self.refresh_btn.setEnabled(False)  # 禁用刷新按钮
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # 无限进度条
        self.statusBar().showMessage("正在应用设置...")
        
        # 启动工作线程
        self.worker_thread = WorkerThread(self.settings, self.adapter, adapter_name, speed_duplex)
        self.worker_thread.finished.connect(self.on_operation_finished)
        self.worker_thread.progress_update.connect(self.on_progress_update)
        self.worker_thread.start()
    
    def update_speed_duplex_options(self, adapter_alias: str):
        """更新速度双工选项为当前适配器支持的选项"""
        try:
            # 获取当前选中的值
            current_selection = self.speed_duplex_combo.currentText()
            
            # 获取适配器支持的选项
            options = self.adapter.get_speed_duplex_options(adapter_alias)
            
            # 更新下拉菜单
            self.speed_duplex_combo.clear()
            self.speed_duplex_combo.addItems(options)
            
            # 尝试保持之前的选择，如果不存在则选择第一个
            if current_selection in options:
                self.speed_duplex_combo.setCurrentText(current_selection)
            else:
                self.speed_duplex_combo.setCurrentIndex(0)
                
        except Exception as e:
            print(f"更新选项失败: {e}")
    
    def on_progress_update(self, message):
        """更新进度信息"""
        self.statusBar().showMessage(message)
    
    def on_operation_finished(self, success, message, updated_adapters):
        """操作完成的处理"""
        # 恢复界面状态
        self.apply_btn.setEnabled(True)
        self.refresh_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if success:
            # 更新适配器列表（如果有数据）
            if updated_adapters:
                self.update_adapter_list(updated_adapters)
            
            self.statusBar().showMessage("设置应用成功")
            QMessageBox.information(self, "成功", message)
        else:
            self.statusBar().showMessage("设置应用失败")
            QMessageBox.critical(self, "失败", message)
    
    def restart_as_admin(self):
        """以管理员身份重启程序"""
        try:
            # 获取当前程序路径
            if getattr(sys, 'frozen', False):
                # 如果是打包后的exe文件
                current_exe = sys.executable
            else:
                # 如果是Python脚本
                current_exe = sys.executable
                script_path = os.path.abspath(__file__)
            
            # 使用ShellExecuteW以管理员身份启动
            if getattr(sys, 'frozen', False):
                # 打包后的exe
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", current_exe, None, None, 1
                )
            else:
                # Python脚本
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", current_exe, f'"{script_path}"', None, 1
                )
            
            # 关闭当前程序
            self.close()
            QApplication.quit()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"无法以管理员身份启动程序: {str(e)}")
    
    def closeEvent(self, event):
        """关闭程序时的处理"""
        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.quit()
            self.worker_thread.wait()
        event.accept()


def main():
    app = QApplication(sys.argv)
    
    # 设置应用程序信息
    app.setApplicationName("网络适配器管理工具")
    app.setApplicationVersion("1.0")
    
    # 创建主窗口
    window = NetworkAdapterGUI()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
