"""
优化版网络适配器管理工具图形界面
解决卡死问题，提升启动速度和响应性
"""

import sys
import os
import ctypes
import logging
import pythoncom
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QComboBox, QPushButton, 
                             QGroupBox, QMessageBox, QProgressBar, QTextEdit, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QPixmap, QIcon

# 导入网络适配器模块
from network_adapter import NetworkAdapter
from network_settings import NetworkSettings
from system_compatibility import SystemCompatibility

# 自定义日志处理器，用于捕获日志到GUI
class GuiLogHandler(logging.Handler):
    def __init__(self):
        super().__init__()
        self.log_messages = []
        self.gui_callback = None
    
    def emit(self, record):
        log_entry = self.format(record)
        self.log_messages.append(log_entry)
        # 如果GUI已经初始化，实时更新显示
        if self.gui_callback:
            self.gui_callback(log_entry)
    
    def set_gui_callback(self, callback):
        self.gui_callback = callback
    
    def get_all_logs(self):
        return '\n'.join(self.log_messages)

# 创建全局日志处理器实例
gui_log_handler = GuiLogHandler()

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(sys.stdout),
        gui_log_handler
    ]
)


class InitializationThread(QThread):
    """初始化线程，避免主线程阻塞"""
    finished = pyqtSignal(bool, str, object)  # 成功标志，错误信息，适配器对象
    progress_update = pyqtSignal(str)  # 进度更新信号
    
    def __init__(self):
        super().__init__()
    
    def run(self):
        # 每个线程必须初始化COM
        try:
            try:
                pythoncom.CoInitialize()
            except:
                pass
            self.progress_update.emit("正在初始化网络适配器模块...")
            
            # 创建适配器对象（延迟初始化）
            adapter = NetworkAdapter(lazy_init=True)
            self.progress_update.emit("正在进行系统健康检查...")
            
            # 进行健康检查
            health = adapter.health_check()
            if not health['wmi_available']:
                self.finished.emit(False, "WMI服务不可用，请检查系统配置", None)
                return
            if not health['powershell_available']:
                self.finished.emit(False, "PowerShell不可用，请检查系统配置", None)
                return
            
            self.progress_update.emit("初始化完成")
            self.finished.emit(True, "", adapter)
            
        except Exception as e:
            error_msg = f"初始化失败: {str(e)}"
            logging.error(error_msg)
            self.finished.emit(False, error_msg, None)
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass


class WorkerThread(QThread):
    """后台工作线程，避免界面卡死"""
    finished = pyqtSignal(bool, str, list)
    progress_update = pyqtSignal(str)
    
    def __init__(self, settings, adapter, adapter_name, speed_duplex):
        super().__init__()
        self.settings = settings
        self.adapter = adapter
        self.adapter_name = adapter_name
        self.speed_duplex = speed_duplex
    
    def run(self):
        # 每个线程必须初始化COM
        try:
            try:
                pythoncom.CoInitialize()
            except:
                pass
            
            logging.info(f"开始应用网络设置: {self.adapter_name} -> {self.speed_duplex}")
            self.progress_update.emit("正在应用网络设置...")
            
            success, message = self.settings.set_adapter_speed_duplex(
                self.adapter_name, self.speed_duplex)
            
            if success:
                logging.info("网络设置应用成功，等待网络适配器重新初始化")
                self.progress_update.emit("等待网络适配器重新初始化...")
                import time
                time.sleep(2)  # 减少等待时间
                
                logging.info("网络设置应用成功，获取更新状态")
                try:
                    updated_status = self.settings.get_current_speed_duplex(self.adapter_name)
                    self.finished.emit(True, message, [{'adapter_name': self.adapter_name, 'new_status': updated_status}])
                except Exception as status_error:
                    logging.warning(f"获取更新状态失败: {str(status_error)}")
                    self.finished.emit(True, message, [])
            else:
                logging.error(f"网络设置应用失败: {message}")
                self.finished.emit(False, message, [])
                
        except Exception as e:
            logging.error(f"操作异常: {str(e)}")
            self.finished.emit(False, f"操作失败: {str(e)}", [])
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass


class RefreshThread(QThread):
    """刷新适配器的后台线程"""
    finished = pyqtSignal(bool, str, list)
    progress_update = pyqtSignal(str)
    
    def __init__(self, adapter):
        super().__init__()
        self.adapter = adapter
    
    def run(self):
        # 每个线程必须初始化COM
        try:
            try:
                pythoncom.CoInitialize()
            except:
                pass
            logging.info("开始刷新适配器列表")
            self.progress_update.emit("正在刷新适配器列表...")
            adapters = self.adapter.get_all_adapters()
            logging.info(f"刷新完成，找到 {len(adapters)} 个适配器")
            self.finished.emit(True, "", adapters)
        except Exception as e:
            error_msg = f"刷新适配器失败: {str(e)}"
            logging.error(error_msg)
            self.finished.emit(False, error_msg, [])
        finally:
            try:
                pythoncom.CoUninitialize()
            except:
                pass


class NetworkAdapterGUI(QMainWindow):
    def __init__(self):
        try:
            print("Initializing NetworkAdapterGUI...")
            super().__init__()
            
            # 初始化变量
            print("Setting up variables...")
            self.adapter = None
            self.settings = NetworkSettings()
            self.current_adapters = []
            self.worker_thread = None
            self.refresh_thread = None
            self.init_thread = None
            self.log_visible = False
            self.initialization_complete = False
            # 动态刷新状态
            self._dynamic_refresh_active = False
            self._dynamic_target_alias = None
            self._dynamic_target_value = None
            self._dynamic_attempt_idx = 0
            # 延后提示相关
            self._pending_success_message = None
            print("Variables initialized successfully")
            
            # 先初始化UI
            print("Initializing UI...")
            self.init_ui()
            print("UI initialized successfully")
            
            # 设置日志回调
            print("Setting up logging...")
            gui_log_handler.set_gui_callback(self.append_log_message)
            
            logging.info(f"程序启动 - 管理员模式: {self.settings.is_admin}")
            logging.info("网络适配器管理工具已启动")
            
            # 显示启动状态
            print("Setting up initial state...")
            self.statusBar().showMessage("正在初始化...")
            self.refresh_btn.setEnabled(False)
            self.apply_btn.setEnabled(False)
            
            # 检查管理员权限：若无管理员，直接静默提权重启并退出当前进程
            if not self.settings.is_admin:
                QTimer.singleShot(50, lambda: self.restart_as_admin(silent=True))
                return
            
            # 启动后台初始化
            print("Starting background initialization...")
            self.start_initialization()
            print("NetworkAdapterGUI initialization completed")
            
        except Exception as e:
            print(f"ERROR in NetworkAdapterGUI.__init__(): {e}")
            import traceback
            traceback.print_exc()
            raise
    
    def closeEvent(self, event):
        """程序关闭时清理资源"""
        try:
            # 停止所有后台线程
            threads_to_stop = [
                ('worker_thread', self.worker_thread),
                ('refresh_thread', self.refresh_thread), 
                ('init_thread', self.init_thread)
            ]
            
            for thread_name, thread_obj in threads_to_stop:
                if thread_obj and thread_obj.isRunning():
                    print(f"Stopping {thread_name}...")
                    thread_obj.quit()
                    if not thread_obj.wait(3000):  # 等待3秒
                        print(f"Warning: {thread_name} did not stop gracefully")
                        thread_obj.terminate()
            
            # 清理WMI连接
            if self.adapter:
                try:
                    if hasattr(self.adapter, 'wmi_conn') and self.adapter.wmi_conn:
                        self.adapter.wmi_conn = None
                except:
                    pass
            
            print("Resource cleanup completed")
            event.accept()
            
        except Exception as e:
            print(f"Error during cleanup: {e}")
            event.accept()
    
    def start_initialization(self):
        """启动后台初始化"""
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # 无限进度条
        
        self.init_thread = InitializationThread()
        self.init_thread.finished.connect(self.on_initialization_finished)
        self.init_thread.progress_update.connect(self.on_progress_update)
        self.init_thread.start()
    
    def on_initialization_finished(self, success, error_msg, adapter):
        """初始化完成处理"""
        self.progress_bar.setVisible(False)
        
        if success:
            self.adapter = adapter
            self.initialization_complete = True
            self.refresh_btn.setEnabled(True)
            self.apply_btn.setEnabled(True)
            self.statusBar().showMessage("初始化完成 - 就绪")
            
            # 自动刷新适配器列表
            QTimer.singleShot(500, self.refresh_adapters)
            
        else:
            logging.error(f"初始化失败: {error_msg}")
            self.statusBar().showMessage("初始化失败")
            
            # 如果是WMI/权限相关问题且不是管理员，直接静默提权重启
            if ("WMI" in error_msg or "权限" in error_msg) and not self.settings.is_admin:
                self.restart_as_admin(silent=True)
                return
            
            # 其他错误：提供详细诊断信息
            try:
                compatibility = SystemCompatibility()
                report = compatibility.get_compatibility_report()
                
                # 构建详细错误信息
                detailed_msg = f"初始化失败: {error_msg}\n\n"
                detailed_msg += "系统诊断信息:\n"
                detailed_msg += f"• PowerShell: {'可用' if report['powershell']['available'] else '不可用'}\n"
                detailed_msg += f"• WMI: {'可用' if report['wmi']['available'] else '不可用'}\n"
                detailed_msg += f"• 管理员权限: {'是' if report['system_info'].get('is_admin', False) else '否'}\n"
                
                if report['recommendations']:
                    detailed_msg += "\n建议:\n"
                    for i, rec in enumerate(report['recommendations'][:3], 1):  # 只显示前3个建议
                        detailed_msg += f"{i}. {rec}\n"
                
                detailed_msg += "\n是否要重试？"
                
                reply = QMessageBox.critical(self, "初始化失败", detailed_msg,
                                           QMessageBox.Retry | QMessageBox.Close)
            except Exception:
                # 如果兼容性检查也失败，使用简单错误信息
                reply = QMessageBox.critical(self, "初始化失败", 
                                           f"{error_msg}\n\n是否要重试？",
                                           QMessageBox.Retry | QMessageBox.Close)
            
            if reply == QMessageBox.Retry:
                QTimer.singleShot(1000, self.start_initialization)
            else:
                self.close()
    
    def show_admin_warning(self):
        """显示管理员权限警告，提供自动重启选项"""
        if not self.settings.is_admin:
            reply = QMessageBox.question(
                self, 
                "需要管理员权限", 
                "网络适配器管理需要管理员权限才能正常工作。\n\n"
                "没有管理员权限将无法：\n"
                "• 获取网络适配器信息\n"
                "• 修改网络设置\n"
                "• 查看详细状态\n\n"
                "是否要以管理员身份重新启动程序？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes  # 默认选择Yes
            )
            
            if reply == QMessageBox.Yes:
                self.restart_as_admin()
            else:
                # 用户选择继续，但功能受限
                self.statusBar().showMessage("功能受限模式 - 建议以管理员身份运行")
    
    def restart_as_admin(self, silent: bool = False):
        """以管理员身份重启程序。
        silent=True 时不弹窗，尽量使用 pythonw.exe 以避免命令行窗口。
        """
        try:
            import ctypes
            import sys
            
            # 获取当前程序路径
            if getattr(sys, 'frozen', False):
                # 如果是打包后的exe文件
                current_exe = sys.executable
                args = None
            else:
                # 如果是Python脚本
                # 优先使用 pythonw.exe 避免命令行窗口
                pyexe = sys.executable
                pywexe = os.path.join(os.path.dirname(pyexe), 'pythonw.exe')
                current_exe = pywexe if os.path.exists(pywexe) else pyexe
                script_path = os.path.abspath(__file__)
                args = f'"{script_path}"'
            
            # 关闭当前程序
            self.close()
            QApplication.processEvents()
            
            # 使用ShellExecuteW以管理员身份启动
            # 显示状态：0隐藏窗口，1正常显示
            show_cmd = 0 if (not getattr(sys, 'frozen', False)) else 1
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", current_exe, args, None, show_cmd
            )
            
            # 退出当前程序
            QApplication.quit()
            sys.exit(0)
            
        except Exception as e:
            if not silent:
                # 如果启动失败，显示错误
                QMessageBox.critical(self, "错误", f"无法以管理员身份启动程序: {str(e)}")
                self.show()
    
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("网络适配器管理工具")
        self.setMinimumSize(500, 600)
        self.setMaximumWidth(500)
        self.resize(500, 600)
        
        # 创建菜单栏
        menubar = self.menuBar()
        
        # 帮助菜单
        help_menu = menubar.addMenu('帮助')
        
        # 系统诊断动作
        system_diag_action = help_menu.addAction('系统诊断')
        system_diag_action.triggered.connect(self.show_system_diagnosis)
        
        # 关于动作
        about_action = help_menu.addAction('关于')
        about_action.triggered.connect(self.show_about)
        
        # 设置窗口图标
        try:
            # 支持多种部署方式的资源路径查找
            possible_paths = [
                os.path.join(os.path.dirname(__file__), "img", "NA.ico"),  # 源码运行
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "img", "NA.ico"),  # 绝对路径
                os.path.join(os.getcwd(), "img", "NA.ico"),  # 当前工作目录
                "img/NA.ico",  # 相对路径
                "NA.ico"  # 同目录
            ]
            
            for icon_path in possible_paths:
                if os.path.exists(icon_path):
                    self.setWindowIcon(QIcon(icon_path))
                    break
        except Exception as e:
            logging.warning(f"加载窗口图标失败: {str(e)}")
        
        # 创建中央widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 主布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        
        # 标题和Logo区域
        header_layout = QVBoxLayout()
        header_layout.setSpacing(10)
        
        # Logo
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)
        try:
            # 支持多种部署方式的Logo路径查找
            possible_logo_paths = [
                os.path.join(os.path.dirname(__file__), "img", "NA (蓝透明).jpg"),  # 源码运行
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "img", "NA (蓝透明).jpg"),  # 绝对路径
                os.path.join(os.getcwd(), "img", "NA (蓝透明).jpg"),  # 当前工作目录
                "img/NA (蓝透明).jpg",  # 相对路径
                "NA (蓝透明).jpg"  # 同目录
            ]
            
            logo_loaded = False
            for logo_path in possible_logo_paths:
                if os.path.exists(logo_path):
                    pixmap = QPixmap(logo_path)
                    if not pixmap.isNull():
                        scaled_pixmap = pixmap.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                        logo_label.setPixmap(scaled_pixmap)
                        logo_label.setStyleSheet("margin: 15px 0;")
                        logo_loaded = True
                        break
            
            if not logo_loaded:
                raise Exception("未找到Logo文件")
        except Exception:
            # 使用文本作为备用Logo
            logo_label.setText("🔧")
            logo_label.setFont(QFont("", 48))
            logo_label.setStyleSheet("color: #2196F3; margin: 15px 0;")
        
        header_layout.addWidget(logo_label)
        
        # 标题
        title_label = QLabel("网络适配器管理工具")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet("color: #333333; margin: 5px 0;")
        header_layout.addWidget(title_label)
        
        main_layout.addLayout(header_layout)
        
        # 适配器选择组
        adapter_group = QGroupBox("适配器选择")
        adapter_layout = QVBoxLayout(adapter_group)
        
        self.adapter_combo = QComboBox()
        self.adapter_combo.currentTextChanged.connect(self.on_adapter_changed)
        self.adapter_combo.setMinimumHeight(35)
        self.adapter_combo.setStyleSheet("QComboBox { font-size: 12px; padding: 5px; }")
        self.adapter_combo.addItem("正在初始化...")
        self.adapter_combo.setEnabled(False)
        adapter_layout.addWidget(self.adapter_combo)
        
        # 仅显示有线网卡开关（默认开启）
        wired_only_layout = QHBoxLayout()
        self.wired_only_checkbox = QCheckBox("仅显示有线网卡")
        self.wired_only_checkbox.setChecked(True)
        self.wired_only_checkbox.stateChanged.connect(lambda _: self.update_adapter_list(self.current_adapters))
        wired_only_layout.addWidget(self.wired_only_checkbox)
        wired_only_layout.addStretch()
        adapter_layout.addLayout(wired_only_layout)
        
        # 当前状态显示
        self.status_label = QLabel("当前状态: 正在初始化")
        adapter_layout.addWidget(self.status_label)
        
        main_layout.addWidget(adapter_group)
        
        # 设置组
        settings_group = QGroupBox("网络设置")
        settings_layout = QVBoxLayout(settings_group)
        
        # 速度双工设置
        speed_duplex_layout = QHBoxLayout()
        speed_duplex_layout.addWidget(QLabel("速度和双工:"))
        self.speed_duplex_combo = QComboBox()
        self.speed_duplex_combo.setMinimumHeight(35)
        self.speed_duplex_combo.setStyleSheet("QComboBox { font-size: 12px; padding: 5px; }")
        self.speed_duplex_combo.addItem("请等待初始化完成")
        self.speed_duplex_combo.setEnabled(False)
        speed_duplex_layout.addWidget(self.speed_duplex_combo)
        settings_layout.addLayout(speed_duplex_layout)
        
        main_layout.addWidget(settings_group)
        
        # 按钮组
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        
        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.clicked.connect(self.refresh_adapters)
        self.refresh_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; }")
        self.refresh_btn.setMinimumWidth(160)
        self.refresh_btn.setEnabled(False)
        button_layout.addWidget(self.refresh_btn)
        
        self.apply_btn = QPushButton("应用设置")
        self.apply_btn.clicked.connect(self.apply_settings)
        self.apply_btn.setStyleSheet("QPushButton { background-color: #2196F3; color: white; font-weight: bold; }")
        self.apply_btn.setMinimumWidth(160)
        self.apply_btn.setEnabled(False)
        button_layout.addWidget(self.apply_btn)
        
        button_layout.addStretch()
        
        # 右侧按钮
        right_button_layout = QHBoxLayout()
        right_button_layout.setSpacing(5)
        
        self.log_btn = QPushButton("显示日志")
        self.log_btn.clicked.connect(self.toggle_log_display)
        self.log_btn.setMinimumWidth(70)
        self.log_btn.setMaximumWidth(70)
        right_button_layout.addWidget(self.log_btn)
        
        button_layout.addLayout(right_button_layout)
        main_layout.addLayout(button_layout)
        
        # 日志显示区域
        self.log_widget = QTextEdit()
        self.log_widget.setVisible(False)
        self.log_widget.setMaximumHeight(200)
        self.log_widget.setReadOnly(True)
        self.log_widget.setStyleSheet("""
            QTextEdit {
                background-color: #f8f8f8;
                color: #333333;
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 9pt;
                border: 1px solid #cccccc;
            }
        """)
        main_layout.addWidget(self.log_widget)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # 状态栏
        self.statusBar().showMessage("正在启动...")
    
    def toggle_log_display(self):
        """切换日志显示状态"""
        if self.log_visible:
            self.log_widget.setVisible(False)
            self.log_btn.setText("显示日志")
            self.resize(500, 600)
            self.log_visible = False
        else:
            self.log_widget.setVisible(True)
            self.log_btn.setText("隐藏日志")
            all_logs = gui_log_handler.get_all_logs()
            if all_logs:
                self.log_widget.setPlainText(all_logs)
                cursor = self.log_widget.textCursor()
                cursor.movePosition(cursor.End)
                self.log_widget.setTextCursor(cursor)
            self.resize(500, 620)
            self.log_visible = True
    
    def append_log_message(self, message):
        """实时添加日志消息"""
        if self.log_visible:
            self.log_widget.append(message)
            cursor = self.log_widget.textCursor()
            cursor.movePosition(cursor.End)
            self.log_widget.setTextCursor(cursor)
    
    def refresh_adapters(self):
        """刷新适配器列表"""
        if not self.initialization_complete or not self.adapter:
            QMessageBox.warning(self, "警告", "请等待初始化完成")
            return
        
        if self.refresh_thread and self.refresh_thread.isRunning():
            logging.info("停止当前刷新线程")
            self.refresh_thread.quit()
            self.refresh_thread.wait()
        
        self.refresh_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        
        logging.info("启动刷新适配器线程")
        self.refresh_thread = RefreshThread(self.adapter)
        self.refresh_thread.finished.connect(self.on_refresh_finished)
        self.refresh_thread.progress_update.connect(self.on_progress_update)
        self.refresh_thread.start()
    
    def on_refresh_finished(self, success, error_msg, adapters):
        """刷新完成处理"""
        self.refresh_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if success:
            logging.info("刷新适配器成功")
            self.update_adapter_list(adapters)
            # 若存在动态刷新序列，检查是否需要继续
            self._maybe_continue_dynamic_refresh()
        else:
            logging.error(f"刷新适配器失败: {error_msg}")
            self.statusBar().showMessage("刷新失败")
            # 自动重试一次：先尝试重连WMI，然后延时重新刷新
            try:
                if self.adapter:
                    self.adapter.reconnect_wmi()
            except Exception as e:
                logging.warning(f"重连WMI失败: {e}")
            
            def retry_once():
                logging.info("自动重试刷新适配器...")
                self.refresh_adapters()
            QTimer.singleShot(800, retry_once)
            
            # 同时提示用户本次失败，但会自动重试
            QMessageBox.information(self, "正在重试", f"刷新失败，将自动重试一次。\n\n原因: {error_msg}")

    def _maybe_continue_dynamic_refresh(self):
        """在动态刷新序列中，根据当前设置是否已生效决定是否继续下一次刷新。"""
        if not self._dynamic_refresh_active:
            return
        alias = self._dynamic_target_alias
        target = self._dynamic_target_value
        if not alias or not target:
            # 无法判断，终止序列
            self._dynamic_refresh_active = False
            return
        try:
            current = self.settings.get_current_speed_duplex(alias)
        except Exception as e:
            logging.warning(f"检查当前速度双工失败: {e}")
            current = None
        
        # 对比是否已达成目标（去除首尾空格）
        if current and target and current.strip() == target.strip():
            self.statusBar().showMessage("设置已生效")
            self._dynamic_refresh_active = False
            # 刷新确认成功后再弹窗提示
            if self._pending_success_message:
                QMessageBox.information(self, "成功", self._pending_success_message)
                self._pending_success_message = None
            return
        
        # 尚未生效，继续下一次刷新（两档：800ms -> 1500ms）
        if self._dynamic_attempt_idx == 0:
            # 第二次：+1500ms（兜底）
            self._dynamic_attempt_idx = 1
            QTimer.singleShot(1500, self.refresh_adapters)
            self.statusBar().showMessage("最后一次刷新以确认设置...")
        else:
            # 两次之后仍未变化，结束并给出提示
            self._dynamic_refresh_active = False
            self.statusBar().showMessage("设置可能未立即生效，可稍后手动刷新或尝试重启适配器")
            # 不弹出成功提示，避免误导；清理待提示信息
            self._pending_success_message = None
    
    def update_adapter_list(self, adapters):
        """更新适配器列表显示"""
        current_selection = self.adapter_combo.currentText()
        
        self.current_adapters = adapters or []
        self.adapter_combo.clear()
        self.adapter_combo.setEnabled(True)
        
        # 根据开关过滤无线网卡（名称包含 Wireless 或 Wi-Fi 或 WLAN 等常见关键字）
        def is_wireless(name: str) -> bool:
            if not name:
                return False
            low = name.lower()
            return ('wireless' in low) or ('wi-fi' in low) or ('wifi' in low) or ('wlan' in low)
        
        show_wired_only = getattr(self, 'wired_only_checkbox', None) and self.wired_only_checkbox.isChecked()
        filtered = []
        for a in self.current_adapters:
            if show_wired_only and is_wireless(a.get('name', '')):
                continue
            filtered.append(a)
        
        if filtered:
            for adapter in filtered:
                self.adapter_combo.addItem(adapter['name'], userData=adapter.get('alias'))
            
            if current_selection and current_selection not in ["正在初始化...", "请等待初始化完成"]:
                index = self.adapter_combo.findText(current_selection)
                if index >= 0:
                    self.adapter_combo.setCurrentIndex(index)
            
            self.statusBar().showMessage(f"找到 {len(filtered)} 个适配器")
        else:
            self.adapter_combo.addItem("未找到可用的网络适配器")
            self.statusBar().showMessage("未找到可用的网络适配器")
    
    def on_adapter_changed(self):
        """适配器选择改变处理"""
        if not self.initialization_complete:
            return
            
        current_text = self.adapter_combo.currentText()
        current_alias = self.adapter_combo.currentData()
        
        if not current_text or not self.current_adapters:
            return
        
        for adapter in self.current_adapters:
            if adapter['name'] == current_text:
                alias = current_alias or adapter.get('alias') or adapter['name']
                actual_speed_duplex = self.settings.get_current_speed_duplex(alias)
                status_text = (f"当前状态: {actual_speed_duplex} | "
                             f"IP: {adapter['ip_address']}")
                self.status_label.setText(status_text)
                self.update_speed_duplex_options(alias)
                break
    
    def update_speed_duplex_options(self, adapter_alias: str):
        """更新速度双工选项"""
        if not adapter_alias or not adapter_alias.strip():
            self.speed_duplex_combo.clear()
            self.speed_duplex_combo.addItem("请先选择适配器")
            return
        
        try:
            current_selection = self.speed_duplex_combo.currentText()
            options = self.adapter.get_speed_duplex_options(adapter_alias, use_fallback=True)
            
            self.speed_duplex_combo.clear()
            self.speed_duplex_combo.setEnabled(True)
            
            if options:
                self.speed_duplex_combo.addItems(options)
                if current_selection in options:
                    self.speed_duplex_combo.setCurrentText(current_selection)
                else:
                    self.speed_duplex_combo.setCurrentIndex(0)
            else:
                self.speed_duplex_combo.addItem("无可用选项")
                
        except Exception as e:
            logging.warning(f"更新速度双工选项失败: {str(e)}")
            self.speed_duplex_combo.clear()
            self.speed_duplex_combo.addItem("获取选项失败")
    
    def apply_settings(self):
        """应用网络设置"""
        if not self.initialization_complete:
            QMessageBox.warning(self, "警告", "请等待初始化完成")
            return
            
        if not self.adapter_combo.currentText():
            QMessageBox.warning(self, "警告", "请先选择一个网络适配器")
            return
        
        if not self.settings.is_admin:
            QMessageBox.warning(self, "权限不足", "需要管理员权限才能修改网络设置")
            return
        
        adapter_name = self.adapter_combo.currentText()
        adapter_alias = self.adapter_combo.currentData() or adapter_name
        speed_duplex = self.speed_duplex_combo.currentText()
        
        reply = QMessageBox.question(self, "确认操作", 
                                   f"确定要修改适配器 '{adapter_name}' 的设置吗?\n\n"
                                   f"速度和双工模式: {speed_duplex}",
                                   QMessageBox.Yes | QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            self.start_operation(adapter_alias, speed_duplex)
    
    def start_operation(self, adapter_name, speed_duplex):
        """启动后台操作"""
        self.apply_btn.setEnabled(False)
        self.refresh_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        self.statusBar().showMessage("正在应用设置...")
        
        self.worker_thread = WorkerThread(self.settings, self.adapter, adapter_name, speed_duplex)
        self.worker_thread.finished.connect(self.on_operation_finished)
        self.worker_thread.progress_update.connect(self.on_progress_update)
        self.worker_thread.start()
    
    def on_operation_finished(self, success, message, status_data):
        """操作完成处理"""
        self.apply_btn.setEnabled(True)
        self.refresh_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if success:
            self.statusBar().showMessage("设置应用成功，正在刷新状态...")
            # 延后弹窗：待刷新确认后再提示成功
            self._pending_success_message = message
            # 从 status_data 或按钮当前选择推断目标
            selected_alias = self.adapter_combo.currentData() or self.adapter_combo.currentText()
            target_value = self.speed_duplex_combo.currentText()
            try:
                if status_data and isinstance(status_data, list) and status_data:
                    maybe = status_data[0]
                    if isinstance(maybe, dict) and maybe.get('adapter_name'):
                        selected_alias = maybe['adapter_name']
                    if isinstance(maybe, dict) and maybe.get('new_status') and maybe['new_status'] != 'Unknown':
                        target_value = maybe['new_status']
            except Exception:
                pass
            
            # 设置动态刷新状态
            self._dynamic_refresh_active = True
            self._dynamic_target_alias = selected_alias
            self._dynamic_target_value = target_value
            self._dynamic_attempt_idx = 0
            
            # 第一次刷新：800ms
            QTimer.singleShot(800, self.refresh_adapters)
        else:
            self.statusBar().showMessage("设置应用失败")
            QMessageBox.critical(self, "失败", message)
    
    def on_progress_update(self, message):
        """更新进度信息"""
        logging.info(f"进度更新: {message}")
        self.statusBar().showMessage(message)
    
    def show_system_diagnosis(self):
        """显示系统诊断对话框"""
        try:
            compatibility = SystemCompatibility()
            report = compatibility.get_compatibility_report()
            
            # 构建诊断报告文本
            diag_text = "系统兼容性诊断报告\n"
            diag_text += "=" * 40 + "\n\n"
            
            # 系统信息
            diag_text += "系统信息:\n"
            sys_info = report['system_info']
            diag_text += f"  平台: {sys_info.get('platform', 'Unknown')}\n"
            diag_text += f"  Python版本: {sys_info.get('python_version', 'Unknown').split()[0]}\n"
            diag_text += f"  管理员权限: {'是' if sys_info.get('is_admin', False) else '否'}\n\n"
            
            # PowerShell信息
            diag_text += "PowerShell兼容性:\n"
            ps_info = report['powershell']
            diag_text += f"  可用性: {'是' if ps_info['available'] else '否'}\n"
            if ps_info['available']:
                diag_text += f"  路径: {ps_info['path']}\n"
                diag_text += f"  版本: {ps_info['version']}\n"
                diag_text += f"  执行策略: {ps_info['execution_policy']}\n"
            diag_text += "\n"
            
            # WMI信息
            diag_text += "WMI兼容性:\n"
            wmi_info = report['wmi']
            diag_text += f"  可用性: {'是' if wmi_info['available'] else '否'}\n"
            diag_text += f"  服务运行: {'是' if wmi_info['service_running'] else '否'}\n"
            if wmi_info['error']:
                diag_text += f"  错误: {wmi_info['error']}\n"
            diag_text += "\n"
            
            # 网络命令兼容性
            diag_text += "网络命令兼容性:\n"
            net_info = report['network_commands']
            diag_text += f"  netsh: {'可用' if net_info['netsh_available'] else '不可用'}\n"
            diag_text += f"  Get-NetAdapter: {'可用' if net_info['get_netadapter_available'] else '不可用'}\n"
            diag_text += f"  wmic: {'可用' if net_info['wmic_available'] else '不可用'}\n\n"
            
            # 建议
            if report['recommendations']:
                diag_text += "建议:\n"
                for i, rec in enumerate(report['recommendations'], 1):
                    diag_text += f"  {i}. {rec}\n"
            
            # 创建对话框
            msg_box = QMessageBox(self)
            msg_box.setWindowTitle("系统诊断")
            msg_box.setText("系统兼容性诊断完成")
            msg_box.setDetailedText(diag_text)
            msg_box.setIcon(QMessageBox.Information)
            
            # 设置对话框大小，让详细信息区域更大
            msg_box.resize(800, 600)
            
            # 查找详细文本区域并设置最小大小
            for widget in msg_box.findChildren(QTextEdit):
                widget.setMinimumSize(750, 400)
                widget.setMaximumSize(750, 400)
            
            msg_box.exec_()
            
        except Exception as e:
            QMessageBox.critical(self, "错误", f"系统诊断失败: {str(e)}")
    
    def show_about(self):
        """显示关于对话框"""
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton, QScrollArea
        from PyQt5.QtCore import QUrl
        from PyQt5.QtGui import QDesktopServices
        
        # 创建自定义对话框
        dialog = QDialog(self)
        dialog.setWindowTitle("关于")
        dialog.setFixedSize(500, 600)
        
        layout = QVBoxLayout(dialog)
        
        # 创建滚动区域
        scroll_area = QScrollArea()
        scroll_widget = QWidget()
        scroll_layout = QVBoxLayout(scroll_widget)
        
        # 标题
        title_label = QLabel("网络适配器管理工具 v1.1")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold; margin: 10px 0;")
        scroll_layout.addWidget(title_label)
        
        # 描述
        desc_label = QLabel("Windows系统网络适配器速度和双工模式管理工具，支持图形化界面操作。\n为NA（广软网协）而做")
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("margin: 10px 0;")
        scroll_layout.addWidget(desc_label)
        
        # 功能特性
        features_text = """🔧 核心功能:
• 网络适配器管理 - 查看和修改网络适配器速度和双工模式
• 实时状态显示 - 显示IP地址、连接状态和网络速度
• 智能过滤 - 支持仅显示有线网卡，过滤无线适配器
• 多线程处理 - 后台操作，界面响应流畅

🌐 系统兼容性:
• Windows版本支持 - Windows 7/8/10/11 (32位/64位)
• PowerShell兼容 - 支持PowerShell 5.x 和 PowerShell 7.x
• 多路径检测 - 自动检测系统中可用的PowerShell版本
• WMI兼容性 - 智能WMI连接管理，支持多线程环境
• 权限管理 - 多种管理员权限检测方法，自动提权

🛡️ 健壮性设计:
• 系统诊断 - 内置兼容性检查和诊断工具
• 错误恢复 - 智能错误处理和降级方案
• 资源管理 - 支持多种部署方式（源码/打包exe）
• 日志系统 - 详细的操作日志和错误追踪

开发信息:
• 基于Python和PyQt5开发
• 使用WMI和PowerShell进行系统管理
• 开源项目，欢迎贡献代码

许可证: MIT License"""
        
        features_label = QLabel(features_text)
        features_label.setWordWrap(True)
        features_label.setStyleSheet("margin: 10px 0;")
        scroll_layout.addWidget(features_label)
        
        # GitHub链接按钮
        github_btn = QPushButton("🔗 项目地址: https://github.com/CurtisYan/NetAdapterTool")
        github_btn.setStyleSheet("""
            QPushButton {
                background-color: #0366d6;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                text-align: left;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #0256cc;
            }
            QPushButton:pressed {
                background-color: #024ea4;
            }
        """)
        github_btn.clicked.connect(lambda: QDesktopServices.openUrl(QUrl("https://github.com/CurtisYan/NetAdapterTool")))
        scroll_layout.addWidget(github_btn)
        
        scroll_widget.setLayout(scroll_layout)
        scroll_area.setWidget(scroll_widget)
        scroll_area.setWidgetResizable(True)
        layout.addWidget(scroll_area)
        
        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn)
        
        dialog.exec_()
    
    def closeEvent(self, event):
        """关闭程序处理"""
        logging.info("程序关闭中...")
        
        # 停止所有线程
        threads = [self.worker_thread, self.refresh_thread, self.init_thread]
        for thread in threads:
            if thread and thread.isRunning():
                thread.quit()
                thread.wait(3000)  # 最多等待3秒
        
        logging.info("程序已关闭")
        event.accept()


def main():
    try:
        # 隐藏控制台窗口
        try:
            import ctypes
            import ctypes.wintypes
            
            # 获取控制台窗口句柄
            kernel32 = ctypes.windll.kernel32
            user32 = ctypes.windll.user32
            
            # 获取当前进程的控制台窗口
            console_window = kernel32.GetConsoleWindow()
            if console_window:
                # 隐藏控制台窗口 (SW_HIDE = 0)
                user32.ShowWindow(console_window, 0)
        except Exception as e:
            # 如果隐藏失败，继续运行程序
            pass
        
        print("Starting Network Adapter Tool...")
        print(f"Python version: {sys.version}")
        print(f"Working directory: {os.getcwd()}")
        
        app = QApplication(sys.argv)
        
        # 设置应用程序信息
        app.setApplicationName("网络适配器管理工具")
        app.setApplicationVersion("1.1")
        print("QApplication created successfully")
        
        # 设置应用程序图标
        try:
            # 支持多种部署方式的图标路径查找
            possible_icon_paths = [
                os.path.join(os.path.dirname(__file__), "img", "NA.ico"),  # 源码运行
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "img", "NA.ico"),  # 绝对路径
                os.path.join(os.getcwd(), "img", "NA.ico"),  # 当前工作目录
                "img/NA.ico",  # 相对路径
                "NA.ico"  # 同目录
            ]
            
            icon_loaded = False
            for icon_path in possible_icon_paths:
                print(f"Looking for icon at: {icon_path}")
                if os.path.exists(icon_path):
                    app.setWindowIcon(QIcon(icon_path))
                    print("Icon loaded successfully")
                    icon_loaded = True
                    break
            
            if not icon_loaded:
                print("Icon file not found in any location")
        except Exception as e:
            print(f"Icon loading error: {e}")
        
        # 创建主窗口
        print("Creating main window...")
        window = NetworkAdapterGUI()
        print("Main window created successfully")
        
        window.show()
        print("Main window shown, starting event loop...")
        
        sys.exit(app.exec_())
        
    except Exception as e:
        print(f"CRITICAL ERROR in main(): {e}")
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")  # 保持控制台窗口打开


if __name__ == "__main__":
    main()
