"""
网络适配器管理工具图形界面
使用PyQt5实现简洁的GUI界面
"""

import sys
import os
import ctypes
import logging
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QComboBox, QPushButton, 
                             QGroupBox, QMessageBox, QProgressBar, QDialog,
                             QTextBrowser, QScrollArea, QTextEdit)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl
from PyQt5.QtGui import QFont, QPixmap, QDesktopServices, QIcon
from PyQt5.QtSvg import QSvgWidget
from network_adapter import NetworkAdapter
from network_settings import NetworkSettings

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
        logging.StreamHandler(sys.stdout),  # 输出到控制台
        gui_log_handler  # 输出到GUI
    ]
)


class AboutDialog(QDialog):
    """关于对话框"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("关于 - 网络适配器管理工具")
        self.setFixedSize(450, 600)
        self.setWindowFlags(Qt.Dialog | Qt.WindowCloseButtonHint)
        self.init_ui()
    
    def init_ui(self):
        """初始化关于对话框界面"""
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 主内容布局（不使用滚动区域）
        content_layout = QVBoxLayout()
        content_layout.setSpacing(8)
        
        # 应用图标和标题
        title_layout = QHBoxLayout()
        
        # 尝试加载应用图标
        icon_label = QLabel()
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "img", "NA (蓝透明).jpg")
            if os.path.exists(icon_path):
                pixmap = QPixmap(icon_path)
                if not pixmap.isNull():
                    # 调整图标大小
                    scaled_pixmap = pixmap.scaled(48, 48, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    icon_label.setPixmap(scaled_pixmap)
                else:
                    raise ValueError("图片加载失败")
            else:
                raise FileNotFoundError("图标文件不存在")
        except (FileNotFoundError, ValueError, OSError):
            # 图标加载失败时使用默认图标
            icon_label.setText("💻")
            icon_label.setFont(QFont("", 32))
            icon_label.setAlignment(Qt.AlignCenter)
        title_layout.addWidget(icon_label)
        
        # 标题信息
        title_info_layout = QVBoxLayout()
        app_title = QLabel("网络适配器管理工具")
        app_title.setFont(QFont("Microsoft YaHei", 16, QFont.Bold))
        app_title.setAlignment(Qt.AlignLeft)
        title_info_layout.addWidget(app_title)
        
        version_label = QLabel("版本 1.0")
        version_label.setFont(QFont("Microsoft YaHei", 9))
        version_label.setStyleSheet("color: #666666;")
        title_info_layout.addWidget(version_label)
        
        title_layout.addLayout(title_info_layout)
        title_layout.addStretch()
        content_layout.addLayout(title_layout)
        
        # 分隔线
        separator1 = QLabel()
        separator1.setStyleSheet("border-bottom: 1px solid #E0E0E0; margin: 5px 0;")
        content_layout.addWidget(separator1)
        
        # 归属信息 - 突出显示
        ownership_layout = QVBoxLayout()
        ownership_layout.setSpacing(3)
        ownership_text = QLabel("该应用归")
        ownership_text.setFont(QFont("Microsoft YaHei", 11))
        ownership_text.setAlignment(Qt.AlignCenter)
        ownership_layout.addWidget(ownership_text)
        
        # 广软网络管理工作站 - 大字体可点击
        workstation_label = QLabel('<a href="https://service.seig.edu.cn/join" style="text-decoration: none; color: #1976D2;">广软网络管理工作站</a>')
        workstation_label.setFont(QFont("Microsoft YaHei", 18, QFont.Bold))
        workstation_label.setAlignment(Qt.AlignCenter)
        workstation_label.setOpenExternalLinks(True)
        workstation_label.setStyleSheet("""
            QLabel {
                padding: 8px;
                border: 2px solid #1976D2;
                border-radius: 6px;
                background-color: #E3F2FD;
                margin: 5px 0;
            }
            QLabel:hover {
                background-color: #BBDEFB;
            }
        """)
        ownership_layout.addWidget(workstation_label)
        
        ownership_text2 = QLabel("所有")
        ownership_text2.setFont(QFont("Microsoft YaHei", 11))
        ownership_text2.setAlignment(Qt.AlignCenter)
        ownership_layout.addWidget(ownership_text2)
        
        content_layout.addLayout(ownership_layout)
        
        # 欢迎信息
        welcome_label = QLabel("欢迎使用和分享")
        welcome_label.setFont(QFont("Microsoft YaHei", 11))
        welcome_label.setAlignment(Qt.AlignCenter)
        welcome_label.setStyleSheet("color: #4CAF50; font-weight: bold; margin: 3px 0;")
        content_layout.addWidget(welcome_label)
        
        # 分隔线
        separator2 = QLabel()
        separator2.setStyleSheet("border-bottom: 1px solid #E0E0E0; margin: 5px 0;")
        content_layout.addWidget(separator2)
        
        # GitHub 链接
        github_layout = QHBoxLayout()
        github_layout.setAlignment(Qt.AlignCenter)
        
        # GitHub 图标
        try:
            github_svg_path = os.path.join(os.path.dirname(__file__), "img", "github-fill.svg")
            if os.path.exists(github_svg_path):
                github_icon = QSvgWidget(github_svg_path)
                github_icon.setFixedSize(18, 18)
            else:
                raise FileNotFoundError("GitHub图标文件不存在")
        except (FileNotFoundError, OSError):
            github_icon = QLabel("⭐")
            github_icon.setFont(QFont("", 16))
            github_icon.setStyleSheet("color: #333333;")
        github_layout.addWidget(github_icon)
        
        github_text = QLabel('<a href="https://github.com/CurtisYan/NetAdapterTool" style="text-decoration: none; color: #333333;">GitHub 项目地址</a>')
        github_text.setFont(QFont("Microsoft YaHei", 11))
        github_text.setOpenExternalLinks(True)
        github_layout.addWidget(github_text)
        
        content_layout.addLayout(github_layout)
        
        # 贡献信息
        contribute_label = QLabel("欢迎 Fork、Pull Request 和 Issue")
        contribute_label.setFont(QFont("Microsoft YaHei", 11))
        contribute_label.setAlignment(Qt.AlignCenter)
        contribute_label.setStyleSheet("color: #666666; margin-top: 5px;")
        content_layout.addWidget(contribute_label)
        
        # 技术信息
        tech_info = QLabel("基于 Python + PyQt5 开发\n使用 WMI 和 PowerShell 进行网络管理\n支持 Windows 10/11 系统")
        tech_info.setFont(QFont("Microsoft YaHei", 9))
        tech_info.setAlignment(Qt.AlignCenter)
        tech_info.setStyleSheet("color: #888888; margin-top: 8px;")
        content_layout.addWidget(tech_info)
        
        # 弹性空间
        content_layout.addStretch()
        
        # 作者信息（移到最下面，降低醒目度）
        author_label = QLabel("作者：Curtis Yan")
        author_label.setFont(QFont("Microsoft YaHei", 9))
        author_label.setAlignment(Qt.AlignCenter)
        author_label.setStyleSheet("color: #999999; margin-top: 10px;")
        content_layout.addWidget(author_label)
        
        # 添加内容布局到主布局
        layout.addLayout(content_layout)
        
        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                border: none;
                padding: 8px 20px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
        """)
        close_btn.setFixedWidth(100)
        
        close_layout = QHBoxLayout()
        close_layout.addStretch()
        close_layout.addWidget(close_btn)
        close_layout.addStretch()
        
        layout.addLayout(close_layout)


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
            logging.info(f"开始应用网络设置: {self.adapter_name} -> {self.speed_duplex}")
            # 第一步：应用设置
            self.progress_update.emit("正在应用网络设置...")
            success, message = self.settings.set_adapter_speed_duplex(
                self.adapter_name, self.speed_duplex)
            
            if success:
                logging.info("网络设置应用成功，等待网络适配器重新初始化")
                # 等待网络适配器重新初始化（网络设置更改后需要时间）
                self.progress_update.emit("等待网络适配器重新初始化...")
                import time
                time.sleep(3)  # 等待3秒让适配器完全重新初始化
                
                # 设置应用成功，获取更新后的状态
                logging.info("网络设置应用成功，获取更新状态")
                try:
                    # 获取当前适配器的最新状态
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


class RefreshThread(QThread):
    """刷新适配器的后台线程"""
    finished = pyqtSignal(bool, str, list)  # 成功标志，错误信息，适配器列表
    progress_update = pyqtSignal(str)  # 进度更新信号
    
    def __init__(self, adapter):
        super().__init__()
        self.adapter = adapter
    
    def run(self):
        try:
            logging.info("开始刷新适配器列表")
            self.progress_update.emit("正在刷新适配器列表...")
            adapters = self.adapter.get_all_adapters()
            logging.info(f"刷新完成，找到 {len(adapters)} 个适配器")
            self.finished.emit(True, "", adapters)
        except Exception as e:
            error_msg = f"刷新适配器失败: {str(e)}"
            logging.error(error_msg)
            self.finished.emit(False, error_msg, [])


class NetworkAdapterGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.adapter = NetworkAdapter()
        self.settings = NetworkSettings()
        self.current_adapters = []
        self.worker_thread = None
        self.refresh_thread = None
        self.log_visible = False  # 日志区域是否可见
        
        self.init_ui()
        
        # 设置日志回调
        gui_log_handler.set_gui_callback(self.append_log_message)
        
        logging.info(f"程序启动 - 管理员模式: {self.settings.is_admin}")
        logging.info("网络适配器管理工具已启动")
        
        # 如果不是管理员，自动以管理员身份重启
        if not self.settings.is_admin:
            self.auto_restart_as_admin()
        else:
            self.refresh_adapters()
    
    def init_ui(self):
        """初始化用户界面"""
        self.setWindowTitle("网络适配器管理工具")
        self.setMinimumSize(500, 600)
        self.setMaximumWidth(500)
        self.resize(500, 600)
        
        # 设置窗口和任务栏图标
        try:
            icon_path = os.path.join(os.path.dirname(__file__), "img", "NA.ico")
            if os.path.exists(icon_path):
                from PyQt5.QtGui import QIcon
                self.setWindowIcon(QIcon(icon_path))
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
        
        # NA（蓝透明）图片
        logo_label = QLabel()
        logo_label.setAlignment(Qt.AlignCenter)
        try:
            logo_path = os.path.join(os.path.dirname(__file__), "img", "NA (蓝透明).jpg")
            if os.path.exists(logo_path):
                pixmap = QPixmap(logo_path)
                if not pixmap.isNull():
                    # 调整图片大小，让它更突出
                    scaled_pixmap = pixmap.scaled(120, 120, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                    logo_label.setPixmap(scaled_pixmap)
                    logo_label.setStyleSheet("margin: 15px 0;")
                else:
                    raise ValueError("图片加载失败")
            else:
                raise FileNotFoundError("Logo文件不存在")
        except (FileNotFoundError, ValueError, OSError):
            # Logo加载失败时使用文字标识
            logo_label.setText("")
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
        # 初始化时不添加任何选项，等待适配器选择后再更新
        self.speed_duplex_combo.setMinimumHeight(35)  # 增大高度
        self.speed_duplex_combo.setStyleSheet("QComboBox { font-size: 12px; padding: 5px; } QComboBox::drop-down { width: 25px; } QComboBox QAbstractItemView { min-width: 250px; }")
        speed_duplex_layout.addWidget(self.speed_duplex_combo)
        settings_layout.addLayout(speed_duplex_layout)
        
        main_layout.addWidget(settings_group)
        
        # 按钮组
        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)  # 设置统一间距
        
        self.refresh_btn = QPushButton("刷新")
        self.refresh_btn.clicked.connect(self.refresh_adapters)
        self.refresh_btn.setStyleSheet("QPushButton { background-color: #4CAF50; color: white; font-weight: bold; }")
        self.refresh_btn.setMinimumWidth(160)
        button_layout.addWidget(self.refresh_btn)
        
        self.apply_btn = QPushButton("应用设置")
        self.apply_btn.clicked.connect(self.apply_settings)
        self.apply_btn.setStyleSheet("QPushButton { background-color: #2196F3; color: white; font-weight: bold; }")
        self.apply_btn.setMinimumWidth(160)
        button_layout.addWidget(self.apply_btn)
        
        # 添加弹性空间，将右侧按钮推到右边
        button_layout.addStretch()
        
        # 创建右侧按钮子布局，控制间距
        right_button_layout = QHBoxLayout()
        right_button_layout.setSpacing(5)  # 设置较小间距
        
        self.about_btn = QPushButton("关于")
        self.about_btn.clicked.connect(self.show_about)
        self.about_btn.setMinimumWidth(70)
        self.about_btn.setMaximumWidth(70)
        right_button_layout.addWidget(self.about_btn)
        
        self.log_btn = QPushButton("显示日志")
        self.log_btn.clicked.connect(self.toggle_log_display)
        self.log_btn.setMinimumWidth(70)
        self.log_btn.setMaximumWidth(70)
        right_button_layout.addWidget(self.log_btn)
        
        # 将右侧按钮布局添加到主布局
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
        if self.settings.is_admin:
            self.statusBar().showMessage("就绪 - 管理员模式")
        else:
            self.statusBar().showMessage("就绪 - 普通用户模式（修改设置需要提权）")
    
    def toggle_log_display(self):
        """切换日志显示状态"""
        if self.log_visible:
            # 隐藏日志
            self.log_widget.setVisible(False)
            self.log_btn.setText("显示日志")
            self.resize(500, 600)
            self.log_visible = False
        else:
            # 显示日志
            self.log_widget.setVisible(True)
            self.log_btn.setText("隐藏日志")
            # 加载所有历史日志
            all_logs = gui_log_handler.get_all_logs()
            if all_logs:
                self.log_widget.setPlainText(all_logs)
                # 滚动到底部
                cursor = self.log_widget.textCursor()
                cursor.movePosition(cursor.End)
                self.log_widget.setTextCursor(cursor)
            self.resize(500, 620)
            self.log_visible = True
    
    def append_log_message(self, message):
        """实时添加日志消息"""
        if self.log_visible:
            self.log_widget.append(message)
            # 自动滚动到底部
            cursor = self.log_widget.textCursor()
            cursor.movePosition(cursor.End)
            self.log_widget.setTextCursor(cursor)
    
    def refresh_adapters(self):
        """刷新适配器列表（多线程）"""
        # 如果已有刷新线程在运行，先停止它
        if self.refresh_thread and self.refresh_thread.isRunning():
            logging.info("停止当前刷新线程")
            self.refresh_thread.quit()
            self.refresh_thread.wait()
        
        # 禁用刷新按钮，显示进度
        self.refresh_btn.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # 无限进度条
        
        # 启动刷新线程
        logging.info("启动刷新适配器线程")
        self.refresh_thread = RefreshThread(self.adapter)
        self.refresh_thread.finished.connect(self.on_refresh_finished)
        self.refresh_thread.progress_update.connect(self.on_progress_update)
        self.refresh_thread.start()
    
    def update_adapter_list(self, adapters):
        """更新适配器列表显示"""
        # 保存当前选中的适配器
        current_selection = self.adapter_combo.currentText()
        
        self.current_adapters = adapters
        self.adapter_combo.clear()
        self.adapter_combo.setEnabled(True)  # 确保下拉框可用
        
        if self.current_adapters:
            for adapter in self.current_adapters:
                # 显示设备名，内部保存接口别名（InterfaceAlias）
                self.adapter_combo.addItem(adapter['name'], userData=adapter.get('alias'))
            
            # 尝试恢复之前的选择
            if current_selection and current_selection not in ["请点击刷新按钮重试", "请先刷新适配器"]:
                index = self.adapter_combo.findText(current_selection)
                if index >= 0:
                    self.adapter_combo.setCurrentIndex(index)
            
            self.statusBar().showMessage(f"找到 {len(self.current_adapters)} 个适配器")
        else:
            self.adapter_combo.addItem("未找到可用的网络适配器")
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
    
    def show_admin_prompt(self):
        """显示管理员权限提示"""
        reply = QMessageBox.question(self, "需要管理员权限", 
                                   "此程序需要管理员权限才能正常工作\n\n"
                                   "是否要以管理员身份重新启动程序？",
                                   QMessageBox.Yes | QMessageBox.No)
        if reply == QMessageBox.Yes:
            # 立即隐藏窗口，避免用户看到两个窗口
            self.hide()
            self.restart_as_admin()
        else:
            # 用户选择不以管理员身份运行，关闭程序
            self.close()
    
    def apply_settings(self):
        """应用网络设置"""
        if not self.adapter_combo.currentText():
            QMessageBox.warning(self, "警告", "请先选择一个网络适配器")
            return
        
        if not self.settings.is_admin:
            self.show_admin_prompt()
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
        if not adapter_alias or not adapter_alias.strip():
            self.speed_duplex_combo.clear()
            self.speed_duplex_combo.addItem("请先选择适配器")
            return
            
        try:
            # 获取当前选中的值
            current_selection = self.speed_duplex_combo.currentText()
            
            # 获取适配器支持的选项
            options = self.adapter.get_speed_duplex_options(adapter_alias)
            
            # 更新下拉菜单
            self.speed_duplex_combo.clear()
            self.speed_duplex_combo.setEnabled(True)  # 确保下拉框可用
            
            if options:
                self.speed_duplex_combo.addItems(options)
                # 尝试恢复之前的选择
                if current_selection in options:
                    self.speed_duplex_combo.setCurrentText(current_selection)
                elif options:
                    self.speed_duplex_combo.setCurrentIndex(0)
            else:
                # 如果没有获取到选项，添加一个提示但保持可用
                self.speed_duplex_combo.addItem("无可用选项 - 请检查适配器")
                
        except Exception as e:
            logging.warning(f"更新速度双工选项失败: {str(e)}")
            self.speed_duplex_combo.clear()
            self.speed_duplex_combo.setEnabled(True)  # 保持可用
            self.speed_duplex_combo.addItem("获取选项失败 - 请重试")
    
    def on_refresh_finished(self, success, error_msg, adapters):
        """刷新完成的处理"""
        # 恢复界面状态
        self.refresh_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if success:
            logging.info("刷新适配器成功")
            self.update_adapter_list(adapters)
        else:
            logging.error(f"刷新适配器失败: {error_msg}")
            self.statusBar().showMessage("刷新失败 - 点击刷新按钮重试")
            
            # 显示错误对话框，并提供重试选项
            reply = QMessageBox.question(self, "刷新失败", 
                                       f"{error_msg}\n\n是否要重试？",
                                       QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                # 重试刷新
                self.refresh_adapters()
            else:
                # 用户选择不重试，保持界面可用状态
                self.adapter_combo.clear()
                self.adapter_combo.addItem("请点击刷新按钮重试")
                self.speed_duplex_combo.clear()
                self.speed_duplex_combo.addItem("请先刷新适配器")
                self.status_label.setText("当前状态: 刷新失败")
    
    def on_progress_update(self, message):
        """更新进度信息"""
        logging.info(f"进度更新: {message}")
        self.statusBar().showMessage(message)
    
    def on_operation_finished(self, success, message, status_data):
        """操作完成的处理"""
        # 恢复界面状态
        self.apply_btn.setEnabled(True)
        self.refresh_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        if success:
            # 更新当前适配器的状态显示
            if status_data and len(status_data) > 0:
                adapter_name = status_data[0].get('adapter_name')
                new_status = status_data[0].get('new_status')
                if adapter_name and new_status:
                    # 更新状态显示
                    current_text = self.adapter_combo.currentText()
                    if current_text == adapter_name or self.adapter_combo.currentData() == adapter_name:
                        # 获取当前适配器的IP地址
                        current_ip = "未知"
                        for adapter in self.current_adapters:
                            if adapter['name'] == current_text:
                                current_ip = adapter['ip_address']
                                break
                        
                        status_text = f"当前状态: {new_status} | IP: {current_ip}"
                        self.status_label.setText(status_text)
                        logging.info(f"状态已更新: {new_status}")
            
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
            
            # 立即关闭当前程序，避免两个窗口同时存在
            self.close()
            QApplication.processEvents()  # 处理关闭事件
            
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
            
            # 退出当前程序
            QApplication.quit()
            sys.exit(0)
            
        except Exception as e:
            # 如果启动失败，显示窗口并提示错误
            self.show()
            QMessageBox.critical(self, "错误", f"无法以管理员身份启动程序: {str(e)}")
    
    def auto_restart_as_admin(self):
        """自动以管理员身份重启程序"""
        try:
            # 获取当前程序路径
            if getattr(sys, 'frozen', False):
                # 如果是打包后的exe文件
                current_exe = sys.executable
            else:
                # 如果是Python脚本
                current_exe = sys.executable
                script_path = os.path.abspath(__file__)
            
            # 立即关闭当前程序，避免两个窗口同时存在
            self.close()
            QApplication.processEvents()  # 处理关闭事件
            
            # 使用ShellExecuteW以管理员身份启动
            if getattr(sys, 'frozen', False):
                # 打包后的exe - 使用SW_HIDE隐藏窗口
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", current_exe, None, None, 0
                )
            else:
                # Python脚本 - 使用pythonw.exe隐藏控制台窗口
                pythonw_exe = sys.executable.replace('python.exe', 'pythonw.exe')
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", pythonw_exe, f'"{script_path}"', None, 0
                )
            
            # 退出当前程序
            QApplication.quit()
            sys.exit(0)
            
        except Exception as e:
            # 如果启动失败，显示窗口并提示错误
            self.show()
            QMessageBox.critical(self, "错误", f"无法以管理员身份启动程序: {str(e)}")
    
    def show_about(self):
        """显示关于对话框"""
        about_dialog = AboutDialog(self)
        about_dialog.exec_()
    
    def closeEvent(self, event):
        """关闭程序时的处理"""
        logging.info("程序关闭中...")
        
        # 停止工作线程
        if self.worker_thread and self.worker_thread.isRunning():
            logging.info("停止工作线程")
            self.worker_thread.quit()
            self.worker_thread.wait()
        
        # 停止刷新线程
        if self.refresh_thread and self.refresh_thread.isRunning():
            logging.info("停止刷新线程")
            self.refresh_thread.quit()
            self.refresh_thread.wait()
        
        logging.info("程序已关闭")
        event.accept()


def main():
    app = QApplication(sys.argv)
    
    # 设置应用程序信息
    app.setApplicationName("网络适配器管理工具")
    app.setApplicationVersion("1.0")
    
    # 设置应用程序图标
    try:
        icon_path = os.path.join(os.path.dirname(__file__), "img", "NA.ico")
        if os.path.exists(icon_path):
            app.setWindowIcon(QIcon(icon_path))
    except Exception:
        pass  # 静默失败，不影响程序启动
    
    # 创建主窗口
    window = NetworkAdapterGUI()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
