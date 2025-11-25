import sys
import os
import io
import threading
import datetime
import traceback

from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QLabel, QFrame, 
                             QTextEdit, QRadioButton, QComboBox, QButtonGroup,
                             QGraphicsDropShadowEffect, QSizePolicy, QScrollArea, QLineEdit)
from PyQt6.QtCore import Qt, QSize, pyqtSignal, QObject
from PyQt6.QtGui import QColor, QFont, QIcon, QCursor, QPixmap

# Import the logic from existing scripts
import format_docx
import extract_sketch_maps
import insert_collected_images
import merge_by_volumes

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # 开发环境：使用脚本所在目录
        base_path = os.path.dirname(os.path.abspath(__file__))

    return os.path.join(base_path, relative_path)

# Redirect stdout to capture print statements
class Stream(io.StringIO):
    def __init__(self, log_func):
        super().__init__()
        self.log_func = log_func

    def write(self, text):
        if text.strip():
            self.log_func(text)
    
    def flush(self):
        pass

class AppLogic:
    def __init__(self, log_func, finish_func):
        self.log_func = log_func
        self.finish_func = finish_func

    def run_task(self, task_type, merge_option=None, merge_value=None):
        # Redirect stdout
        original_stdout = sys.stdout
        sys.stdout = Stream(self.log_func)
        
        try:
            if task_type == "format":
                print(">>> 开始执行：格式化文档...")
                format_docx.run_batch()
            elif task_type == "extract":
                print(">>> 开始执行：提取素描图...")
                extract_sketch_maps.run_batch()
            elif task_type == "insert":
                print(">>> 开始执行：插入素描图...")
                insert_collected_images.run_batch()
            elif task_type == "merge":
                print(">>> 开始执行：分册合并...")
                if merge_option == "routes_per_volume":
                    print(f"分册模式：每册 {merge_value} 条路线")
                    merge_by_volumes.run_batch_with_routes_per_volume(merge_value)
                elif merge_option == "total_volumes":
                    print(f"分册模式：总共分为 {merge_value} 册")
                    merge_by_volumes.run_batch_with_total_volumes(merge_value)
                else:
                    merge_by_volumes.run_batch()
            elif task_type == "all":
                print(">>> 🚀 开始一键全自动运行...")
                
                print("\n--- 步骤 1/4: 格式化文档 ---")
                format_docx.run_batch()
                
                print("\n--- 步骤 2/4: 提取素描图 ---")
                extract_sketch_maps.run_batch()
                
                print("\n--- 步骤 3/4: 插入素描图 ---")
                insert_collected_images.run_batch()
                
                print("\n--- 步骤 4/4: 分册合并 ---")
                if merge_option == "routes_per_volume":
                    print(f"分册模式：每册 {merge_value} 条路线")
                    merge_by_volumes.run_batch_with_routes_per_volume(merge_value)
                elif merge_option == "total_volumes":
                    print(f"分册模式：总共分为 {merge_value} 册")
                    merge_by_volumes.run_batch_with_total_volumes(merge_value)
                else:
                    merge_by_volumes.run_batch()
                
                print("\n>>> 🎉 所有任务执行完毕！")
        except Exception as e:
            print(f"\n!!! 发生错误: {str(e)}")
            print(traceback.format_exc())
        finally:
            # Restore stdout
            sys.stdout = original_stdout
            self.finish_func()

class WorkerSignals(QObject):
    log = pyqtSignal(str)
    finished = pyqtSignal()

class CardFrame(QFrame):
    """自定义卡片控件，带阴影和圆角"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("CardFrame")
        # 添加阴影效果
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setColor(QColor(0, 0, 0, 20))
        shadow.setOffset(0, 2)
        self.setGraphicsEffect(shadow)

class ModernWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("DGSS 野外路线电子手簿一键排版工具")
        self.resize(1280, 800)
        self.setObjectName("MainWindow")
        
        # 设置窗口图标
        icon_path = resource_path("app_icon.png")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        
        # 信号处理
        self.signals = WorkerSignals()
        self.signals.log.connect(self.append_log_safe)
        self.signals.finished.connect(self.on_task_finished)
        
        # 业务逻辑
        self.logic = AppLogic(self.emit_log, self.emit_finished)
        
        # 初始化界面布局
        self.setup_ui()
        
        # 应用样式表
        self.apply_stylesheet()

    def emit_log(self, text):
        self.signals.log.emit(text)

    def emit_finished(self):
        self.signals.finished.emit()

    def append_log_safe(self, message):
        self.add_log(message)

    def on_task_finished(self):
        self.add_log(">>> 就绪。", "#22c55e")
        self.enable_buttons(True)

    def enable_buttons(self, enabled):
        self.btn_auto.setEnabled(enabled)
        for btn in self.step_buttons:
            btn.setEnabled(enabled)

    def start_task(self, task_type):
        self.enable_buttons(False)
        self.log_area.clear()
        self.add_log(f"正在启动任务: {task_type}...", "#3b82f6")
        
        merge_option, merge_value = self.get_merge_args()
        
        # 在后台线程运行
        threading.Thread(
            target=self.logic.run_task, 
            args=(task_type, merge_option, merge_value), 
            daemon=True
        ).start()

    def get_merge_args(self):
        if self.radio_routes.isChecked():
            # 从输入框获取路线数
            try:
                val = int(self.input_routes.text().strip())
                if val <= 0:
                    raise ValueError("路线数必须大于0")
                return "routes_per_volume", val
            except ValueError:
                self.add_log("错误：请输入有效的正整数作为每册路线数", "#e74c3c")
                return None, None
        elif self.radio_volumes.isChecked():
            # 从输入框获取总册数
            try:
                val = int(self.input_volumes.text().strip())
                if val <= 0:
                    raise ValueError("总册数必须大于0")
                return "total_volumes", val
            except ValueError:
                self.add_log("错误：请输入有效的正整数作为总册数", "#e74c3c")
                return None, None
        return None, None

    def setup_ui(self):
        """构建整体UI骨架"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # 全局垂直布局
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 1. 顶部 Header 区域
        self.create_header(main_layout)

        # 2. 下方内容区域 (使用浅灰背景容器)
        content_container = QWidget()
        content_container.setObjectName("ContentContainer")
        content_layout = QHBoxLayout(content_container)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(20)

        # 2.1 左侧控制面板
        self.create_left_panel(content_layout)

        # 2.2 右侧主区域 (日志 + 开发者信息)
        self.create_right_panel(content_layout)

        main_layout.addWidget(content_container)

    def create_header(self, parent_layout):
        """创建顶部标题区域"""
        header_frame = QFrame()
        header_frame.setObjectName("HeaderFrame")
        header_frame.setFixedHeight(80)
        
        v_layout = QVBoxLayout(header_frame)
        v_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # 主标题
        title = QLabel("DGSS 野外路线电子手簿一键排版工具")
        title.setObjectName("MainTitle")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        v_layout.addWidget(title)
        
        parent_layout.addWidget(header_frame)

    def create_left_panel(self, parent_layout):
        """创建左侧功能控制区 (紧凑模式)"""
        left_widget = QWidget()
        left_widget.setFixedWidth(340)
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(15)

        # --- 模块1: 智能自动化 ---
        group_auto = CardFrame()
        layout_auto = QVBoxLayout(group_auto)
        layout_auto.setContentsMargins(15, 15, 15, 15)
        
        lbl_auto = QLabel("🚀 智能自动化")
        lbl_auto.setObjectName("CardTitle")
        layout_auto.addWidget(lbl_auto)
        
        # 绿色大按钮
        self.btn_auto = QPushButton("一键全自动运行  (推荐)")
        self.btn_auto.setObjectName("BtnGreen")
        self.btn_auto.setFixedHeight(50)
        self.btn_auto.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_auto.clicked.connect(lambda: self.start_task("all"))
        layout_auto.addWidget(self.btn_auto)
        
        left_layout.addWidget(group_auto)

        # --- 模块2: 分步操作 ---
        group_step = CardFrame()
        layout_step = QVBoxLayout(group_step)
        layout_step.setContentsMargins(15, 15, 15, 15)
        layout_step.setSpacing(10)
        
        lbl_step = QLabel("🛠 分步操作")
        lbl_step.setObjectName("CardTitle")
        layout_step.addWidget(lbl_step)
        
        self.step_buttons = []
        steps = [
            ("1. 格式化文档 (Format)", "format"),
            ("2. 提取素描图 (Extract)", "extract"),
            ("3. 插入素描图 (Insert)", "insert"),
            ("4. 分册合并 (Merge)", "merge")
        ]
        
        for text, task_key in steps:
            btn = QPushButton(text)
            btn.setObjectName("BtnBlue")
            btn.setFixedHeight(40)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            # 使用闭包捕获 task_key
            btn.clicked.connect(lambda checked, k=task_key: self.start_task(k))
            layout_step.addWidget(btn)
            self.step_buttons.append(btn)
            
        left_layout.addWidget(group_step)
        
        # left_layout.addStretch() 

        # --- 模块3: 分册设置 ---
        group_settings = CardFrame()
        size_policy = QSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
        group_settings.setSizePolicy(size_policy)
        layout_settings = QVBoxLayout(group_settings)
        layout_settings.setContentsMargins(15, 15, 15, 15)
        layout_settings.setSpacing(10)

        lbl_set = QLabel("⚙️ 分册设置")
        lbl_set.setObjectName("CardTitle")
        layout_settings.addWidget(lbl_set)
        
        # 单选按钮组
        radio_layout = QHBoxLayout()
        self.radio_default = QRadioButton("默认")
        self.radio_default.setChecked(True)
        self.radio_routes = QRadioButton("指定路线")
        self.radio_volumes = QRadioButton("指定总册")
        
        bg = QButtonGroup(self)
        bg.addButton(self.radio_default)
        bg.addButton(self.radio_routes)
        bg.addButton(self.radio_volumes)
        
        radio_layout.addWidget(self.radio_default)
        radio_layout.addWidget(self.radio_routes)
        radio_layout.addWidget(self.radio_volumes)
        radio_layout.addStretch()
        layout_settings.addLayout(radio_layout)
        
        # 参数输入框
        param_layout = QHBoxLayout()
        lbl_param = QLabel("参数:")
        lbl_param.setStyleSheet("font-weight: bold; color: #64748b;")
        
        # 路线数量输入框
        route_input_layout = QVBoxLayout()
        lbl_routes = QLabel("每册路线数:")
        lbl_routes.setStyleSheet("color: #64748b; font-size: 12px;")
        self.input_routes = QLineEdit()
        self.input_routes.setPlaceholderText("如: 12")
        self.input_routes.setFixedHeight(32)
        self.input_routes.setEnabled(False)
        route_input_layout.addWidget(lbl_routes)
        route_input_layout.addWidget(self.input_routes)
        
        # 总册数输入框
        volume_input_layout = QVBoxLayout()
        lbl_volumes = QLabel("总册数:")
        lbl_volumes.setStyleSheet("color: #64748b; font-size: 12px;")
        self.input_volumes = QLineEdit()
        self.input_volumes.setPlaceholderText("如: 3")
        self.input_volumes.setFixedHeight(32)
        self.input_volumes.setEnabled(False)
        volume_input_layout.addWidget(lbl_volumes)
        volume_input_layout.addWidget(self.input_volumes)

        # 联动逻辑
        self.radio_default.toggled.connect(self.update_input_state)
        self.radio_routes.toggled.connect(self.update_input_state)
        self.radio_volumes.toggled.connect(self.update_input_state)

        param_layout.addWidget(lbl_param)
        param_layout.addLayout(route_input_layout)
        param_layout.addLayout(volume_input_layout)
        layout_settings.addLayout(param_layout)
        layout_settings.addStretch()
        
        left_layout.addWidget(group_settings)

        parent_layout.addWidget(left_widget)

    def update_input_state(self):
        # 当选择默认模式时，禁用所有输入框
        if self.radio_default.isChecked():
            self.input_routes.setEnabled(False)
            self.input_volumes.setEnabled(False)
        # 当选择指定路线模式时，只启用路线输入框
        elif self.radio_routes.isChecked():
            self.input_routes.setEnabled(True)
            self.input_volumes.setEnabled(False)
            self.input_routes.setFocus()  # 自动聚焦到输入框
        # 当选择指定总册模式时，只启用总册输入框
        elif self.radio_volumes.isChecked():
            self.input_routes.setEnabled(False)
            self.input_volumes.setEnabled(True)
            self.input_volumes.setFocus()  # 自动聚焦到输入框

    def create_right_panel(self, parent_layout):
        """创建右侧区域 (上下两部分)"""
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(20)
        
        # --- 上半部分: 运行日志 ---
        log_frame = QFrame()
        log_frame.setObjectName("LogFrame")
        log_layout = QVBoxLayout(log_frame)
        log_layout.setContentsMargins(0, 0, 0, 0)
        log_layout.setSpacing(0)
        
        # 日志标题栏
        log_header = QFrame()
        log_header.setObjectName("LogHeader")
        log_header.setFixedHeight(40)
        header_layout = QHBoxLayout(log_header)
        header_layout.setContentsMargins(15, 0, 15, 0)
        
        lbl_log_title = QLabel("📄 运行日志控制台")
        lbl_log_title.setStyleSheet("color: #f1f5f9; font-weight: bold; font-size: 14px;")
        
        lbl_status = QLabel("System Ready")
        lbl_status.setObjectName("StatusBadge")
        
        header_layout.addWidget(lbl_log_title)
        header_layout.addStretch()
        header_layout.addWidget(lbl_status)
        
        log_layout.addWidget(log_header)
        
        # 日志内容区
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setObjectName("LogArea")
        self.add_log("系统初始化完成...")
        self.add_log("加载配置: Default_Config.json")
        self.add_log(f"当前工作目录: {os.getcwd()}")
        self.add_log("等待用户指令...", color="#4ade80")
        
        log_layout.addWidget(self.log_area)
        
        # 设置上半部分的高度比例 (55%)
        size_policy_log = QSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
        size_policy_log.setVerticalStretch(55)
        log_frame.setSizePolicy(size_policy_log)
        
        right_layout.addWidget(log_frame)
        
        # --- 下半部分: 开发者信息 & 二维码 ---
        dev_frame = CardFrame()
        dev_frame.setStyleSheet("#CardFrame { background-color: white; border: 1px solid #e2e8f0; border-radius: 8px; }")
        
        dev_layout = QHBoxLayout(dev_frame)
        dev_layout.setContentsMargins(0, 0, 0, 0)
        dev_layout.setSpacing(0)
        
        # 左侧: 信息
        info_widget = QWidget()
        info_layout = QVBoxLayout(info_widget)
        info_layout.setContentsMargins(20, 40, 20, 40)
        info_layout.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignLeft)
        
        lbl_dev_title = QLabel("开发者信息")
        lbl_dev_title.setStyleSheet("font-size: 20px; font-weight: bold; color: #1e293b;")
        
        line_blue = QFrame()
        line_blue.setFixedSize(50, 4)
        line_blue.setStyleSheet("background-color: #3b82f6; border-radius: 2px;")
        
        info_layout.addWidget(lbl_dev_title)
        info_layout.addWidget(line_blue)
        info_layout.addSpacing(20)
        
        # 信息行生成函数
        def create_info_row(label, text, is_email=False):
            row = QWidget()
            h = QHBoxLayout(row)
            h.setContentsMargins(0, 0, 0, 0)
            h.setSpacing(15)
            
            lbl = QLabel(label)
            lbl.setFixedWidth(50)
            lbl.setStyleSheet("font-weight: bold; color: #334155;")
            
            val = QLabel(text)
            if is_email:
                val.setStyleSheet("background-color: #f8fafc; border: 1px solid #f1f5f9; padding: 2px 6px; border-radius: 4px; font-family: Consolas; color: #475569;")
            else:
                val.setStyleSheet("color: #475569; font-weight: 500;")
                # val.setWordWrap(True)
                
            h.addWidget(lbl)
            h.addWidget(val)
            h.addStretch()
            return row

        info_layout.addWidget(create_info_row("单位:", "浙江省宁波地质院 基础地质调查研究中心"))
        info_layout.addWidget(create_info_row("姓名:", "丁正鹏"))
        info_layout.addWidget(create_info_row("邮箱:", "zhengpengding@outlook.com", is_email=True))
        info_layout.addStretch()

        dev_layout.addWidget(info_widget)
        
        # 右侧: 二维码区域
        qr_container = QWidget()
        qr_container.setFixedWidth(340)
        qr_container.setStyleSheet("background-color: #f8fafc; border-left: 1px solid #e2e8f0; border-top-right-radius: 8px; border-bottom-right-radius: 8px;")
        qr_layout = QVBoxLayout(qr_container)
        qr_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        # 二维码图片
        qr_label = QLabel()
        qr_label.setFixedSize(300, 300)
        qr_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        qr_path = resource_path("recived money.png")
        if os.path.exists(qr_path):
            pixmap = QPixmap(qr_path)
            qr_label.setPixmap(pixmap.scaled(300, 300, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation))
        else:
            qr_label.setText("QR CODE\nNot Found")
            qr_label.setStyleSheet("""
                background-color: #1e293b; 
                border: 4px solid white; 
                border-radius: 8px;
                color: rgba(255,255,255,0.5);
                font-size: 12px;
            """)

        # 阴影给二维码
        qr_shadow = QGraphicsDropShadowEffect()
        qr_shadow.setBlurRadius(10)
        qr_shadow.setColor(QColor(0,0,0,30))
        qr_label.setGraphicsEffect(qr_shadow)
        
        lbl_coffee = QLabel("☕ 请作者喝杯咖啡")
        lbl_coffee.setStyleSheet("color: #e67e22; font-weight: bold; margin-top: 10px;")
        lbl_coffee.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        qr_layout.addWidget(qr_label)
        qr_layout.addWidget(lbl_coffee)
        
        dev_layout.addWidget(qr_container)
        
        # 设置下半部分的高度比例 (45%)
        size_policy_dev = QSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Expanding)
        size_policy_dev.setVerticalStretch(45)
        dev_frame.setSizePolicy(size_policy_dev)
        
        right_layout.addWidget(dev_frame)

        parent_layout.addWidget(right_widget)

    def add_log(self, message, color=None):
        """向日志控制台添加信息"""
        now = datetime.datetime.now().strftime("%H:%M:%S")
        timestamp = f'<span style="color: #64748b;">[{now}]</span> '
        
        # 简单的HTML转义，防止内容破坏格式
        message = str(message).replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>")
        
        if color:
            msg_content = f'<span style="color: {color}; font-weight: bold;">{message}</span>'
        else:
            msg_content = f'<span style="color: #cbd5e1;">{message}</span>'
            
        full_html = f'<div style="margin-bottom: 4px; border-left: 3px solid #22c55e; padding-left: 8px;">{timestamp}{msg_content}</div>'
        self.log_area.append(full_html)
        
        # 自动滚动到底部
        scrollbar = self.log_area.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def apply_stylesheet(self):
        style = """
        /* 全局字体与背景 - 使用渐变背景 */
        QWidget {
            font-family: 'Microsoft YaHei', 'Segoe UI', 'PingFang SC', sans-serif;
            font-size: 14px;
            color: #1e293b;
        }
        QMainWindow {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                       stop:0 #f8fafc, stop:1 #e0e7ff);
        }
        
        /* 顶部 Header - 天蓝渐变 */
        #HeaderFrame {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                       stop:0 #0ea5e9, stop:1 #38bdf8);
            border-bottom: none;
            padding: 10px;
        }
        #MainTitle {
            font-size: 26px;
            font-weight: bold;
            color: white;
            letter-spacing: 1px;
        }

        /* 浅色背景容器 */
        #ContentContainer {
            background-color: transparent;
        }

        /* 卡片通用样式 - 毛玻璃效果 */
        #CardFrame {
            background-color: rgba(255, 255, 255, 0.9);
            border-radius: 12px;
            border: 1px solid rgba(148, 163, 184, 0.2);
        }
        #CardTitle {
            font-weight: bold;
            font-size: 15px;
            color: #334155;
            margin-bottom: 8px;
        }

        /* 绿色渐变按钮 - 更圆润 */
        #BtnGreen {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                       stop:0 #10b981, stop:1 #059669);
            color: white;
            border-radius: 10px;
            font-size: 16px;
            font-weight: bold;
            border: none;
            padding: 12px 20px;
            letter-spacing: 1px;
        }
        #BtnGreen:hover {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                       stop:0 #059669, stop:1 #047857);
        }
        #BtnGreen:pressed {
            background: #047857;
            padding-top: 14px;
            padding-bottom: 10px;
        }
        #BtnGreen:disabled {
            background-color: #cbd5e1;
        }

        /* 天蓝渐变按钮 */
        #BtnBlue {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                       stop:0 #0ea5e9, stop:1 #0284c7);
            color: white;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            border: none;
            text-align: left;
            padding: 10px 16px;
        }
        #BtnBlue:hover {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                       stop:0 #0284c7, stop:1 #0369a1);
        }
        #BtnBlue:pressed {
            background: #075985;
            padding-top: 11px;
            padding-bottom: 9px;
        }
        #BtnBlue:disabled {
            background-color: #cbd5e1;
        }

        /* 单选按钮 - 现代化样式 */
        QRadioButton {
            spacing: 8px;
            color: #475569;
            font-weight: 500;
        }
        QRadioButton::indicator {
            width: 18px;
            height: 18px;
            border-radius: 9px;
            border: 2px solid #94a3b8;
            background-color: white;
        }
        QRadioButton::indicator:hover {
            border-color: #0ea5e9;
        }
        QRadioButton::indicator:checked {
            background-color: #0ea5e9;
            border-color: #0ea5e9;
        }
        
        /* 下拉框 */
        QComboBox {
            border: 2px solid #e2e8f0;
            border-radius: 6px;
            padding: 6px 10px;
            background-color: white;
            font-weight: 500;
        }
        QComboBox:hover {
            border-color: #a5b4fc;
        }
        QComboBox:focus {
            border-color: #0ea5e9;
        }
        QComboBox::drop-down {
            subcontrol-origin: padding;
            subcontrol-position: top right;
            width: 25px;
            border-left: none;
        }
        
        /* 输入框 - 现代化 */
        QLineEdit {
            border: 2px solid #e2e8f0;
            border-radius: 8px;
            padding: 8px 12px;
            background-color: white;
            font-size: 14px;
            font-weight: 500;
            selection-background-color: #a5b4fc;
        }
        QLineEdit:hover {
            border-color: #cbd5e1;
        }
        QLineEdit:focus {
            border-color: #0ea5e9;
            outline: none;
            background-color: #f0f9ff;
        }
        QLineEdit:disabled {
            background-color: #f1f5f9;
            color: #94a3b8;
            border-color: #e2e8f0;
        }

        /* 日志控制台 - 深色主题优化 */
        #LogFrame {
            background-color: #1e293b;
            border-radius: 12px;
            border: 1px solid #334155;
        }
        #LogHeader {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                       stop:0 #334155, stop:1 #475569);
            border-top-left-radius: 12px;
            border-top-right-radius: 12px;
            border-bottom: 1px solid #475569;
        }
        #StatusBadge {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                                       stop:0 #059669, stop:1 #10b981);
            color: white;
            font-size: 10px;
            font-weight: bold;
            padding: 4px 10px;
            border-radius: 10px;
            border: none;
        }
        #LogArea {
            background-color: #1e293b;
            border: none;
            color: #e2e8f0;
            font-family: 'Consolas', 'Courier New', 'JetBrains Mono', monospace;
            font-size: 13px;
            padding: 12px;
            border-bottom-left-radius: 12px;
            border-bottom-right-radius: 12px;
            selection-background-color: #475569;
        }
        
        /* 滚动条美化 */
        QScrollBar:vertical {
            border: none;
            background-color: #1e293b;
            width: 10px;
            border-radius: 5px;
        }
        QScrollBar::handle:vertical {
            background-color: #475569;
            min-height: 30px;
            border-radius: 5px;
        }
        QScrollBar::handle:vertical:hover {
            background-color: #64748b;
        }
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        """
        self.setStyleSheet(style)

if __name__ == "__main__":
    # 高DPI缩放策略
    if hasattr(Qt.HighDpiScaleFactorRoundingPolicy, 'PassThrough'):
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
            
    app = QApplication(sys.argv)
    window = ModernWindow()
    window.show()
    sys.exit(app.exec())
