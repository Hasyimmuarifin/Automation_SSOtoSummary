from PyQt6 import QtCore, QtGui, QtWidgets
from utils.resources import resource_path
import os, sys
import datetime
from typing import List

# ---------- Helper ----------
class ResourceHelper:
    """
    Helper class to manage resource paths (files, assets, styles).
    Ensures compatibility when running as an executable (frozen) 
    or directly from source code.
    """

    @staticmethod
    def get_path(relative_path: str) -> str:
        base_path = (
            os.path.dirname(sys.executable)
            if getattr(sys, "frozen", False)
            else os.path.dirname(os.path.abspath(__file__))
        )
        return os.path.join(base_path, relative_path)

# ---------- UI ----------
class Ui_MainWindow(object):
    """
    Class to define the main application UI with modern design.
    """

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 700)
        MainWindow.setMinimumSize(800, 700)
        MainWindow.setMaximumSize(800, 700)  # Fixed size to prevent scrolling

        # === Load stylesheet (QSS) ===
        qss_path = resource_path("style/style.qss")
        if os.path.exists(qss_path):
            try:
                with open(qss_path, "r", encoding="utf-8") as f:
                    MainWindow.setStyleSheet(f.read())
            except Exception:
                MainWindow.setStyleSheet(self.get_default_stylesheet())
        else:
            MainWindow.setStyleSheet(self.get_default_stylesheet())

        # === Central widget & main layout ===
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        
        # # Set background image
        # bg_path = resource_path("assets/bgbg.png")
        # if os.path.exists(bg_path):
        #     self.centralwidget.setStyleSheet(f"""
        #         QWidget#centralwidget {{
        #             background-image: url({bg_path.replace(os.sep, '/')});
        #             background-position: center;
        #             background-repeat: no-repeat;
        #             background-attachment: fixed;
        #         }}
        #     """)
        
        # Main layout without scroll area
        self.vlayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.vlayout.setContentsMargins(40, 25, 40, 25)
        self.vlayout.setSpacing(15)

        # === Header Container ===
        self.header_container = QtWidgets.QWidget()
        self.header_container.setObjectName("header_container")
        self.header_layout = QtWidgets.QVBoxLayout(self.header_container)
        self.header_layout.setContentsMargins(0, 0, 0, 0)
        self.header_layout.setSpacing(8)

        # === Company Logo ===
        self.logo_container = QtWidgets.QWidget()
        self.logo_layout = QtWidgets.QVBoxLayout(self.logo_container)
        self.logo_layout.setContentsMargins(0, 0, 0, 0)
        
        self.logo = QtWidgets.QLabel()
        logo_path = resource_path("assets/ITM_logo.png")
        if os.path.exists(logo_path):
            pixmap = QtGui.QPixmap(logo_path)
            self.logo.setPixmap(pixmap.scaled(150, 75, QtCore.Qt.AspectRatioMode.KeepAspectRatio, 
                                             QtCore.Qt.TransformationMode.SmoothTransformation))
        self.logo.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.logo.setObjectName("label_logo")
        self.logo_layout.addWidget(self.logo)
        
        self.header_layout.addWidget(self.logo_container)

        # === Company Name ===
        self.label_company = QtWidgets.QLabel("PT Indo Tambangraya Megah Tbk")
        self.label_company.setFont(QtGui.QFont("Segoe UI", 13, QtGui.QFont.Weight.Bold))
        self.label_company.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label_company.setObjectName("label_company")
        self.header_layout.addWidget(self.label_company)

        # === Application Title ===
        self.label_title = QtWidgets.QLabel("SSO Results Automation System")
        self.label_title.setFont(QtGui.QFont("Segoe UI", 18, QtGui.QFont.Weight.Bold))
        self.label_title.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label_title.setObjectName("label_title")
        self.header_layout.addWidget(self.label_title)

        # === Application Subtitle ===
        self.label_sub = QtWidgets.QLabel("Select input and output files to begin automated processing")
        self.label_sub.setFont(QtGui.QFont("Segoe UI", 9))
        self.label_sub.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label_sub.setObjectName("label_sub")
        self.header_layout.addWidget(self.label_sub)

        self.vlayout.addWidget(self.header_container)

        # === Month Selection Card ===
        self.month_card = QtWidgets.QFrame()
        self.month_card.setObjectName("month_card")
        self.month_card_layout = QtWidgets.QVBoxLayout(self.month_card)
        self.month_card_layout.setContentsMargins(25, 15, 25, 15)
        self.month_card_layout.setSpacing(12)

        # Card shadow effect
        card_shadow = QtWidgets.QGraphicsDropShadowEffect()
        card_shadow.setBlurRadius(20)
        card_shadow.setOffset(0, 4)
        card_shadow.setColor(QtGui.QColor(0, 0, 0, 60))
        self.month_card.setGraphicsEffect(card_shadow)

        # Month data
        months: List[str] = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]

        # === This Month Section ===
        this_month_container = QtWidgets.QWidget()
        this_month_layout = QtWidgets.QHBoxLayout(this_month_container)
        this_month_layout.setContentsMargins(0, 0, 0, 0)
        this_month_layout.setSpacing(12)

        self.this_month_label = QtWidgets.QLabel("📅 This Month")
        self.this_month_label.setFont(QtGui.QFont("Segoe UI", 10, QtGui.QFont.Weight.Medium))
        this_month_layout.addWidget(self.this_month_label)
        this_month_layout.addStretch()

        self.month_combo = QtWidgets.QComboBox()
        self.month_combo.setMinimumWidth(140)
        self.month_combo.setFixedHeight(35)
        self.month_combo.addItems(months)
        self.month_combo.setObjectName("combo_this_month")
        
        delegate = QtWidgets.QStyledItemDelegate(self.month_combo)
        self.month_combo.setItemDelegate(delegate)
        for i in range(self.month_combo.count()):
            self.month_combo.setItemData(i, QtCore.Qt.AlignmentFlag.AlignCenter, 
                                        QtCore.Qt.ItemDataRole.TextAlignmentRole)
        
        this_month_layout.addWidget(self.month_combo)
        self.month_card_layout.addWidget(this_month_container)

        # Separator line
        separator = QtWidgets.QFrame()
        separator.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        separator.setObjectName("separator")
        self.month_card_layout.addWidget(separator)

        # === Next Month Section ===
        next_month_container = QtWidgets.QWidget()
        next_month_layout = QtWidgets.QHBoxLayout(next_month_container)
        next_month_layout.setContentsMargins(0, 0, 0, 0)
        next_month_layout.setSpacing(12)

        self.next_month_label = QtWidgets.QLabel("📅 Next Month")
        self.next_month_label.setFont(QtGui.QFont("Segoe UI", 10, QtGui.QFont.Weight.Medium))
        next_month_layout.addWidget(self.next_month_label)
        next_month_layout.addStretch()

        self.checkBox_enableNextMonth = QtWidgets.QCheckBox("Enable")
        self.checkBox_enableNextMonth.setFont(QtGui.QFont("Segoe UI", 9))
        self.checkBox_enableNextMonth.setObjectName("checkBox_enableNextMonth")
        next_month_layout.addWidget(self.checkBox_enableNextMonth)

        self.month_combo2 = QtWidgets.QComboBox()
        self.month_combo2.setMinimumWidth(140)
        self.month_combo2.setFixedHeight(35)
        self.month_combo2.addItems(months)
        self.month_combo2.setEnabled(False)
        self.month_combo2.setObjectName("combo_next_month")
        
        delegate2 = QtWidgets.QStyledItemDelegate(self.month_combo2)
        self.month_combo2.setItemDelegate(delegate2)
        for i in range(self.month_combo2.count()):
            self.month_combo2.setItemData(i, QtCore.Qt.AlignmentFlag.AlignCenter, 
                                         QtCore.Qt.ItemDataRole.TextAlignmentRole)
        
        self.checkBox_enableNextMonth.toggled.connect(self.month_combo2.setEnabled)
        next_month_layout.addWidget(self.month_combo2)
        self.month_card_layout.addWidget(next_month_container)

        self.vlayout.addWidget(self.month_card)

        # === File Selection Card ===
        self.file_card = QtWidgets.QFrame()
        self.file_card.setObjectName("file_card")
        self.file_layout = QtWidgets.QVBoxLayout(self.file_card)
        self.file_layout.setContentsMargins(25, 20, 25, 20)
        self.file_layout.setSpacing(15)

        # Card shadow
        file_shadow = QtWidgets.QGraphicsDropShadowEffect()
        file_shadow.setBlurRadius(20)
        file_shadow.setOffset(0, 4)
        file_shadow.setColor(QtGui.QColor(0, 0, 0, 60))
        self.file_card.setGraphicsEffect(file_shadow)

        # === Input File Section ===
        input_label = QtWidgets.QLabel("📁 Input File")
        input_label.setFont(QtGui.QFont("Segoe UI", 10, QtGui.QFont.Weight.Medium))
        input_label.setObjectName("section_label")
        self.file_layout.addWidget(input_label)

        self.input_layout = QtWidgets.QHBoxLayout()
        self.input_layout.setSpacing(10)
        
        self.input_line = QtWidgets.QLineEdit()
        self.input_line.setReadOnly(True)
        self.input_line.setPlaceholderText("No file selected...")
        self.input_line.setMinimumHeight(38)
        self.input_line.setObjectName("input_line")
        self.input_layout.addWidget(self.input_line, stretch=1)

        self.input_btn = QtWidgets.QPushButton("Browse")
        self.input_btn.setObjectName("btnInput")
        self.input_btn.setFixedSize(110, 38)
        self.input_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.input_layout.addWidget(self.input_btn)

        self.file_layout.addLayout(self.input_layout)

        # === Output File Section ===
        output_label = QtWidgets.QLabel("📁 Output File")
        output_label.setFont(QtGui.QFont("Segoe UI", 10, QtGui.QFont.Weight.Medium))
        output_label.setObjectName("section_label")
        self.file_layout.addWidget(output_label)

        self.output_layout = QtWidgets.QHBoxLayout()
        self.output_layout.setSpacing(10)
        
        self.output_line = QtWidgets.QLineEdit()
        self.output_line.setReadOnly(True)
        self.output_line.setPlaceholderText("No file selected...")
        self.output_line.setMinimumHeight(38)
        self.output_line.setObjectName("output_line")
        self.output_layout.addWidget(self.output_line, stretch=1)

        self.output_btn = QtWidgets.QPushButton("Browse")
        self.output_btn.setObjectName("btnOutput")
        self.output_btn.setFixedSize(110, 38)
        self.output_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.output_layout.addWidget(self.output_btn)

        self.file_layout.addLayout(self.output_layout)

        # === Start Button ===
        self.start_btn = QtWidgets.QPushButton("▶  Start Processing")
        self.start_btn.setObjectName("btnStart")
        self.start_btn.setMinimumSize(180, 48)
        self.start_btn.setFont(QtGui.QFont("Segoe UI", 11, QtGui.QFont.Weight.Bold))
        self.start_btn.setCursor(QtGui.QCursor(QtCore.Qt.CursorShape.PointingHandCursor))
        self.file_layout.addWidget(self.start_btn, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

        self.vlayout.addWidget(self.file_card)

        # === Footer ===
        self.vlayout.addStretch()
        self.footer = QtWidgets.QLabel("© 2025 PT Indo Tambangraya Megah Tbk  •  Automated SSO Processing System")
        self.footer.setFont(QtGui.QFont("Segoe UI", 8))
        self.footer.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.footer.setObjectName("label_footer")
        self.vlayout.addWidget(self.footer)
        
        MainWindow.setCentralWidget(self.centralwidget)

        # Set default to current month
        current_month = datetime.datetime.now().month
        self.month_combo.setCurrentIndex(current_month - 1)
        self.month_combo2.setCurrentIndex(current_month - 1)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "SSO Automation System"))

    def get_default_stylesheet(self) -> str:
        """Modern stylesheet with semi-transparent cards"""
        return """
            QWidget#centralwidget {
                background: transparent;
            }
            
            QWidget#header_container {
                background: transparent;
            }
            
            QLabel#label_company {
                color: #1a1a1a;
                padding: 3px;
                background: transparent;
            }
            
            QLabel#label_title {
                color: #0d47a1;
                padding: 5px;
                background: transparent;
            }
            
            QLabel#label_sub {
                color: #424242;
                padding: 3px;
                background: transparent;
            }
            
            QLabel#section_label {
                color: #2c3e50;
                margin-bottom: 5px;
            }
            
            QFrame#month_card, QFrame#file_card {
                background: rgba(255, 255, 255, 0.92);
                border-radius: 14px;
                border: 1px solid rgba(224, 230, 237, 0.8);
            }
            
            QFrame#separator {
                background: #e0e6ed;
                max-height: 1px;
                border: none;
            }
            
            QComboBox {
                background: #f8fafc;
                border: 2px solid #d1dce5;
                border-radius: 8px;
                padding: 6px 12px;
                font-size: 10pt;
                color: #2c3e50;
                font-family: "Segoe UI";
            }
            
            QComboBox:hover {
                border-color: #3498db;
                background: white;
            }
            
            QComboBox:focus {
                border-color: #2980b9;
            }
            
            QComboBox::drop-down {
                border: none;
                width: 28px;
            }
            
            QComboBox::down-arrow {
                image: none;
                border-left: 4px solid transparent;
                border-right: 4px solid transparent;
                border-top: 5px solid #2c3e50;
                margin-right: 8px;
            }
            
            QComboBox QAbstractItemView {
                background: white;
                border: 2px solid #d1dce5;
                border-radius: 8px;
                selection-background-color: #3498db;
                selection-color: white;
                padding: 4px;
                outline: none;
            }
            
            QCheckBox {
                spacing: 6px;
                color: #2c3e50;
            }
            
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 4px;
                border: 2px solid #d1dce5;
                background: #f8fafc;
            }
            
            QCheckBox::indicator:hover {
                border-color: #3498db;
            }
            
            QCheckBox::indicator:checked {
                background: #3498db;
                border-color: #3498db;
            }
            
            QLineEdit {
                background: #f8fafc;
                border: 2px solid #d1dce5;
                border-radius: 8px;
                padding: 8px 12px;
                font-size: 9pt;
                color: #2c3e50;
                font-family: "Segoe UI";
            }
            
            QLineEdit:focus {
                border-color: #3498db;
                background: white;
            }
            
            QPushButton#btnInput, QPushButton#btnOutput {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #4a90e2, stop:1 #357abd);
                color: white;
                border: none;
                border-radius: 8px;
                font-size: 9pt;
                font-weight: 600;
                font-family: "Segoe UI";
            }
            
            QPushButton#btnInput:hover, QPushButton#btnOutput:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #5ba3f5, stop:1 #4a90e2);
            }
            
            QPushButton#btnInput:pressed, QPushButton#btnOutput:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #357abd, stop:1 #2c6aa0);
            }
            
            QPushButton#btnStart {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #27ae60, stop:1 #229954);
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 11pt;
                font-weight: bold;
                padding: 12px 25px;
                font-family: "Segoe UI";
            }
            
            QPushButton#btnStart:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2ecc71, stop:1 #27ae60);
            }
            
            QPushButton#btnStart:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #229954, stop:1 #1e8449);
            }
            
            QPushButton:disabled {
                background: #bdc3c7;
                color: #7f8c8d;
            }
            
            QLabel#label_footer {
                color: #424242;
                padding: 10px;
                background: transparent;
            }
        """