from PyQt6 import QtCore, QtGui, QtWidgets
from utils.resources import resource_path
import os, sys

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

class Ui_MainWindow(object):
    """
    Class to define the main application UI.
    """

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(700, 600)

        # === Load stylesheet (QSS) ===
        qss_path = resource_path("style/style.qss")
        if os.path.exists(qss_path):
            with open(qss_path, "r") as f:
                MainWindow.setStyleSheet(f.read())

        # === Central widget & main layout ===
        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.vlayout = QtWidgets.QVBoxLayout(self.centralwidget)
        self.vlayout.setContentsMargins(40, 40, 40, 40)
        self.vlayout.setSpacing(20)

        # === Company Logo ===
        self.logo = QtWidgets.QLabel()
        logo_path = resource_path("assets/ITM_logo.png")
        if os.path.exists(logo_path):
            self.logo.setPixmap(QtGui.QPixmap(logo_path))
        self.logo.setScaledContents(True)
        self.logo.setFixedSize(180, 90)
        self.logo.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.vlayout.addWidget(self.logo, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

        # === Company Name ===
        self.label_company = QtWidgets.QLabel("PT Indo Tambangraya Megah Tbk")
        font = QtGui.QFont("Arial", 14, QtGui.QFont.Weight.Bold)
        self.label_company.setFont(font)
        self.label_company.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.vlayout.addWidget(self.label_company)

        # === Application Title ===
        self.label_title = QtWidgets.QLabel("SSO Results Automation System")
        font_title = QtGui.QFont("Arial", 20, QtGui.QFont.Weight.Bold)
        self.label_title.setFont(font_title)
        self.label_title.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.vlayout.addWidget(self.label_title)

        # === Application Subtitle ===
        self.label_sub = QtWidgets.QLabel("Please select the input and output file to begin automated processing.")
        font_sub = QtGui.QFont("Arial", 10)
        self.label_sub.setFont(font_sub)
        self.label_sub.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.vlayout.addWidget(self.label_sub)

        # === This Month Label ===
        self.this_month = QtWidgets.QLabel("This Month :")
        font_sub = QtGui.QFont("Arial", 10)
        self.this_month.setFont(font_sub)
        self.this_month.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.vlayout.addWidget(self.this_month)

        # === Month ComboBox ===
        self.month_combo = QtWidgets.QComboBox()
        self.month_combo.setFixedWidth(120)
        self.month_combo.setFixedHeight(35)

        # Tambahkan bulan January - December
        self.month_combo.addItems([
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ])
        # Buat delegate agar teks di tengah
        delegate = QtWidgets.QStyledItemDelegate(self.month_combo)
        self.month_combo.setItemDelegate(delegate)

        for i in range(self.month_combo.count()):
            self.month_combo.setItemData(i, QtCore.Qt.AlignmentFlag.AlignCenter, QtCore.Qt.ItemDataRole.TextAlignmentRole)

        # Set default ke bulan saat ini
        import datetime
        current_month = datetime.datetime.now().month
        self.month_combo.setCurrentIndex(current_month - 1)

        # Tambahkan ke layout utama
        self.vlayout.addWidget(self.month_combo, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

        # === White Box (main container) ===
        self.box_widget = QtWidgets.QFrame()
        self.box_layout = QtWidgets.QVBoxLayout(self.box_widget)
        self.box_layout.setContentsMargins(20, 20, 20, 20)
        self.box_layout.setSpacing(15)

        # Add shadow effect
        shadow = QtWidgets.QGraphicsDropShadowEffect()
        shadow.setBlurRadius(20)
        shadow.setOffset(0, 4)
        shadow.setColor(QtGui.QColor(0, 0, 0, 80))
        self.box_widget.setGraphicsEffect(shadow)

        # === Input File Row ===
        self.input_layout = QtWidgets.QHBoxLayout()
        self.input_line = QtWidgets.QLineEdit()
        self.input_line.setReadOnly(True)

        # Auto-expand LineEdit
        self.input_layout.addWidget(self.input_line, stretch=1)

        self.input_btn = QtWidgets.QPushButton("📂 Select Input")
        self.input_btn.setObjectName("btnInput")
        self.input_btn.setFixedHeight(40)
        self.input_btn.setFixedWidth(120)  # consistent width
        self.input_layout.addWidget(self.input_btn)

        self.box_layout.addLayout(self.input_layout)

        # === Output File Row ===
        self.output_layout = QtWidgets.QHBoxLayout()
        self.output_line = QtWidgets.QLineEdit()
        self.output_line.setReadOnly(True)

        # Auto-expand LineEdit
        self.output_layout.addWidget(self.output_line, stretch=1)

        self.output_btn = QtWidgets.QPushButton("📂 Select Output")
        self.output_btn.setObjectName("btnOutput")
        self.output_btn.setFixedHeight(40)
        self.output_btn.setFixedWidth(120)  # same width as input
        self.output_layout.addWidget(self.output_btn)

        self.box_layout.addLayout(self.output_layout)

        # === Start Button (center) ===
        self.start_btn = QtWidgets.QPushButton("▶️ Start")
        self.start_btn.setFixedHeight(45)
        self.start_btn.setFixedWidth(150)
        self.box_layout.addWidget(self.start_btn, 
        alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

        # Add white box to main layout
        self.vlayout.addWidget(self.box_widget, 
        alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

        # === Footer ===
        self.footer = QtWidgets.QLabel("© 2025 PT Indo Tambangraya Megah Tbk – Automated SSO Processing")
        font_footer = QtGui.QFont("Arial", 9)
        self.footer.setFont(font_footer)
        self.footer.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.vlayout.addWidget(self.footer, alignment=QtCore.Qt.AlignmentFlag.AlignBottom)

        MainWindow.setCentralWidget(self.centralwidget)
        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "SSO Automation"))