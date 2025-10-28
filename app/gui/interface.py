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
    Class to define the main application UI.
    """

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(700, 600)

        # === Load stylesheet (QSS) ===
        qss_path = resource_path("style/style.qss")
        if os.path.exists(qss_path):
            try:
                with open(qss_path, "r", encoding="utf-8") as f:
                    MainWindow.setStyleSheet(f.read())
            except Exception:
                # jika gagal baca stylesheet, lanjut tanpa crash
                pass

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
        self.logo.setObjectName("label_logo")
        self.vlayout.addWidget(self.logo, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

        # === Company Name ===
        self.label_company = QtWidgets.QLabel("PT Indo Tambangraya Megah Tbk")
        self.label_company.setFont(QtGui.QFont("Arial", 14, QtGui.QFont.Weight.Bold))
        self.label_company.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label_company.setObjectName("label_company")
        self.vlayout.addWidget(self.label_company)

        # === Application Title ===
        self.label_title = QtWidgets.QLabel("SSO Results Automation System")
        self.label_title.setFont(QtGui.QFont("Arial", 20, QtGui.QFont.Weight.Bold))
        self.label_title.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label_title.setObjectName("label_title")
        self.vlayout.addWidget(self.label_title)

        # === Application Subtitle ===
        self.label_sub = QtWidgets.QLabel("Please select the input and output file to begin automated processing.")
        self.label_sub.setFont(QtGui.QFont("Arial", 10))
        self.label_sub.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.label_sub.setObjectName("label_sub")
        self.vlayout.addWidget(self.label_sub)

        # --- Month data (reuse list) ---
        months: List[str] = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ]

        # === This Month (label + combobox) ===
        this_month_group = QtWidgets.QHBoxLayout()
        this_month_group.setSpacing(12)

        self.this_month_label = QtWidgets.QLabel("This Month :")
        self.this_month_label.setFont(QtGui.QFont("Arial", 10))
        self.this_month_label.setObjectName("label_this_month")
        # keep label left-aligned in the horizontal group
        this_month_group.addWidget(self.this_month_label, alignment=QtCore.Qt.AlignmentFlag.AlignVCenter)

        # ComboBox This Month
        self.month_combo = QtWidgets.QComboBox()
        self.month_combo.setFixedSize(120, 35)
        self.month_combo.addItems(months)

        # Buat delegate agar teks di tengah
        delegate = QtWidgets.QStyledItemDelegate(self.month_combo)
        self.month_combo.setItemDelegate(delegate)

        for i in range(self.month_combo.count()):
            self.month_combo.setItemData(i, QtCore.Qt.AlignmentFlag.AlignCenter, QtCore.Qt.ItemDataRole.TextAlignmentRole)
        self.month_combo.setObjectName("combo_next_month")
        this_month_group.addWidget(self.month_combo, alignment=QtCore.Qt.AlignmentFlag.AlignVCenter)

        # Tambahkan ke layout utama
        # add stretch so group stays compact and centered
        wrapper = QtWidgets.QHBoxLayout()
        wrapper.addStretch(1)
        wrapper.addLayout(this_month_group)
        wrapper.addStretch(1)
        self.vlayout.addLayout(wrapper)

        # === Next Month (label + combobox + checkbox) ===
        next_month_group = QtWidgets.QHBoxLayout()
        next_month_group.setSpacing(12)

        self.next_month_label = QtWidgets.QLabel("Next Month :")
        self.next_month_label.setFont(QtGui.QFont("Arial", 10))
        self.next_month_label.setObjectName("label_next_month")
        # keep label left-aligned in the horizontal group
        next_month_group.addWidget(self.next_month_label, alignment=QtCore.Qt.AlignmentFlag.AlignVCenter)

        # ComboBox Next Month
        self.month_combo2 = QtWidgets.QComboBox()
        self.month_combo2.setFixedSize(120, 35)
        self.month_combo2.addItems(months)

        # Buat delegate agar teks di tengah
        delegate2 = QtWidgets.QStyledItemDelegate(self.month_combo2)
        self.month_combo2.setItemDelegate(delegate2)

        for i in range(self.month_combo2.count()):
            self.month_combo2.setItemData(i, QtCore.Qt.AlignmentFlag.AlignCenter, QtCore.Qt.ItemDataRole.TextAlignmentRole)
        self.month_combo2.setObjectName("combo_next_month2")
        next_month_group.addWidget(self.month_combo2, alignment=QtCore.Qt.AlignmentFlag.AlignVCenter)

        # Checkbox should appear to the RIGHT of the Next Month combobox
        self.checkBox_enableNextMonth = QtWidgets.QCheckBox()
        self.checkBox_enableNextMonth.setToolTip("Enable processing for next month")
        self.checkBox_enableNextMonth.setObjectName("checkBox_enableNextMonth")

        # Default: disabled combobox until checkbox enabled
        self.month_combo2.setEnabled(False)
        # connect toggle to enable/disable combobox
        self.checkBox_enableNextMonth.toggled.connect(self.month_combo2.setEnabled)
        next_month_group.addWidget(self.checkBox_enableNextMonth, alignment=QtCore.Qt.AlignmentFlag.AlignVCenter)

        # add stretch so group stays compact and centered
        wrapper = QtWidgets.QHBoxLayout()
        wrapper.addStretch(1)
        wrapper.addLayout(next_month_group)
        wrapper.addStretch(1)
        self.vlayout.addLayout(wrapper)

        # === White Box (main container) ===
        self.box_widget = QtWidgets.QFrame()
        self.box_widget.setObjectName("box_widget")
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
        self.input_line.setObjectName("input_line")
        # Auto-expand LineEdit
        self.input_layout.addWidget(self.input_line, stretch=1)

        self.input_btn = QtWidgets.QPushButton("📂 Select Input")
        self.input_btn.setObjectName("btnInput")
        self.input_btn.setFixedSize(120, 40)
        # connect signals to your handlers later, e.g. self.input_btn.clicked.connect(self.on_select_input)
        self.input_layout.addWidget(self.input_btn)

        self.box_layout.addLayout(self.input_layout)

        # === Output File Row ===
        self.output_layout = QtWidgets.QHBoxLayout()
        self.output_line = QtWidgets.QLineEdit()
        self.output_line.setReadOnly(True)
        self.output_line.setObjectName("output_line")
        # Auto-expand LineEdit
        self.output_layout.addWidget(self.output_line, stretch=1)

        self.output_btn = QtWidgets.QPushButton("📂 Select Output")
        self.output_btn.setObjectName("btnOutput")
        self.output_btn.setFixedSize(120, 40)
        self.output_layout.addWidget(self.output_btn)

        self.box_layout.addLayout(self.output_layout)

        # === Start Button (center) ===
        self.start_btn = QtWidgets.QPushButton("▶️ Start")
        self.start_btn.setObjectName("btnStart")
        self.start_btn.setFixedSize(150, 45)
        self.box_layout.addWidget(self.start_btn, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

        # Add white box to main layout
        self.vlayout.addWidget(self.box_widget, alignment=QtCore.Qt.AlignmentFlag.AlignCenter)

        # === Footer ===
        self.footer = QtWidgets.QLabel("© 2025 PT Indo Tambangraya Megah Tbk – Automated SSO Processing")
        self.footer.setFont(QtGui.QFont("Arial", 9))
        self.footer.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.footer.setObjectName("label_footer")
        self.vlayout.addWidget(self.footer, alignment=QtCore.Qt.AlignmentFlag.AlignBottom)

        MainWindow.setCentralWidget(self.centralwidget)

        # Set default ke bulan saat ini
        import datetime
        current_month = datetime.datetime.now().month
        self.month_combo.setCurrentIndex(current_month - 1)
        self.month_combo2.setCurrentIndex(current_month - 1)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "SSO Automation"))