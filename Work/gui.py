import sys
import os
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
    QPushButton, QLabel, QFileDialog, QMessageBox, QComboBox, 
    QSpinBox, QCheckBox, QGroupBox, QFormLayout, QTableView, QTabWidget,
    QHeaderView, QSplitter, QDoubleSpinBox, QScrollArea, QFrame, QGridLayout,
    QApplication, QDialog, QListWidget, QListWidgetItem, QLineEdit
)
from PySide6.QtCore import Qt, QAbstractTableModel, QThread, Signal
from PySide6.QtGui import QFont, QColor, QPixmap, QIcon

from utils import VALID_MEAN_GRADES, GradingError

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    return os.path.join(base_path, relative_path)

# ─── Lazy Import Helpers ────────────────────────────────────────
# Heavy modules (pandas, matplotlib, numpy, scipy, seaborn, fpdf)
# are only imported when first needed, making window appear in ~1-2s.
_pd = None
_excel_handler = None
_statistics_engine = None
_grading_engine = None
_graph_generator = None
_report_generator = None

def _get_pd():
    global _pd
    if _pd is None:
        import pandas as _pd_mod
        _pd = _pd_mod
    return _pd

def _get_excel_handler():
    global _excel_handler
    if _excel_handler is None:
        import excel_handler as _eh
        _excel_handler = _eh
    return _excel_handler

def _get_statistics_engine():
    global _statistics_engine
    if _statistics_engine is None:
        import statistics_engine as _se
        _statistics_engine = _se
    return _statistics_engine

def _get_grading_engine():
    global _grading_engine
    if _grading_engine is None:
        import grading_engine as _ge
        _grading_engine = _ge
    return _grading_engine

def _get_graph_generator():
    global _graph_generator
    if _graph_generator is None:
        import graph_generator as _gg
        _graph_generator = _gg
    return _graph_generator

def _get_report_generator():
    global _report_generator
    if _report_generator is None:
        from report_generator import generate_pdf_report as _gen
        _report_generator = _gen
    return _report_generator

class CalculationWorker(QThread):
    finished = Signal(object, object, bytes, bytes, str)  # df, stats, hist_bytes, bar_bytes, err_msg
    def __init__(self, df, mappings, weights, max_marks, fail_col, fail_threshold, total_fail_threshold, enable_hard_fail, boundaries, grace_limits):
        super().__init__()
        self.df = df
        self.mappings = mappings
        self.weights = weights
        self.max_marks = max_marks
        self.fail_col = fail_col
        self.fail_threshold = fail_threshold
        self.total_fail_threshold = total_fail_threshold
        self.enable_hard_fail = enable_hard_fail
        self.boundaries = boundaries
        self.grace_limits = grace_limits

    def run(self):
        try:
            import grading_engine
            # Dynamically handle backward compatibility if user mixes file versions
            import inspect
            sig = inspect.signature(grading_engine.run_grading)
            kwargs = {
                "df": self.df, "mappings": self.mappings, "weights": self.weights, 
                "max_marks": self.max_marks, "fail_col": self.fail_col, 
                "fail_threshold": self.fail_threshold, "total_fail_threshold": self.total_fail_threshold,
                "enable_hard_fail": self.enable_hard_fail, "boundaries": self.boundaries
            }
            if "grace_limits" in sig.parameters:
                kwargs["grace_limits"] = self.grace_limits
            result_df = grading_engine.run_grading(**kwargs)
            stats = _get_statistics_engine().calculate_stats(result_df, "Total_Score")
            hist_bytes = _get_graph_generator().generate_hist(result_df) or b""
            bar_bytes = _get_graph_generator().generate_bar(result_df) or b""
            self.finished.emit(result_df, stats, hist_bytes, bar_bytes, "")
        except Exception as e:
            self.finished.emit(None, None, b"", b"", str(e))

# ─── Theme Colors ───────────────────────────────────────────────
COLORS = {
    "bg":           "#141820",
    "card":         "#1c2230",
    "card_border":  "#2a3045",
    "primary":      "#00b4d8",
    "primary_dim":  "rgba(0,180,216,0.15)",
    "accent":       "#9b59b6",
    "text":         "#c8d0e0",
    "text_muted":   "#6b7a99",
    "destructive":  "#e74c3c",
    "success":      "#27ae60",
    "warning":      "#f39c12",
    "input_bg":     "#232a3a",
    "input_border": "#2e3a52",
}

# ─── Stylesheet ─────────────────────────────────────────────────
STYLESHEET = f"""
QMainWindow, QWidget#central {{
    background-color: {COLORS['bg']};
}}
QLabel {{
    color: {COLORS['text']};
    font-size: 13px;
}}
QLabel#sectionTitle {{
    font-size: 16px;
    font-weight: 600;
    color: {COLORS['primary']};
    padding-bottom: 4px;
}}
QLabel#headerTitle {{
    font-size: 17px;
    font-weight: 700;
    color: {COLORS['text']};
}}
QLabel#badge {{
    font-size: 11px;
    color: {COLORS['accent']};
    background-color: rgba(155,89,182,0.15);
    border: 1px solid rgba(155,89,182,0.3);
    border-radius: 10px;
    padding: 2px 10px;
}}
QLabel#muted {{
    color: {COLORS['text_muted']};
    font-size: 12px;
}}
QLabel#status {{
    color: {COLORS['text_muted']};
    font-style: italic;
    font-size: 12px;
}}
QLabel#valid {{
    color: {COLORS['success']};
    font-size: 13px;
    font-weight: 500;
}}
QLabel#invalid {{
    color: {COLORS['warning']};
    font-size: 13px;
    font-weight: 500;
}}
QLabel#failArrow {{
    color: {COLORS['destructive']};
    font-size: 14px;
    font-weight: 700;
}}
QGroupBox {{
    background-color: rgba(28,34,48,0.6);
    border: 1px solid {COLORS['card_border']};
    border-radius: 12px;
    padding: 18px;
    padding-top: 36px;
    margin-top: 8px;
    font-size: 0px;  
}}
QPushButton {{
    background-color: rgba(0,180,216,0.12);
    color: {COLORS['primary']};
    border: 1px solid rgba(0,180,216,0.3);
    border-radius: 8px;
    padding: 8px 18px;
    font-size: 13px;
    font-weight: 500;
}}
QPushButton:hover {{
    background-color: rgba(0,180,216,0.25);
    border-color: rgba(0,180,216,0.5);
}}
QPushButton:pressed {{
    background-color: rgba(0,180,216,0.35);
}}
QPushButton#action {{
    background-color: {COLORS['primary']};
    color: {COLORS['bg']};
    font-weight: 600;
    border: none;
    padding: 10px 28px;
    font-size: 14px;
}}
QPushButton#action:hover {{
    background-color: #00a0c0;
}}
QPushButton#export {{
    background-color: {COLORS['card']};
    color: {COLORS['text_muted']};
    border: 1px solid {COLORS['card_border']};
}}
QPushButton#export:hover {{
    color: {COLORS['text']};
    border-color: {COLORS['text_muted']};
}}
QComboBox {{
    background-color: {COLORS['input_bg']};
    color: {COLORS['text']};
    border: 1px solid {COLORS['input_border']};
    border-radius: 8px;
    padding: 6px 12px;
    font-size: 13px;
    min-height: 28px;
}}
QComboBox:hover {{
    border-color: rgba(0,180,216,0.4);
}}
QComboBox::drop-down {{
    border: none;
    width: 24px;
}}
QComboBox QAbstractItemView {{
    background-color: {COLORS['card']};
    color: {COLORS['text']};
    border: 1px solid {COLORS['card_border']};
    selection-background-color: rgba(0,180,216,0.2);
    selection-color: {COLORS['primary']};
    border-radius: 8px;
    padding: 4px;
}}
QSpinBox, QDoubleSpinBox {{
    background-color: {COLORS['input_bg']};
    color: {COLORS['text']};
    border: 1px solid {COLORS['input_border']};
    border-radius: 8px;
    padding: 6px 10px;
    font-size: 13px;
    font-family: 'Consolas', 'JetBrains Mono', monospace;
    min-height: 28px;
}}
QSpinBox:hover, QDoubleSpinBox:hover {{
    border-color: rgba(0,180,216,0.4);
}}
QCheckBox {{
    color: {COLORS['text']};
    font-size: 13px;
    spacing: 8px;
}}
QCheckBox::indicator {{
    width: 18px;
    height: 18px;
    border: 2px solid {COLORS['input_border']};
    border-radius: 4px;
    background-color: {COLORS['input_bg']};
}}
QCheckBox::indicator:checked {{
    background-color: {COLORS['primary']};
    border-color: {COLORS['primary']};
}}
QTabWidget::pane {{
    background-color: rgba(28,34,48,0.6);
    border: 1px solid {COLORS['card_border']};
    border-radius: 12px;
    padding: 12px;
}}
QTabBar::tab {{
    background-color: transparent;
    color: {COLORS['text_muted']};
    border: 1px solid transparent;
    border-radius: 8px;
    padding: 8px 18px;
    font-size: 13px;
    margin-right: 4px;
}}
QTabBar::tab:selected {{
    background-color: rgba(0,180,216,0.15);
    color: {COLORS['primary']};
    border-color: rgba(0,180,216,0.3);
}}
QTabBar::tab:hover {{
    background-color: rgba(0,180,216,0.08);
    color: {COLORS['text']};
}}
QTableView {{
    background-color: transparent;
    color: {COLORS['text']};
    border: none;
    gridline-color: {COLORS['card_border']};
    font-size: 12px;
    selection-background-color: rgba(0,180,216,0.15);
}}
QTableView::item {{
    padding: 6px 10px;
    border-bottom: 1px solid rgba(42,48,69,0.5);
}}
QHeaderView::section {{
    background-color: transparent;
    color: {COLORS['text_muted']};
    border: none;
    border-bottom: 1px solid {COLORS['card_border']};
    padding: 8px 10px;
    font-size: 11px;
    font-weight: 600;
    text-transform: uppercase;
}}
QTableCornerButton::section {{
     background-color: transparent;
}}
QScrollArea {{
    border: none;
    background: transparent;
}}
QScrollBar:vertical {{
    background: transparent;
    width: 8px;
    border-radius: 4px;
}}
QScrollBar::handle:vertical {{
    background: {COLORS['card_border']};
    border-radius: 4px;
    min-height: 30px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0;
}}
QFrame#header {{
    background-color: rgba(28,34,48,0.5);
    border-bottom: 1px solid {COLORS['card_border']};
    padding: 8px;
}}
"""

class PandasModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0] if self._data is not None else 0

    def columnCount(self, parent=None):
        return self._data.shape[1] if self._data is not None else 0

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid() or self._data is None:
            return None
        if role == Qt.DisplayRole:
            val = self._data.iloc[index.row(), index.column()]
            if _get_pd().isna(val): return ""
            if isinstance(val, float): return f"{val:.2f}"
            return str(val)
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole and self._data is not None:
            return str(self._data.columns[col])
        return None

class BestOfDialog(QDialog):
    def __init__(self, cols, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Create Best Of Column")
        self.setStyleSheet(STYLESHEET)
        layout = QVBoxLayout(self)
        
        layout.addWidget(QLabel("Select evaluation columns:"))
        self.list_widget = QListWidget()
        for c in cols:
            item = QListWidgetItem(str(c))
            item.setCheckState(Qt.Unchecked)
            self.list_widget.addItem(item)
        layout.addWidget(self.list_widget)
        
        layout.addWidget(QLabel("Take the top N scores:"))
        self.n_spin = QSpinBox()
        self.n_spin.setRange(1, 20)
        self.n_spin.setValue(1)
        layout.addWidget(self.n_spin)
        
        layout.addWidget(QLabel("New Column Name:"))
        self.name_input = QLineEdit()
        self.name_input.setText("Best_Scores")
        layout.addWidget(self.name_input)
        
        h = QHBoxLayout()
        btn_ok = QPushButton("Calculate & Add")
        btn_ok.setObjectName("action")
        btn_ok.clicked.connect(self.accept)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        h.addWidget(btn_cancel)
        h.addWidget(btn_ok)
        layout.addLayout(h)

    def get_data(self):
        checked_cols = []
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            if item.checkState() == Qt.Checked:
                checked_cols.append(item.text())
        return checked_cols, self.n_spin.value(), self.name_input.text().strip()

class GraceDialog(QDialog):
    def __init__(self, parent=None, current_limits=None):
        super().__init__(parent)
        self.setWindowTitle("Configure Grace Limits")
        self.resize(300, 400)
        self.current_limits = current_limits or {g: 0.0 for g in ["A", "A-", "B", "B-", "C", "C-", "D"]}
        self.spins = {}
        
        layout = QVBoxLayout(self)
        
        form_layout = QFormLayout()
        for g in ["A", "A-", "B", "B-", "C", "C-", "D"]:
            spin = QDoubleSpinBox()
            spin.setRange(0.0, 100.0)
            spin.setDecimals(1)
            spin.setValue(self.current_limits.get(g, 0.0))
            form_layout.addRow(f"To Grade {g}:", spin)
            self.spins[g] = spin
            
        layout.addLayout(form_layout)
        
        btn_layout = QHBoxLayout()
        btn_save = QPushButton("Save Config")
        btn_save.clicked.connect(self.accept)
        btn_cancel = QPushButton("Cancel")
        btn_cancel.clicked.connect(self.reject)
        
        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)
        
    def get_limits(self):
        return {g: spin.value() for g, spin in self.spins.items()}

class GlassGroup(QWidget):
    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(18, 18, 18, 18)
        self.main_layout.setSpacing(12)
        self.setStyleSheet(f"""
            GlassGroup {{
                background-color: rgba(28,34,48,0.6);
                border: 1px solid {COLORS['card_border']};
                border-radius: 12px;
            }}
        """)
        lbl = QLabel(title)
        lbl.setObjectName("sectionTitle")
        self.main_layout.addWidget(lbl)
    def addWidget(self, w):
        self.main_layout.addWidget(w)
    def addLayout(self, l):
        self.main_layout.addLayout(l)

class ErrorDialog(QDialog):
    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(18, 18, 18, 18)
        self.main_layout.setSpacing(12)
        self.setStyleSheet(f"""
            GlassGroup {{
                background-color: rgba(28,34,48,0.6);
                border: 1px solid {COLORS['card_border']};
                border-radius: 12px;
            }}
        """)
        lbl = QLabel(title)
        lbl.setObjectName("sectionTitle")
        self.main_layout.addWidget(lbl)
    def addWidget(self, w):
        self.main_layout.addWidget(w)
    def addLayout(self, l):
        self.main_layout.addLayout(l)
    def addStretch(self, s=1):
        self.main_layout.addStretch(s)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Relative Grading System")
        self.setWindowIcon(QIcon(resource_path("App Logo.ico")))
        self.resize(1400, 900)
        
        self.df = None
        self.result_df = None
        self.stats = None
        
        self.map_combos = {}
        self.weight_spins = {}
        self.max_spins = {}
        self.boundary_spins = {}
        
        self.setup_ui()
        self.setStyleSheet(STYLESHEET)
        
    def setup_ui(self):
        central = QWidget()
        central.setObjectName("central")
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)
        
        # ── Header ──
        header = QFrame()
        header.setObjectName("header")
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(20, 10, 20, 10)
        
        # ── Logo ──
        logo_lbl = QLabel()
        logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaled(44, 44, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            logo_lbl.setPixmap(pixmap)
            # Also set the application window/taskbar icon
            self.setWindowIcon(QIcon(logo_path))
        else:
            logo_lbl.setText("🎓")
            logo_lbl.setStyleSheet(f"background-color: {COLORS['primary_dim']}; border-radius: 8px; padding: 6px 8px; font-size: 18px;")
        logo_lbl.setFixedSize(50, 50)
        logo_lbl.setAlignment(Qt.AlignCenter)
        logo_lbl.setStyleSheet(f"background-color: {COLORS['primary_dim']}; border-radius: 10px; padding: 3px;")
        header_layout.addWidget(logo_lbl)

        title_col = QVBoxLayout()
        title_col.setSpacing(2)
        title = QLabel("Relative Grading System")
        title.setObjectName("headerTitle")
        subtitle = QLabel("Advanced Enterprise Edition")
        subtitle.setObjectName("muted")
        title_col.addWidget(title)
        title_col.addWidget(subtitle)
        header_layout.addLayout(title_col)
        
        badge = QLabel("v2.0  Enterprise")
        badge.setObjectName("badge")
        header_layout.addWidget(badge)
        header_layout.addStretch()
        
        self.btn_exp_excel = QPushButton("💾 Export Excel")
        self.btn_exp_excel.setObjectName("export")
        self.btn_exp_excel.clicked.connect(self.on_export_excel)
        
        self.btn_exp_pdf = QPushButton("📄 Export PDF")
        self.btn_exp_pdf.setObjectName("export")
        self.btn_exp_pdf.clicked.connect(self.on_export_pdf)
        
        self.btn_calculate = QPushButton("⚡ Calculate Grades")
        self.btn_calculate.setObjectName("action")
        self.btn_calculate.clicked.connect(self.on_calculate)
        
        header_layout.addWidget(self.btn_exp_excel)
        header_layout.addWidget(self.btn_exp_pdf)
        header_layout.addWidget(self.btn_calculate)
        
        root_layout.addWidget(header)
        
        # ── Body ──
        body = QWidget()
        body_layout = QHBoxLayout(body)
        body_layout.setContentsMargins(20, 20, 20, 20)
        
        self.main_splitter = QSplitter(Qt.Horizontal)
        self.main_splitter.setChildrenCollapsible(False)
        self.main_splitter.setHandleWidth(8)
        
        # Left Panel (Scrollable Config)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll.setMinimumWidth(350)
        
        scroll_content = QWidget()
        self.left_layout = QVBoxLayout(scroll_content)
        self.left_layout.setSpacing(16)
        
        self.build_data_source()
        self.build_eval_schema()
        self.build_grading_params()
        self.build_dist_shifts()
        
        # Connect signal ONLY after everything is built
        self.grade_combo.currentTextChanged.connect(self.on_mean_grade_changed)
        # Sync σ distribution with default mean grade selection (A-)
        self.on_mean_grade_changed(self.grade_combo.currentText())
        
        self.left_layout.addStretch()
        scroll.setWidget(scroll_content)
        self.main_splitter.addWidget(scroll)
        
        # Right Panel (Tabs)
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0,0,0,0)
        
        self.tabs = QTabWidget()
        
        # Data View Tab
        self.tab_data = QWidget()
        data_layout = QVBoxLayout(self.tab_data)
        data_layout.setContentsMargins(0, 0, 0, 0)
        self.table = QTableView()
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        data_layout.addWidget(self.table)
        self.tabs.addTab(self.tab_data, "Data & Results")
        
        # Analytics Tab
        self.tab_analytics = QScrollArea()
        self.tab_analytics.setWidgetResizable(True)
        analytics_content = QWidget()
        self.analytics_layout = QVBoxLayout(analytics_content)
        self.analytics_layout.setContentsMargins(20, 20, 20, 20)
        self.analytics_layout.setSpacing(25)
        
        self.lbl_stats = QLabel("Run calculation to see analytics.")
        self.lbl_stats.setObjectName("muted")
        self.lbl_stats.setAlignment(Qt.AlignCenter)
        self.lbl_stats.setWordWrap(True)
        self.analytics_layout.addWidget(self.lbl_stats)
        
        self.lbl_hist = QLabel()
        self.lbl_hist.setAlignment(Qt.AlignCenter)
        self.analytics_layout.addWidget(self.lbl_hist)
        
        self.lbl_bar = QLabel()
        self.lbl_bar.setAlignment(Qt.AlignCenter)
        self.analytics_layout.addWidget(self.lbl_bar)
        
        self.tab_analytics.setWidget(analytics_content)
        self.tabs.addTab(self.tab_analytics, "Visual Analytics")
        
        right_layout.addWidget(self.tabs)
        self.main_splitter.addWidget(right_panel)
        
        # Set initial splitter sizes (left panel ~40%, right tab ~60%)
        self.main_splitter.setSizes([450, 750])
        
        body_layout.addWidget(self.main_splitter)
        
        root_layout.addWidget(body, 1)
        self.generate_mapping_rows()

    def build_data_source(self):
        self.data_source = GlassGroup("1. Data Source")
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)
        self.btn_template = QPushButton("📄 Generate Template")
        self.btn_upload = QPushButton("📤 Upload Excel Data")
        btn_row.addWidget(self.btn_template)
        btn_row.addWidget(self.btn_upload)
        self.data_source.addLayout(btn_row)
        
        self.btn_best_of = QPushButton("✨ Create 'Best Of' Column")
        self.btn_best_of.setEnabled(False)
        self.btn_best_of.clicked.connect(self.on_best_of)
        self.data_source.addWidget(self.btn_best_of)

        self.status_label = QLabel("Ready to load data...")
        self.status_label.setObjectName("status")
        self.data_source.addWidget(self.status_label)
        
        self.btn_template.clicked.connect(self.on_generate_template)
        self.btn_upload.clicked.connect(self.on_upload_excel)
        self.left_layout.addWidget(self.data_source)

    def build_eval_schema(self):
        self.eval_schema = GlassGroup("2. Evaluation Schema")
        top_row = QHBoxLayout()
        top_row.setSpacing(10)
        top_row.addWidget(QLabel("Number of Evaluation Categories:"))
        self.num_spin = QSpinBox()
        self.num_spin.setRange(1, 10)
        self.num_spin.setValue(5)
        self.num_spin.setFixedWidth(70)
        top_row.addWidget(self.num_spin)
        self.btn_generate = QPushButton("Generate Mapping Rows")
        self.btn_generate.clicked.connect(self.generate_mapping_rows)
        top_row.addWidget(self.btn_generate)
        top_row.addStretch()
        self.eval_schema.addLayout(top_row)
        
        header = QHBoxLayout()
        header.setSpacing(10)
        h1 = QLabel("")
        h1.setFixedWidth(70)
        header.addWidget(h1)
        h2 = QLabel("EXCEL COLUMN")
        h2.setObjectName("muted")
        header.addWidget(h2, 1)
        h3 = QLabel("WEIGHT %")
        h3.setObjectName("muted")
        h3.setFixedWidth(100)
        header.addWidget(h3)
        h4 = QLabel("MAX MARKS")
        h4.setObjectName("muted")
        h4.setFixedWidth(80)
        header.addWidget(h4)
        self.eval_schema.addLayout(header)
        
        self.rows_container = QVBoxLayout()
        self.rows_container.setSpacing(6)
        self.eval_schema.addLayout(self.rows_container)
        
        self.weight_label = QLabel("Total weight: 0%. It must equal 100%.")
        self.weight_label.setObjectName("invalid")
        self.eval_schema.addWidget(self.weight_label)
        self.left_layout.addWidget(self.eval_schema)

    def build_grading_params(self):
        self.grading_params = GlassGroup("3. Grading Parameters")
        row1 = QHBoxLayout()
        row1.setSpacing(10)
        row1.addWidget(QLabel("Target Class Mean Grade:"))
        self.grade_combo = QComboBox()
        self.grade_combo.addItems(VALID_MEAN_GRADES)
        self.grade_combo.setCurrentText("A-")
        row1.addWidget(self.grade_combo)
        
        self.grace_limits_dict = {g: 0.0 for g in ["A", "A-", "B", "B-", "C", "C-", "D"]}
        self.btn_grace = QPushButton("⚙ Configure Grace Limits")
        self.btn_grace.clicked.connect(self.open_grace_dialog)
        row1.addWidget(self.btn_grace)
        
        row1.addStretch()
        self.grading_params.addLayout(row1)
        
        self.hard_fail_check = QCheckBox("Enable Hard Fail Policy")
        self.hard_fail_check.setChecked(True) # Enabled by default as requested
        self.grading_params.addWidget(self.hard_fail_check)
        
        self.fail_widget = QWidget()
        fail_layout = QVBoxLayout(self.fail_widget)
        fail_layout.setContentsMargins(24, 0, 0, 0)
        fail_layout.setSpacing(8)
        
        row1 = QHBoxLayout()
        row1.addWidget(QLabel("If"))
        self.fail_col = QComboBox()
        row1.addWidget(self.fail_col, 1)
        row1.addWidget(QLabel("<"))
        self.fail_threshold = QSpinBox()
        self.fail_threshold.setRange(0, 100)
        self.fail_threshold.setFixedWidth(70)
        row1.addWidget(self.fail_threshold)
        arrow1 = QLabel("→ F")
        arrow1.setObjectName("failArrow")
        row1.addWidget(arrow1)
        row1.addStretch()
        fail_layout.addLayout(row1)

        row2 = QHBoxLayout()
        row2.addWidget(QLabel("OR If Total Score <"))
        self.total_fail_threshold = QSpinBox()
        self.total_fail_threshold.setRange(0, 100)
        self.total_fail_threshold.setFixedWidth(70)
        row2.addWidget(self.total_fail_threshold)
        arrow2 = QLabel("→ F")
        arrow2.setObjectName("failArrow")
        row2.addWidget(arrow2)
        row2.addStretch()
        fail_layout.addLayout(row2)
        self.fail_widget.setVisible(self.hard_fail_check.isChecked())
        self.grading_params.addWidget(self.fail_widget)
        
        self.hard_fail_check.toggled.connect(self.fail_widget.setVisible)
        self.left_layout.addWidget(self.grading_params)

    def build_dist_shifts(self):
        self.dist_shifts = GlassGroup("4. Distribution Shifts (\u03c3)")
        self.GRADE_ORDER = ["A", "A-", "B", "B-", "C", "C-", "D"]
        self.DEFAULT_SIGMA = {
            "A": 1.50, "A-": 1.00, "B": 0.50, "B-": 0.00,
            "C": -0.50, "C-": -1.00, "D": -1.50
        }
        grid = QGridLayout()
        grid.setSpacing(8)

        for i, grade in enumerate(self.GRADE_ORDER):
            val = self.DEFAULT_SIGMA[grade]
            row = i % 4
            col_offset = (i // 4) * 2
            lbl = QLabel(f"\u03c3 Cutoff {grade}:")
            lbl.setObjectName("muted")
            grid.addWidget(lbl, row, col_offset)
            spin = QDoubleSpinBox()
            spin.setRange(-10.0, 10.0)
            spin.setSingleStep(0.1)
            spin.setDecimals(2)
            spin.setValue(val)
            spin.setReadOnly(False) 
            grid.addWidget(spin, row, col_offset + 1)
            self.boundary_spins[grade] = spin

        hint = QLabel("\u2139 Selecting a mean grade auto-sets its \u03c3 to 0 and scales the rest.")
        hint.setObjectName("muted")
        hint.setWordWrap(True)
        self.dist_shifts.addLayout(grid)
        self.dist_shifts.addWidget(hint)
        self.left_layout.addWidget(self.dist_shifts)

    def on_mean_grade_changed(self, selected_grade: str):
        if selected_grade not in self.GRADE_ORDER:
            return

        anchor_idx = self.GRADE_ORDER.index(selected_grade)
        STEP = 0.50  
        new_values = {}
        for i, grade in enumerate(self.GRADE_ORDER):
            distance = anchor_idx - i  
            new_values[grade] = round(distance * STEP, 2)

        for grade, spin in self.boundary_spins.items():
            spin.blockSignals(True)
            spin.setValue(new_values.get(grade, spin.value()))
            spin.blockSignals(False)

    def generate_mapping_rows(self):
        while self.rows_container.count():
            item = self.rows_container.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
                
        self.map_combos.clear()
        self.weight_spins.clear()
        self.max_spins.clear()
        
        count = self.num_spin.value()
        cols = ["[None]"]
        if self.df is not None:
             cols += list(self.df.columns)
             
        for i in range(1, count + 1):
            comp_name = f"Evaluation {i}"
            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(10)
            
            lbl = QLabel(f"Eval {i}:")
            lbl.setObjectName("muted")
            lbl.setFixedWidth(70)
            row_layout.addWidget(lbl)
            
            cb = QComboBox()
            cb.addItems(cols)
            cb.currentTextChanged.connect(self.update_fail_col_dropdown)
            row_layout.addWidget(cb, 1)
            
            sb = QSpinBox()
            sb.setRange(0, 100)
            sb.setSuffix(" %")
            sb.setFixedWidth(100)
            sb.valueChanged.connect(self.check_weight_sum)
            row_layout.addWidget(sb)
            
            mb = QDoubleSpinBox()
            mb.setRange(1, 1000)
            mb.setValue(100)
            mb.setDecimals(1)
            mb.setFixedWidth(80)
            row_layout.addWidget(mb)
            
            self.map_combos[comp_name] = cb
            self.weight_spins[comp_name] = sb
            self.max_spins[comp_name] = mb
            
            self.rows_container.addWidget(row_widget)
            
        self.check_weight_sum()
        self.update_fail_col_dropdown()
        
    def check_weight_sum(self):
        total = sum(sb.value() for sb in self.weight_spins.values())
        if total == 100:
            self.weight_label.setText(f"✅ Total weight: {total}%. Valid!")
            self.weight_label.setObjectName("valid")
        else:
            self.weight_label.setText(f"⚠ Total weight: {total}%. It must equal 100%.")
            self.weight_label.setObjectName("invalid")
        self.weight_label.setStyleSheet(self.weight_label.styleSheet())
            
    def update_fail_col_dropdown(self):
        current = self.fail_col.currentText()
        self.fail_col.clear()
        self.fail_col.addItem("[None]")
        if self.df is not None:
            for cb in self.map_combos.values():
                mapped_val = cb.currentText()
                if mapped_val != "[None]":
                    self.fail_col.addItem(mapped_val)
        index = self.fail_col.findText(current)
        if index >= 0:
            self.fail_col.setCurrentIndex(index)

    def populate_columns(self):
        if self.df is not None:
            cols = list(self.df.columns)
            for cb in self.map_combos.values():
                cb.clear()
                cb.addItem("[None]")
                cb.addItems(cols)
            self.update_fail_col_dropdown()

    def on_generate_template(self):
        count = self.num_spin.value()
        path, _ = QFileDialog.getSaveFileName(self, "Save Template", "template.xlsx", "Excel Files (*.xlsx)")
        if path:
            import excel_handler
            cols = ["Roll_No", "Name"] + [f"Eval_{i}" for i in range(1, count+1)]
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "Marks"
            ws.append(cols)
            wb.save(path)
            self.status_label.setText(f"✅ Template generated: {os.path.basename(path)}")
            QMessageBox.information(self, "Success", f"Template with {count} evaluation columns saved.")

    def on_upload_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Excel", "", "Excel Files (*.xlsx *.xls)")
        if path:
            try:
                self.df = _get_excel_handler().read_excel(path)
                self.status_label.setText(f"✅ Loaded: {os.path.basename(path)}")
                self.status_label.setObjectName("valid")
                self.status_label.setStyleSheet(self.status_label.styleSheet())
                
                self.btn_best_of.setEnabled(True)
                self.populate_columns()
                model = PandasModel(self.df)
                self.table.setModel(model)
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def on_best_of(self):
        if self.df is None: return
        numeric_cols = [c for c in self.df.columns if _get_pd().api.types.is_numeric_dtype(self.df[c])]
        if not numeric_cols:
            QMessageBox.warning(self, "Error", "No numeric columns available in dataset.")
            return
            
        dlg = BestOfDialog(numeric_cols, self)
        if dlg.exec() == QDialog.Accepted:
            cols, n, new_name = dlg.get_data()
            if not new_name:
                QMessageBox.warning(self, "Error", "Invalid column name.")
                return
            if len(cols) < n:
                QMessageBox.warning(self, "Error", f"You chose to take the top {n} scores but only selected {len(cols)} columns.")
                return
                
            try:
                # Sum top n columns row-wise
                self.df[new_name] = self.df[cols].apply(lambda row: row.nlargest(n).sum(), axis=1)
                self.populate_columns()
                self.table.setModel(PandasModel(self.df))
                QMessageBox.information(self, "Success", f"Column '{new_name}' added successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to compute: {e}")

    def open_grace_dialog(self):
        dlg = GraceDialog(self, self.grace_limits_dict)
        if dlg.exec():
            self.grace_limits_dict = dlg.get_limits()
            active_count = sum(1 for v in self.grace_limits_dict.values() if v > 0)
            if active_count > 0:
                self.btn_grace.setText(f"⚙ Grace Configured ({active_count} active)")
                self.btn_grace.setStyleSheet("background: rgba(46, 204, 113, 0.2); color: #2ecc71;")
            else:
                self.btn_grace.setText("⚙ Configure Grace Limits")
                self.btn_grace.setStyleSheet("")

    def on_calculate(self):
        if self.df is None:
            QMessageBox.warning(self, "Warning", "Please upload Excel data first.")
            return
            
        total_weight = sum(sb.value() for sb in self.weight_spins.values())
        if total_weight != 100:
            QMessageBox.warning(self, "Invalid Weightages", f"Total evaluation weight is {total_weight}%. It must be exactly 100%.")
            return
            
        mappings = {comp: cb.currentText() for comp, cb in self.map_combos.items() if cb.currentText() != "[None]"}
        weights = {comp: sb.value() for comp, sb in self.weight_spins.items()}
        max_marks = {comp: mb.value() for comp, mb in self.max_spins.items()}
        fail_col = self.fail_col.currentText() if self.fail_col.currentText() != "[None]" else None
        fail_threshold = self.fail_threshold.value()
        total_fail_threshold = self.total_fail_threshold.value()
        enable_hard_fail = self.hard_fail_check.isChecked()
        boundaries = {g: spin.value() for g, spin in self.boundary_spins.items()}
        grace_limits = self.grace_limits_dict
        
        self.btn_calculate.setEnabled(False)
        self.btn_calculate.setText("⚙ Calculating...")
        self.worker = CalculationWorker(
            self.df, mappings, weights, max_marks, 
            fail_col, fail_threshold, total_fail_threshold, enable_hard_fail, boundaries, grace_limits
        )
        self.worker.finished.connect(self.on_calculate_finished)
        self.worker.start()

    def on_calculate_finished(self, result_df, stats, hist_bytes, bar_bytes, err_msg):
        self.btn_calculate.setEnabled(True)
        self.btn_calculate.setText("⚡ Calculate Grades")

        if err_msg:
            QMessageBox.critical(self, "Calculation Error", err_msg)
            return

        self.result_df = result_df
        self.stats = stats
        
        # Update Table
        self.table.setModel(PandasModel(self.result_df))
        
        # Update Analytics Stats
        stats_text = (f"Class Size: {self.stats.get('count',0)}  |  "
                      f"Mean Score: {self.stats.get('mean',0):.2f}  |  "
                      f"Std Dev (σ): {self.stats.get('std',0):.2f}")
        self.lbl_stats.setText(stats_text)
        self.lbl_stats.setObjectName("text")
        self.lbl_stats.setStyleSheet(self.lbl_stats.styleSheet())
        
        # Update Graphs
        if hist_bytes:
            pixmap = QPixmap()
            pixmap.loadFromData(hist_bytes)
            self.lbl_hist.setPixmap(pixmap)
            
        if bar_bytes:
            pixmap = QPixmap()
            pixmap.loadFromData(bar_bytes)
            self.lbl_bar.setPixmap(pixmap)
            
        QMessageBox.information(self, "Success", "Grades calculated successfully! Check the Data and Analytics tabs.")

    def on_export_excel(self):
        if self.result_df is None:
            QMessageBox.warning(self, "Warning", "Please calculate grades first.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save Excel", "grades_results.xlsx", "Excel Files (*.xlsx)")
        if path:
            try:
                self.result_df.to_excel(path, index=False)
                QMessageBox.information(self, "Success", f"Results exported to {path}")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", str(e))

    def on_export_pdf(self):
        if self.result_df is None or self.stats is None:
            QMessageBox.warning(self, "Warning", "Please calculate grades first.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "grades_report.pdf", "PDF Files (*.pdf)")
        if path:
            try:
                _get_report_generator()(self.result_df, self.stats, path)
                QMessageBox.information(self, "Success", f"Report exported to {path}")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", str(e))
