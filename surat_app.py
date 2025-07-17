import sys
import os

# Hilangkan console di Windows
if sys.platform == 'win32':
    import ctypes
    if not 'DEBUG' in os.environ:  # Biarkan console muncul saat debug
        ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PyQt5.QtWidgets import (QApplication, QMainWindow, QTabWidget, 
                             QVBoxLayout, QWidget, QPushButton, 
                             QLineEdit, QTextEdit, QLabel, 
                             QTableWidget, QTableWidgetItem, QDesktopWidget,
                             QHBoxLayout, QFrame, QScrollArea, QGroupBox,
                             QGridLayout, QSpacerItem, QSizePolicy, QMessageBox,
                             QDateEdit, QHeaderView)
from PyQt5.QtCore import Qt, QPropertyAnimation, QRect, QEasingCurve, QDate
from PyQt5.QtGui import QFont, QIcon, QPalette, QColor, QPixmap, QPainter, QBrush
from datetime import datetime

# Koneksi ke Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)
sheet = client.open("Database Surat").sheet1  # Ganti dengan nama sheet Anda

class ModernButton(QPushButton):
    def __init__(self, text, color="#4CAF50", hover_color="#45a049", text_color="white"):
        super().__init__(text)
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {color};
                color: {text_color};
                border: none;
                padding: 12px 24px;
                border-radius: 8px;
                font-size: 14px;
                font-weight: 600;
                text-align: center;
                min-height: 20px;
            }}
            QPushButton:hover {{
                background-color: {hover_color};
                transform: translateY(-2px);
            }}
            QPushButton:pressed {{
                background-color: {color};
                transform: translateY(0px);
            }}
        """)

class ModernLineEdit(QLineEdit):
    def __init__(self, placeholder=""):
        super().__init__()
        self.setPlaceholderText(placeholder)
        self.setStyleSheet("""
            QLineEdit {
                padding: 12px 16px;
                border: 2px solid #e0e0e0;
                border-radius: 8px;
                font-size: 14px;
                background-color: white;
                color: #333;
            }
            QLineEdit:focus {
                border-color: #4CAF50;
                background-color: #fafafa;
            }
            QLineEdit:hover {
                border-color: #c0c0c0;
            }
        """)
        # Auto-resize berdasarkan text
        self.textChanged.connect(self.adjust_width)
        self.setMinimumWidth(200)  # Lebar minimum
        
    def adjust_width(self):
        # Hitung lebar text dengan padding
        metrics = self.fontMetrics()
        text_width = metrics.width(self.text())
        new_width = max(200, text_width + 50)  # Minimum 200px, +50 untuk padding
        self.setMinimumWidth(new_width)

class ModernTextEdit(QTextEdit):
    def __init__(self, placeholder=""):
        super().__init__()
        self.setPlaceholderText(placeholder)
        self.setStyleSheet("""
            QTextEdit {
                padding: 12px 16px;
                border: 2px solid #e0e0e0;
                border-radius: 8px;
                font-size: 14px;
                background-color: white;
                color: #333;
                min-height: 100px;
            }
            QTextEdit:focus {
                border-color: #4CAF50;
                background-color: #fafafa;
            }
            QTextEdit:hover {
                border-color: #c0c0c0;
            }
        """)

class ModernLabel(QLabel):
    def __init__(self, text, size=14, weight="normal", color="#333"):
        super().__init__(text)
        font_weight = "400" if weight == "normal" else "600" if weight == "bold" else "300"
        self.setStyleSheet(f"""
            QLabel {{
                color: {color};
                font-size: {size}px;
                font-weight: {font_weight};
                margin: 4px 0;
            }}
        """)

class ModernDateEdit(QDateEdit):
    def __init__(self):
        super().__init__()
        # Set tanggal default ke hari ini
        self.setDate(QDate.currentDate())
        self.setCalendarPopup(True)
        self.setDisplayFormat("dd/MM/yyyy")
        self.setStyleSheet("""
            QDateEdit {
                padding: 12px 16px;
                border: 2px solid #e0e0e0;
                border-radius: 8px;
                font-size: 14px;
                background-color: white;
                color: #333;
                min-height: 20px;
            }
            QDateEdit:focus {
                border-color: #4CAF50;
                background-color: #fafafa;
            }
            QDateEdit:hover {
                border-color: #c0c0c0;
            }
            QDateEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 30px;
                border-left: 1px solid #c0c0c0;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
                background-color: #f0f0f0;
            }
            QDateEdit::drop-down:hover {
                background-color: #e0e0e0;
            }
            QDateEdit::down-arrow {
                width: 12px;
                height: 12px;
            }
        """)

class ModernGroupBox(QGroupBox):
    def __init__(self, title=""):
        super().__init__(title)
        self.setStyleSheet("""
            QGroupBox {
                font-size: 16px;
                font-weight: 600;
                color: #333;
                padding: 20px;
                margin-top: 20px;
                border: 2px solid #e0e0e0;
                border-radius: 12px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 20px;
                padding: 0 10px 0 10px;
                color: #4CAF50;
            }
        """)

class ModernMessageBox:
    @staticmethod
    def show_success(parent, title, message):
        msg_box = QMessageBox(parent)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QMessageBox QLabel {
                color: #333;
                font-size: 14px;
                padding: 10px;
            }
            QMessageBox QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 600;
                min-width: 80px;
            }
            QMessageBox QPushButton:hover {
                background-color: #45a049;
            }
            QMessageBox QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        msg_box.exec_()
    
    @staticmethod
    def show_error(parent, title, message):
        msg_box = QMessageBox(parent)
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QMessageBox QLabel {
                color: #333;
                font-size: 14px;
                padding: 10px;
            }
            QMessageBox QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 600;
                min-width: 80px;
            }
            QMessageBox QPushButton:hover {
                background-color: #da190b;
            }
            QMessageBox QPushButton:pressed {
                background-color: #c62828;
            }
        """)
        msg_box.exec_()
    
    @staticmethod
    def show_warning(parent, title, message):
        msg_box = QMessageBox(parent)
        msg_box.setIcon(QMessageBox.Warning)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setStandardButtons(QMessageBox.Ok)
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: white;
                color: #333;
                font-size: 14px;
            }
            QMessageBox QLabel {
                color: #333;
                font-size: 14px;
                padding: 10px;
            }
            QMessageBox QPushButton {
                background-color: #FF9800;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 6px;
                font-size: 14px;
                font-weight: 600;
                min-width: 80px;
            }
            QMessageBox QPushButton:hover {
                background-color: #F57C00;
            }
            QMessageBox QPushButton:pressed {
                background-color: #E65100;
            }
        """)
        msg_box.exec_()

class SuratApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("📨 Aplikasi Persuratan SLBN Tanah Grogot")
        
        # Set aplikasi theme
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QTabWidget::pane {
                border: none;
                background-color: #f5f5f5;
            }
            QTabWidget::tab-bar {
                alignment: center;
            }
            QTabBar::tab {
                background-color: #e0e0e0;
                color: #666;
                padding: 16px 32px;
                margin-right: 4px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                font-size: 14px;
                font-weight: 600;
                min-width: 120px;
            }
            QTabBar::tab:selected {
                background-color: #4CAF50;
                color: white;
            }
            QTabBar::tab:hover:!selected {
                background-color: #c0c0c0;
            }
            QTableWidget {
                gridline-color: #e0e0e0;
                background-color: white;
                alternate-background-color: #f8f8f8;
                border: 1px solid #e0e0e0;
                border-radius: 8px;
                font-size: 13px;
                word-wrap: break-word;
            }
            QTableWidget::item {
                padding: 8px;
                border: none;
                word-wrap: break-word;
            }
            QTableWidget::item:selected {
                background-color: #4CAF50;
                color: white;
            }
            QHeaderView::section {
                background-color: #4CAF50;
                color: white;
                padding: 12px;
                border: none;
                font-weight: 600;
            }
            QScrollBar:vertical {
                background: #f0f0f0;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background: #c0c0c0;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a0a0a0;
            }
        """)
        
        # Dapatkan ukuran desktop
        desktop = QDesktopWidget()
        screen_geometry = desktop.screenGeometry()
        
        # Set ukuran window menjadi 85% dari ukuran desktop
        window_width = int(screen_geometry.width() * 0.85)
        window_height = int(screen_geometry.height() * 0.85)
        
        # Posisikan window di tengah layar
        x = (screen_geometry.width() - window_width) // 2
        y = (screen_geometry.height() - window_height) // 2
        
        self.setGeometry(x, y, window_width, window_height)
        self.setMinimumSize(800, 600)
        
        # Status bar dengan style modern
        self.statusBar().setStyleSheet("""
            QStatusBar {
                background-color: #4CAF50;
                color: white;
                font-size: 14px;
                padding: 8px;
            }
        """)
        
        # Tab Interface
        self.tabs = QTabWidget()
        self.tab1 = QWidget()  # Surat Masuk
        self.tab2 = QWidget()  # Surat Keluar
        self.tab3 = QWidget()  # Lihat Data
        
        self.tabs.addTab(self.tab1, "📥 Surat Masuk")
        self.tabs.addTab(self.tab2, "📤 Surat Keluar")
        self.tabs.addTab(self.tab3, "📊 Lihat Data")
        
        # Setup tabs
        self.setup_tab(self.tab1, "Masuk")
        self.setup_tab(self.tab2, "Keluar")
        self.setup_data_view()
        
        self.setCentralWidget(self.tabs)
        
        # Tampilkan pesan selamat datang
        ModernMessageBox.show_success(self, "Selamat Datang", "✅ Selamat datang di Aplikasi Persuratan SLBN Tanah Grogot ✅\n\nBy: Fahmi Firdausi")
    
    def setup_tab(self, tab, jenis):
        # Scroll area untuk konten yang panjang
        scroll_area = QScrollArea()
        scroll_widget = QWidget()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(scroll_widget)
        
        main_layout = QVBoxLayout()
        
        # Header dengan judul
        header = QFrame()
        header.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, 
                    stop:0 #4CAF50, stop:1 #45a049);
                border-radius: 12px;
                padding: 20px;
                margin-bottom: 20px;
            }
        """)
        header_layout = QVBoxLayout()
        
        title_label = QLabel(f"📝 Form Surat {jenis}")
        title_label.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 24px;
                font-weight: 700;
                margin: 0;
            }
        """)
        
        subtitle_label = QLabel(f"Masukkan data surat {jenis.lower()} dengan lengkap")
        subtitle_label.setStyleSheet("""
            QLabel {
                color: rgba(255, 255, 255, 0.8);
                font-size: 14px;
                margin: 0;
            }
        """)
        
        header_layout.addWidget(title_label)
        header_layout.addWidget(subtitle_label)
        header.setLayout(header_layout)
        
        # Form dalam group box
        form_group = ModernGroupBox("📋 Informasi Surat")
        form_layout = QGridLayout()
        
        # Input fields dengan style modern
        setattr(self, f'no_surat_{jenis.lower()}', ModernLineEdit(f"Masukkan nomor surat {jenis.lower()}"))
        setattr(self, f'tanggal_{jenis.lower()}', ModernDateEdit())
        setattr(self, f'pihak_{jenis.lower()}', ModernLineEdit("Pengirim" if jenis == "Masuk" else "Tujuan"))
        setattr(self, f'perihal_{jenis.lower()}', ModernTextEdit("Masukkan perihal surat..."))
        
        # Labels dengan style modern
        form_layout.addWidget(ModernLabel(f"📄 No. Surat {jenis}:", 16, "bold", "#4CAF50"), 0, 0)
        form_layout.addWidget(getattr(self, f'no_surat_{jenis.lower()}'), 0, 1)
        
        form_layout.addWidget(ModernLabel("📅 Tanggal:", 16, "bold", "#4CAF50"), 1, 0)
        form_layout.addWidget(getattr(self, f'tanggal_{jenis.lower()}'), 1, 1)
        
        form_layout.addWidget(ModernLabel("👤 " + ("Pengirim:" if jenis == "Masuk" else "Tujuan:"), 16, "bold", "#4CAF50"), 2, 0)
        form_layout.addWidget(getattr(self, f'pihak_{jenis.lower()}'), 2, 1)
        
        form_layout.addWidget(ModernLabel("📝 Perihal:", 16, "bold", "#4CAF50"), 3, 0)
        form_layout.addWidget(getattr(self, f'perihal_{jenis.lower()}'), 3, 1)
        
        # Button dengan style modern
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        save_btn = ModernButton(f"💾 Simpan Surat {jenis}", "#4CAF50", "#45a049")
        save_btn.clicked.connect(lambda: self.save_data(jenis))
        
        clear_btn = ModernButton("🗑️ Bersihkan Form", "#f44336", "#da190b")
        clear_btn.clicked.connect(lambda: self.clear_form(jenis))
        
        button_layout.addWidget(clear_btn)
        button_layout.addWidget(save_btn)
        
        form_layout.addLayout(button_layout, 4, 0, 1, 2)
        form_group.setLayout(form_layout)
        
        # Layout utama
        main_layout.addWidget(header)
        main_layout.addWidget(form_group)
        main_layout.addStretch()
        
        scroll_widget.setLayout(main_layout)
        
        # Layout tab
        tab_layout = QVBoxLayout()
        tab_layout.setContentsMargins(20, 20, 20, 20)
        tab_layout.addWidget(scroll_area)
        tab.setLayout(tab_layout)
    
    def setup_data_view(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(20, 20, 20, 20)
        
        # Header
        header = QFrame()
        header.setStyleSheet("""
            QFrame {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, 
                    stop:0 #2196F3, stop:1 #1976D2);
                border-radius: 12px;
                padding: 20px;
                margin-bottom: 20px;
            }
        """)
        header_layout = QHBoxLayout()
        
        title_label = QLabel("📊 Database Surat")
        title_label.setStyleSheet("""
            QLabel {
                color: white;
                font-size: 24px;
                font-weight: 700;
            }
        """)
        
        refresh_btn = ModernButton("🔄 Refresh Data", "#2196F3", "#1976D2")
        refresh_btn.clicked.connect(self.load_data)
        
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(refresh_btn)
        header.setLayout(header_layout)
        
        # Tabel data
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        
        # Tabel hanya bisa dilihat, tidak bisa diedit
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        # Set word wrap untuk tabel
        self.table.setWordWrap(True)
        
        # Set resize mode untuk kolom dan baris
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        
        # Set stretch untuk kolom terakhir agar mengisi sisa ruang
        self.table.horizontalHeader().setStretchLastSection(True)
        
        self.load_data()
        
        main_layout.addWidget(header)
        main_layout.addWidget(self.table)
        self.tab3.setLayout(main_layout)
    
    def load_data(self):
        try:
            records = sheet.get_all_records()
            self.table.setRowCount(len(records))
            self.table.setColumnCount(len(records[0]) if records else 0)
            
            if records:
                headers = list(records[0].keys())
                self.table.setHorizontalHeaderLabels(headers)
                
                for row_idx, row in enumerate(records):
                    for col_idx, key in enumerate(headers):
                        item = QTableWidgetItem(str(row[key]))
                        item.setTextAlignment(Qt.AlignLeft | Qt.AlignTop)
                        # Set item tidak bisa diedit
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                        self.table.setItem(row_idx, col_idx, item)
                
                # Auto-resize kolom berdasarkan konten
                self.table.resizeColumnsToContents()
                
                # Auto-resize baris berdasarkan konten
                self.table.resizeRowsToContents()
                
                # Set minimum row height untuk readability
                for row in range(self.table.rowCount()):
                    self.table.setRowHeight(row, max(self.table.rowHeight(row), 40))
                
                # Set minimum column width untuk readability
                for col in range(self.table.columnCount()):
                    current_width = self.table.columnWidth(col)
                    self.table.setColumnWidth(col, max(current_width, 100))
                
                # Set stretch untuk kolom terakhir
                self.table.horizontalHeader().setStretchLastSection(True)
                
            ModernMessageBox.show_success(self, "Berhasil", "✅ Data berhasil dimuat! ✅")
        except Exception as e:
            ModernMessageBox.show_error(self, "Error", f"❌ Terjadi kesalahan saat memuat data:\n{str(e)}")
    
    def save_data(self, jenis):
        try:
            no_surat = getattr(self, f'no_surat_{jenis.lower()}').text()
            tanggal_widget = getattr(self, f'tanggal_{jenis.lower()}')
            tanggal = tanggal_widget.date().toString("dd/MM/yyyy")
            pihak = getattr(self, f'pihak_{jenis.lower()}').text()
            perihal = getattr(self, f'perihal_{jenis.lower()}').toPlainText()
            
            if not all([no_surat, pihak, perihal]):
                ModernMessageBox.show_warning(self, "Peringatan", "⚠️ Mohon lengkapi semua field yang diperlukan!\n\n• No. Surat\n• Pihak\n• Perihal")
                return
            
            # Removed the "File" field - now only saving the essential data
            data = {
                "Tanggal": tanggal,
                "No. Surat": no_surat,
                "Pihak": pihak,
                "Perihal": perihal,
                "Jenis": jenis
            }
            
            sheet.append_row(list(data.values()))
            ModernMessageBox.show_success(self, "Berhasil", f"✅ Data surat {jenis.lower()} berhasil disimpan!\n\nNo. Surat: {no_surat}\nTanggal: {tanggal}")
            self.load_data()
            self.clear_form(jenis)
            
        except Exception as e:
            ModernMessageBox.show_error(self, "Error", f"❌ Terjadi kesalahan saat menyimpan data:\n{str(e)}")
    
    def clear_form(self, jenis):
        getattr(self, f'no_surat_{jenis.lower()}').clear()
        getattr(self, f'tanggal_{jenis.lower()}').setDate(QDate.currentDate())  # Reset ke tanggal hari ini
        getattr(self, f'pihak_{jenis.lower()}').clear()
        getattr(self, f'perihal_{jenis.lower()}').clear()
        ModernMessageBox.show_success(self, "Berhasil", "🧹 Form berhasil dibersihkan!")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Set font aplikasi
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    # Suppress Qt warnings
    os.environ["QT_LOGGING_RULES"] = "qt.qpa.plugin=false"
    
    window = SuratApp()
    window.show()
    sys.exit(app.exec_())
