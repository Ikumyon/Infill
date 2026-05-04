import sys
import os
import subprocess
import json
from pathlib import Path
from typing import Optional, List, Dict, Any

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QGridLayout, QLabel, QLineEdit, 
                             QPushButton, QComboBox, QTextEdit, QCheckBox, 
                             QFileDialog, QMessageBox, QGroupBox,
                             QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont, QIcon
import ctypes

# Windowsタスクバーでアイコンを正しく表示させるための設定
try:
    myappid = 'mycompany.myproduct.subproduct.version' # 任意のユニークなID
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except Exception:
    pass

from core.processor import CSVProcessor, TemplateProcessor, Exporter

def resource_path(relative_path: str) -> str:
    """リソース（アイコン等）の絶対パスを返す（PyInstallerの一時展開先にも対応）"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# UI Styles
STYLE_DROP_NORMAL = """
    QLabel {
        border: 2px dashed #aaa;
        border-radius: 15px;
        background-color: rgba(255, 255, 255, 100);
        color: #888;
        font-size: 16px;
        padding: 20px;
    }
"""

STYLE_DROP_VALID = """
    QLabel {
        border: 3px dashed #409EFF;
        border-radius: 15px;
        background-color: rgba(64, 158, 255, 40);
        color: #409EFF;
        font-size: 18px;
        font-weight: bold;
        padding: 20px;
    }
"""

STYLE_DROP_INVALID = """
    QLabel {
        border: 3px dashed #f56c6c;
        border-radius: 15px;
        background-color: rgba(245, 108, 108, 40);
        color: #f56c6c;
        font-size: 18px;
        font-weight: bold;
        padding: 20px;
    }
"""

class DropZone(QLabel):
    """ドラッグ＆ドロップを受け付けるゾーンのUIコンポーネント"""
    STATE_NORMAL = 0
    STATE_VALID = 1
    STATE_INVALID = 2

    def __init__(self, title: str, description: str, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.title = title
        self.description = description
        self.setAlignment(Qt.AlignCenter)
        self.setWordWrap(True)
        self.set_state(self.STATE_NORMAL)

    def set_state(self, state: int, error_msg: str = "") -> None:
        """ゾーンの状態（通常、有効、無効）に応じて表示を更新する"""
        if state == self.STATE_VALID:
            style = STYLE_DROP_VALID
            text = f"【{self.title}】\n\nここにドロップして確定"
        elif state == self.STATE_INVALID:
            style = STYLE_DROP_INVALID
            text = f"【{self.title}】\n\n{error_msg}"
        else:
            style = STYLE_DROP_NORMAL
            text = f"【{self.title}】\n\n{self.description}"
        
        self.setStyleSheet(style)
        self.setText(text)

class DropOverlay(QWidget):
    """ドラッグ時に表示されるオーバーレイUI"""
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.hide()
        
        # 半透明白の背景
        self.bg = QWidget(self)
        self.bg.setStyleSheet("background-color: rgba(255, 255, 255, 200);")
        
        layout = QHBoxLayout(self)
        layout.setContentsMargins(9, 9, 9, 9)
        layout.setSpacing(6)
        
        self.zone_csv = DropZone("CSVファイル", "CSVをここにドロップ")
        self.zone_tmpl = DropZone("ひな形ファイル", "ひな形をここにドロップ")
        self.zone_dest = DropZone("出力先フォルダ", "フォルダをここにドロップ")
        
        # 子要素のイベント透過設定
        for zone in [self.zone_csv, self.zone_tmpl, self.zone_dest]:
            zone.setAttribute(Qt.WA_TransparentForMouseEvents)
            layout.addWidget(zone)

    def resizeEvent(self, event) -> None:
        self.bg.resize(self.size())
        super().resizeEvent(event)

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            self.update_zones(event.position().toPoint(), event.mimeData().urls())
            event.accept()

    def dragLeaveEvent(self, event) -> None:
        self.hide()
        event.accept()

    def dropEvent(self, event) -> None:
        pos = event.position().toPoint()
        urls = event.mimeData().urls()
        target = self.update_zones(pos, urls)
        self.hide()
        
        if hasattr(self.parent(), "handle_drop"):
            self.parent().handle_drop(target, pos, urls)
        event.accept()

    def update_zones(self, pos, urls) -> Optional[str]:
        if not urls: return None
        
        path = urls[0].toLocalFile()
        is_dir = os.path.isdir(path)
        ext = os.path.splitext(path)[1].lower()

        target = None
        if self.zone_csv.geometry().contains(pos): target = "csv"
        elif self.zone_tmpl.geometry().contains(pos): target = "tmpl"
        elif self.zone_dest.geometry().contains(pos): target = "dest"

        # 各ゾーンのステータスリセット
        self.zone_csv.set_state(DropZone.STATE_NORMAL)
        self.zone_tmpl.set_state(DropZone.STATE_NORMAL)
        self.zone_dest.set_state(DropZone.STATE_NORMAL)

        if target == "csv":
            if is_dir: self.zone_csv.set_state(DropZone.STATE_INVALID, "フォルダは不可")
            elif ext != ".csv": self.zone_csv.set_state(DropZone.STATE_INVALID, "CSVのみ有効")
            else: self.zone_csv.set_state(DropZone.STATE_VALID)
        elif target == "tmpl":
            if is_dir: self.zone_tmpl.set_state(DropZone.STATE_INVALID, "フォルダは不可")
            else: self.zone_tmpl.set_state(DropZone.STATE_VALID)
        elif target == "dest":
            if not is_dir: self.zone_dest.set_state(DropZone.STATE_INVALID, "フォルダのみ有効")
            else: self.zone_dest.set_state(DropZone.STATE_VALID)

        return target

class MainWindow(QMainWindow):
    SETTINGS_FILE = "settings.json"

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Infill - CSV差し込みテキスト出力ツール")
        self.setWindowIcon(QIcon(resource_path("app_icon.ico")))
        self.resize(1100, 850)
        self.setAcceptDrops(True)
        
        self.csv_proc = CSVProcessor()
        self.tmpl_proc = TemplateProcessor()
        
        self.init_ui()
        self.load_settings()
        
        self.overlay = DropOverlay(self)
        self.overlay.resize(self.size())

    def resizeEvent(self, event) -> None:
        if hasattr(self, 'overlay'):
            self.overlay.resize(self.size())
        super().resizeEvent(event)

    def init_ui(self) -> None:
        """UI全体の構成を行う"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(6)
        main_layout.setContentsMargins(9, 9, 9, 9)

        # 各セクションの初期化
        main_layout.addWidget(self._init_file_group())
        
        self.tabs = QTabWidget()
        self.tabs.addTab(self._init_preview_tab(), "プレビュー")
        self.tabs.addTab(self._init_csv_table(), "CSV一覧")
        self.tabs.addTab(self._init_log_text(), "出力ログ")
        main_layout.addWidget(self.tabs)

        main_layout.addWidget(self._init_export_settings())
        main_layout.addLayout(self._init_footer())

        # シグナル接続
        self.export_btn.setEnabled(False)
        self.start_delim_edit.textChanged.connect(self.update_example_label)
        self.end_delim_edit.textChanged.connect(self.update_example_label)
        self.csv_path_edit.returnPressed.connect(lambda: self.load_csv_data(self.csv_path_edit.text()))
        self.tmpl_path_edit.returnPressed.connect(lambda: self.load_template_data(self.tmpl_path_edit.text()))
        self.dest_path_edit.returnPressed.connect(self.validate_files)

    def _init_file_group(self) -> QGroupBox:
        group = QGroupBox("ファイル設定")
        layout = QGridLayout(group)
        layout.setSpacing(6)

        self.csv_path_edit = QLineEdit()
        self.csv_path_edit.setPlaceholderText("CSVファイルをここにドロップ または 参照...")
        self.csv_browse_btn = QPushButton("参照...")
        self.csv_browse_btn.clicked.connect(self.browse_csv)

        self.tmpl_path_edit = QLineEdit()
        self.tmpl_path_edit.setPlaceholderText("ひな形ファイルをここにドロップ または 参照...")
        self.tmpl_browse_btn = QPushButton("参照...")
        self.tmpl_browse_btn.clicked.connect(self.browse_template)

        self.dest_path_edit = QLineEdit()
        self.dest_path_edit.setPlaceholderText("出力先フォルダを選択してください...")
        self.dest_browse_btn = QPushButton("参照...")
        self.dest_browse_btn.clicked.connect(self.browse_dest)

        layout.addWidget(QLabel("CSVファイル"), 0, 0)
        layout.addWidget(self.csv_path_edit, 0, 1)
        layout.addWidget(self.csv_browse_btn, 0, 2)
        layout.addWidget(QLabel("ひな形ファイル"), 1, 0)
        layout.addWidget(self.tmpl_path_edit, 1, 1)
        layout.addWidget(self.tmpl_browse_btn, 1, 2)
        layout.addWidget(QLabel("出力先フォルダ"), 2, 0)
        layout.addWidget(self.dest_path_edit, 2, 1)
        layout.addWidget(self.dest_browse_btn, 2, 2)

        delim_layout = QHBoxLayout()
        delim_layout.setSpacing(6)
        self.start_delim_edit = QLineEdit("{")
        self.start_delim_edit.setMinimumWidth(100)
        self.end_delim_edit = QLineEdit("}")
        self.end_delim_edit.setMinimumWidth(100)
        self.apply_delim_btn = QPushButton("適用")
        self.apply_delim_btn.clicked.connect(self.apply_delimiters)
        self.example_label = QLabel("  例： {name}")
        self.example_label.setStyleSheet("color: #666; font-style: italic;")
        
        delim_layout.addWidget(QLabel("開始文字列"))
        delim_layout.addWidget(self.start_delim_edit)
        delim_layout.addWidget(QLabel("終了文字列"))
        delim_layout.addWidget(self.end_delim_edit)
        delim_layout.addWidget(self.apply_delim_btn)
        delim_layout.addWidget(self.example_label)
        delim_layout.addStretch()
        layout.addLayout(delim_layout, 3, 1)
        return group

    def _init_preview_tab(self) -> QWidget:
        widget = QWidget()
        layout = QHBoxLayout(widget)
        layout.setContentsMargins(9, 9, 9, 9)
        layout.setSpacing(6)
        
        confirm_group = QGroupBox("確認")
        confirm_layout = QVBoxLayout(confirm_group)
        confirm_layout.setContentsMargins(9, 9, 9, 9)
        confirm_layout.setSpacing(6)
        self.csv_cols_label = QLabel("CSV列: -")
        self.tmpl_items_label = QLabel("ひな形項目: -")
        self.status_label = QLabel("状態: 待機中")
        confirm_layout.addWidget(self.csv_cols_label)
        confirm_layout.addWidget(self.tmpl_items_label)
        confirm_layout.addStretch()
        confirm_layout.addWidget(self.status_label)
        
        preview_group = QGroupBox("プレビュー (黄:差し込み / 赤:未置換項目)")
        preview_layout = QVBoxLayout(preview_group)
        preview_layout.setContentsMargins(9, 9, 9, 9)
        preview_layout.setSpacing(6)
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setFont(QFont("Consolas", 10))
        
        ctrl_layout = QHBoxLayout()
        self.preview_row_combo = QComboBox()
        self.preview_row_combo.currentIndexChanged.connect(self.update_preview)
        self.prev_btn = QPushButton("前へ")
        self.prev_btn.clicked.connect(lambda: self.change_preview_row(-1))
        self.next_btn = QPushButton("次へ")
        self.next_btn.clicked.connect(lambda: self.change_preview_row(1))
        
        ctrl_layout.addWidget(QLabel("プレビュー行:"))
        ctrl_layout.addWidget(self.preview_row_combo)
        ctrl_layout.addWidget(self.prev_btn)
        ctrl_layout.addWidget(self.next_btn)
        
        preview_layout.addWidget(self.preview_text)
        preview_layout.addLayout(ctrl_layout)
        
        layout.addWidget(confirm_group, 1)
        layout.addWidget(preview_group, 2)
        return widget

    def _init_csv_table(self) -> QTableWidget:
        self.csv_table = QTableWidget()
        self.csv_table.setAlternatingRowColors(True)
        self.csv_table.setEditTriggers(QTableWidget.NoEditTriggers)
        return self.csv_table

    def _init_log_text(self) -> QTextEdit:
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        return self.log_text

    def _init_export_settings(self) -> QGroupBox:
        group = QGroupBox("出力設定")
        layout = QGridLayout(group)
        layout.setContentsMargins(9, 9, 9, 9)
        layout.setSpacing(6)
        
        # 1. ウィジェットの作成
        self.filename_col_combo = QComboBox()
        self.filename_col_combo.addItem("(連番出力)")
        self.encoding_combo = QComboBox()
        self.encoding_combo.addItems(["UTF-8", "UTF-8 BOM", "Shift_JIS"])
        self.newline_combo = QComboBox()
        self.newline_combo.addItems(["Windows (CRLF)", "Unix (LF)"])
        
        self.overwrite_check = QCheckBox("既存ファイルを上書きする")
        self.open_folder_check = QCheckBox("出力後にフォルダを開く")
        self.open_folder_check.setChecked(True)

        self.combine_single_check = QCheckBox("1つのファイルに結合する")
        self.separator_label = QLabel("区切り:")
        self.separator_edit = QLineEdit("---")
        self.separator_edit.setMaximumWidth(80)
        self.combined_fn_label = QLabel("ファイル名:")
        self.combined_fn_edit = QLineEdit("infill_output.txt")

        # 2. レイアウトへの配置
        # 1段目: 基本設定
        layout.addWidget(QLabel("ファイル名列"), 0, 0)
        layout.addWidget(self.filename_col_combo, 0, 1)
        layout.addWidget(QLabel("文字コード"), 0, 2)
        layout.addWidget(self.encoding_combo, 0, 3)
        layout.addWidget(QLabel("改行コード"), 0, 4)
        layout.addWidget(self.newline_combo, 0, 5)

        # 2段目: 出力オプション
        options_layout = QHBoxLayout()
        options_layout.setSpacing(15)
        options_layout.addWidget(self.overwrite_check)
        options_layout.addWidget(self.open_folder_check)
        layout.addLayout(options_layout, 1, 0, 1, 2)

        # 結合設定のグループ化
        combine_layout = QHBoxLayout()
        combine_layout.setSpacing(6)
        combine_layout.addWidget(self.combine_single_check)
        combine_layout.addWidget(self.separator_label)
        combine_layout.addWidget(self.separator_edit)
        combine_layout.addWidget(self.combined_fn_label)
        combine_layout.addWidget(self.combined_fn_edit)
        combine_layout.addStretch()
        layout.addLayout(combine_layout, 1, 2, 1, 4)

        # 3. シグナル接続
        self.combine_single_check.toggled.connect(self._update_export_ui_state)
        
        # 初期状態の反映
        self._update_export_ui_state()
        
        return group

    def _update_export_ui_state(self) -> None:
        """結合設定の有効/無効状態を一括更新する"""
        is_comb = self.combine_single_check.isChecked()
        
        # 結合用ウィジェットの有効化
        for w in [self.separator_label, self.separator_edit, 
                  self.combined_fn_label, self.combined_fn_edit]:
            w.setEnabled(is_comb)
            
        # 結合時は「ファイル名列」の選択を無効化（混乱防止）
        self.filename_col_combo.setEnabled(not is_comb)

    def _init_footer(self) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setContentsMargins(9, 9, 9, 9)
        layout.setSpacing(6)
        self.footer_status = QLabel("準備完了")
        self.export_btn = QPushButton("出力実行")
        self.export_btn.clicked.connect(self.run_export)
        layout.addWidget(self.footer_status, 1)
        layout.addWidget(self.export_btn, 0)
        return layout

    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.accept()
            self.overlay.show()
            self.overlay.raise_()

    def handle_drop(self, target: str, pos, urls) -> None:
        files = [u.toLocalFile() for u in urls]
        if not files: return

        f = files[0]
        loaded = []
        
        if target == "csv":
            if not os.path.isdir(f) and f.lower().endswith(".csv"):
                self.csv_path_edit.setText(f)
                self.load_csv_data(f)
                loaded.append("CSV")
        elif target == "tmpl":
            if not os.path.isdir(f):
                self.tmpl_path_edit.setText(f)
                self.load_template_data(f)
                loaded.append("ひな形")
        elif target == "dest":
            if os.path.isdir(f):
                self.dest_path_edit.setText(f)
                loaded.append("出力先")
                self.validate_files()

        if loaded:
            self.footer_status.setText(" / ".join(loaded) + "を読み込みました")

    def save_settings(self) -> None:
        settings = {
            "csv_path": self.csv_path_edit.text(),
            "tmpl_path": self.tmpl_path_edit.text(),
            "dest_path": self.dest_path_edit.text(),
            "start_delim": self.start_delim_edit.text(),
            "end_delim": self.end_delim_edit.text(),
            "encoding": self.encoding_combo.currentIndex(),
            "newline": self.newline_combo.currentIndex(),
            "overwrite": self.overwrite_check.isChecked(),
            "combine_single": self.combine_single_check.isChecked(),
            "separator": self.separator_edit.text(),
            "combined_filename": self.combined_fn_edit.text(),
            "open_folder": self.open_folder_check.isChecked()
        }
        try:
            with open(self.SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=4)
        except Exception: pass

    def load_settings(self) -> None:
        if not os.path.exists(self.SETTINGS_FILE): return
        try:
            with open(self.SETTINGS_FILE, "r", encoding="utf-8") as f:
                s = json.load(f)
            
            self.start_delim_edit.setText(s.get("start_delim", "{"))
            self.end_delim_edit.setText(s.get("end_delim", "}"))
            self.encoding_combo.setCurrentIndex(s.get("encoding", 0))
            self.newline_combo.setCurrentIndex(s.get("newline", 0))
            self.overwrite_check.setChecked(s.get("overwrite", False))
            self.combine_single_check.setChecked(s.get("combine_single", False))
            self.separator_edit.setText(s.get("separator", "---"))
            self.combined_fn_edit.setText(s.get("combined_filename", "combined_output.txt"))
            
            is_comb = self.combine_single_check.isChecked()
            for w in [self.separator_edit, self.separator_label, self.combined_fn_edit, self.combined_fn_label]:
                w.setEnabled(is_comb)
            self.filename_col_combo.setEnabled(not is_comb)

            self.open_folder_check.setChecked(s.get("open_folder", True))

            cp, tp, dp = s.get("csv_path", ""), s.get("tmpl_path", ""), s.get("dest_path", "")
            if cp and os.path.exists(cp):
                self.csv_path_edit.setText(cp)
                self.load_csv_data(cp)
            if tp and os.path.exists(tp):
                self.tmpl_path_edit.setText(tp)
                self.load_template_data(tp)
            if dp and os.path.isdir(dp):
                self.dest_path_edit.setText(dp)
        except Exception: pass

    def closeEvent(self, event) -> None:
        self.save_settings()
        event.accept()

    def browse_csv(self) -> None:
        p, _ = QFileDialog.getOpenFileName(self, "CSVファイルを選択", "", "CSV Files (*.csv)")
        if p:
            self.csv_path_edit.setText(p)
            self.load_csv_data(p)

    def browse_template(self) -> None:
        p, _ = QFileDialog.getOpenFileName(self, "ひな形ファイルを選択", "", "Text Files (*.txt)")
        if p:
            self.tmpl_path_edit.setText(p)
            self.load_template_data(p)

    def browse_dest(self) -> None:
        p = QFileDialog.getExistingDirectory(self, "出力先フォルダを選択")
        if p: self.dest_path_edit.setText(p)

    def load_csv_data(self, path: str) -> None:
        if self.csv_proc.load(path):
            self.csv_cols_label.setText(f"CSV列: {', '.join(self.csv_proc.headers)}")
            self.preview_row_combo.clear()
            for i in range(len(self.csv_proc.data)): self.preview_row_combo.addItem(str(i + 1))
            self.filename_col_combo.clear()
            self.filename_col_combo.addItem("(連番出力)")
            self.filename_col_combo.addItems(self.csv_proc.headers)
            self.update_csv_table()
            self.validate_files()
        else: QMessageBox.critical(self, "エラー", "CSVファイルを読み込めませんでした。")

    def update_csv_table(self) -> None:
        h, d = self.csv_proc.headers, self.csv_proc.data
        self.csv_table.setColumnCount(len(h))
        self.csv_table.setRowCount(len(d))
        self.csv_table.setHorizontalHeaderLabels(h)
        for r, row in enumerate(d):
            for c, col in enumerate(h):
                v = row.get(col, "")
                self.csv_table.setItem(r, c, QTableWidgetItem(str(v) if v is not None else ""))
        self.csv_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)

    def load_template_data(self, path: str) -> None:
        s, e = self.start_delim_edit.text(), self.end_delim_edit.text()
        if not s or not e:
            QMessageBox.warning(self, "警告", "開始文字列と終了文字列を入力してください。")
            return
        if self.tmpl_proc.load(path, s, e):
            self.tmpl_items_label.setText(f"ひな形項目: {', '.join(self.tmpl_proc.placeholders)}")
            self.validate_files()
        else: QMessageBox.critical(self, "エラー", "ひな形ファイルを読み込めませんでした。")

    def apply_delimiters(self) -> None:
        if not self.tmpl_proc.content: return
        s, e = self.start_delim_edit.text(), self.end_delim_edit.text()
        if not s or not e:
            QMessageBox.warning(self, "警告", "開始文字列と終了文字列を入力してください。")
            return
        self.tmpl_proc.start_delim, self.tmpl_proc.end_delim = s, e
        self.tmpl_proc.parse_placeholders()
        self.tmpl_items_label.setText(f"ひな形項目: {', '.join(self.tmpl_proc.placeholders)}")
        self.validate_files()
        self.update_preview()
        self.footer_status.setText("プレースホルダー文字列を適用しました。")

    def update_example_label(self) -> None:
        s, e = self.start_delim_edit.text(), self.end_delim_edit.text()
        self.example_label.setText(f"  例： {s}name{e}" if s and e else "  例： -")

    def validate_files(self) -> None:
        if not self.csv_proc.headers or not self.tmpl_proc.content: return
        missing = self.tmpl_proc.validate(self.csv_proc.headers)
        if missing:
            self.status_label.setText(f"状態: エラー (不足: {', '.join(missing)})")
            self.status_label.setStyleSheet("color: #f56c6c;")
            self.export_btn.setEnabled(False)
        else:
            self.status_label.setText("状態: OK")
            self.status_label.setStyleSheet("color: #67c23a;")
            self.export_btn.setEnabled(True)
            self.footer_status.setText(f"{len(self.csv_proc.data)}件出力できます。")
            self.update_preview()

    def update_preview(self) -> None:
        idx = self.preview_row_combo.currentIndex()
        if 0 <= idx < len(self.csv_proc.data):
            row = self.csv_proc.data[idx]
            self.preview_text.setHtml(self.tmpl_proc.preview(row, highlight=True))

    def change_preview_row(self, delta: int) -> None:
        ni = self.preview_row_combo.currentIndex() + delta
        if 0 <= ni < self.preview_row_combo.count(): self.preview_row_combo.setCurrentIndex(ni)

    def run_export(self) -> None:
        dest_dir = self.dest_path_edit.text()
        if not dest_dir:
            dest_dir = QFileDialog.getExistingDirectory(self, "出力先フォルダを選択してください")
            if not dest_dir: return
            self.dest_path_edit.setText(dest_dir)

        self.log_text.clear()
        self.log_text.append(f"--- 出力開始: {len(self.csv_proc.data)}件 ---")
        
        def progress_log(idx: int, msg: str, success: bool, error: str) -> None:
            self.log_text.append(msg)
            # 必要に応じてここでプログレスバーなどの更新が可能

        sc, ec = Exporter.batch_export(
            data=self.csv_proc.data,
            template_proc=self.tmpl_proc,
            dest_dir=dest_dir,
            filename_col=self.filename_col_combo.currentText(),
            encoding=self.encoding_combo.currentText(),
            newline='\r\n' if "CRLF" in self.newline_combo.currentText() else '\n',
            overwrite=self.overwrite_check.isChecked(),
            combine_single=self.combine_single_check.isChecked(),
            separator=self.separator_edit.text(),
            output_filename=self.combined_fn_edit.text(),
            progress_callback=progress_log
        )

        self.log_text.append(f"\n--- 完了: 成功 {sc}件 / 失敗 {ec}件 ---")
        QMessageBox.information(self, "完了", f"出力が完了しました。\n成功: {sc}件\n失敗: {ec}件")
        self.tabs.setCurrentIndex(2)
        if self.open_folder_check.isChecked(): self.open_output_folder(dest_dir)

    def open_output_folder(self, p: str) -> None:
        if sys.platform == 'win32': os.startfile(p)
        else: subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', p])

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
