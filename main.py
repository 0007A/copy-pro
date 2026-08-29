import sys
import os
import csv
import pandas as pd
import re
import hashlib
import base64
import io
import json
import sqlite3
import html
import time
import queue
import threading
from PIL import Image as PILImage
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QGridLayout, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QFileDialog, QMessageBox, QLabel, QTextEdit, 
                             QMenu, QDialog, QDialogButtonBox, QComboBox, QSizePolicy, QScrollArea, QFrame, QCheckBox,
                             QTabWidget, QTabBar, QStackedWidget, QProgressBar, QSizeGrip)
from PyQt6.QtCore import Qt, QThread, QObject, pyqtSignal, QPoint, QSize, QTimer, QSettings, QRect
from PyQt6.QtGui import QIcon, QColor, QPixmap
import hashlib

from overlay import Overlay
from ocr_engine import WindowsOCR
from parser import MCQParser

# Firebase Integration
try:
    import firebase_admin
    from firebase_admin import credentials, firestore
    FIREBASE_AVAILABLE = True
except ImportError:
    FIREBASE_AVAILABLE = False

import importlib
import asyncio

# Neural Speech & SAPI Audio Dynamic Integration (IDE-safe dynamic loading)
edge_tts = None
pygame = None
EDGE_TTS_AVAILABLE = False
try:
    edge_tts = importlib.import_module("edge_tts")
    pygame = importlib.import_module("pygame")
    EDGE_TTS_AVAILABLE = True
except Exception:
    EDGE_TTS_AVAILABLE = False

win32com_client = None
WIN32_SAPI_AVAILABLE = False
try:
    win32com_client = importlib.import_module("win32com.client")
    WIN32_SAPI_AVAILABLE = True
except Exception:
    WIN32_SAPI_AVAILABLE = False

speech_recognition = None
SPEECH_RECOG_AVAILABLE = False
try:
    speech_recognition = importlib.import_module("speech_recognition")
    SPEECH_RECOG_AVAILABLE = True
except Exception:
    SPEECH_RECOG_AVAILABLE = False

vosk_module = None
VOSK_AVAILABLE = False
try:
    vosk_module = importlib.import_module("vosk")
    VOSK_AVAILABLE = True
except Exception:
    VOSK_AVAILABLE = False

sounddevice_module = None
SOUNDDEVICE_AVAILABLE = False
try:
    sounddevice_module = importlib.import_module("sounddevice")
    SOUNDDEVICE_AVAILABLE = True
except Exception:
    SOUNDDEVICE_AVAILABLE = False

class OCRThread(QThread):
    result_ready = pyqtSignal(object) 
    error_occurred = pyqtSignal(str)
    def __init__(self, image_path, ocr_engine, parser, target_type):
        super().__init__()
        self.image_path = image_path
        self.ocr_engine = ocr_engine
        self.parser = parser
        self.target_type = target_type
    def run(self):
        try:
            raw_text = self.ocr_engine.extract_text(self.image_path)
            if self.target_type == "Options":
                parsed_data = self.parser.parse(raw_text)
                self.result_ready.emit(parsed_data)
            else:
                self.result_ready.emit(raw_text.strip())
        except Exception as e:
            self.error_occurred.emit(str(e))

class SyncMetadataDialog(QDialog):
    def __init__(self, current_subject, parent=None):
        super().__init__(parent)
        self.setWindowTitle("☁️ Cloud Sync Metadata")
        self.resize(350, 280)
        layout = QVBoxLayout(self)
        self.topic_input = QTextEdit(); self.topic_input.setPlaceholderText("Topic Name"); self.topic_input.setMaximumHeight(35)
        self.subject_input = QTextEdit(); self.subject_input.setText(current_subject); self.subject_input.setMaximumHeight(35)
        self.test_code_input = QTextEdit(); self.test_code_input.setPlaceholderText("Test Code (e.g. T01)"); self.test_code_input.setMaximumHeight(35)
        self.type_input = QComboBox(); self.type_input.addItems(["MCQ", "Subjective", "True/False"])
        self.complete_cb = QCheckBox("Mark as Completed")
        self.complete_cb.setChecked(True)
        
        layout.addWidget(QLabel("Topic Name:"))
        layout.addWidget(self.topic_input)
        layout.addWidget(QLabel("Subject:"))
        layout.addWidget(self.subject_input)
        layout.addWidget(QLabel("Test Code:"))
        layout.addWidget(self.test_code_input)
        layout.addWidget(QLabel("Default Type:"))
        layout.addWidget(self.type_input)
        layout.addWidget(self.complete_cb)
        
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
    def get_data(self):
        return {
            "topic": self.topic_input.toPlainText().strip(),
            "subject": self.subject_input.toPlainText().strip(),
            "test_code": self.test_code_input.toPlainText().strip(),
            "type": self.type_input.currentText(),
            "is_complete": self.complete_cb.isChecked()
        }

class CloudHistoryDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.parent_app = parent
        self.setWindowTitle("☁️ Cloud History")
        self.resize(600, 400)
        layout = QVBoxLayout(self)
        self.list = QTableWidget(0, 4)
        self.list.setHorizontalHeaderLabels(["Topic", "Subject", "Date", "Action"])
        self.list.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.list.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.list.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.list.itemDoubleClicked.connect(self.load_selected)
        layout.addWidget(QLabel("Double click a topic to load/edit, or click 'Upload' for solution page:"))
        layout.addWidget(self.list)
        
        self.fetch_topics()

    def upload_solution_page(self, topic, subject):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Solution Page Image", "", "Images (*.png *.jpg *.jpeg)")
        if not file_path: return

        try:
            # 🖼️ Load, Resize and Compress Image
            img = PILImage.open(file_path)
            if img.width > 1000:
                h = int(img.height * (1000 / img.width))
                img = img.resize((1000, h), PILImage.Resampling.LANCZOS)
            
            output = io.BytesIO()
            img.save(output, format="JPEG", quality=70)
            img_data = output.getvalue()
            base64_str = base64.b64encode(img_data).decode('utf-8')

            # ☁️ Save to 'solution_pages' collection
            doc_id = hashlib.md5(f"sol_{subject}_{topic}".encode()).hexdigest()
            self.db.collection("solution_pages").document(doc_id).set({
                "topic": topic,
                "subject": subject,
                "imageData": base64_str,
                "timestamp": firestore.SERVER_TIMESTAMP
            })
            QMessageBox.information(self, "Success", f"Solution Page uploaded for '{topic}'!")
        except Exception as e:
            QMessageBox.critical(self, "Upload Error", f"Failed to upload: {e}")

    def delete_topic(self, topic, subject):
        ret = QMessageBox.question(self, "Confirm Delete", f"Are you sure you want to delete '{topic}' from cloud?", 
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if ret == QMessageBox.StandardButton.Yes:
            try:
                docs = self.db.collection("quizzes").where("topic", "==", topic).where("subject", "==", subject).stream()
                batch = self.db.batch()
                count = 0
                for doc in docs:
                    batch.delete(doc.reference)
                    count += 1
                
                # Also delete solution pages if any
                sol_docs = self.db.collection("solution_pages").where("topic", "==", topic).where("subject", "==", subject).stream()
                for doc in sol_docs:
                    batch.delete(doc.reference)
                
                batch.commit()
                QMessageBox.information(self, "Deleted", f"Successfully deleted {count} questions from '{topic}'.")
                self.list.setRowCount(0)
                self.fetch_topics() # Refresh list
            except Exception as e:
                QMessageBox.critical(self, "Error", str(e))

    def fetch_topics(self):
        if not self.db: return
        try:
            docs = self.db.collection("quizzes").order_by("timestamp", direction=firestore.Query.DESCENDING).limit(100).stream()
            topics = {}
            for doc in docs:
                d = doc.to_dict()
                key = (d.get("topic"), d.get("subject"))
                if key not in topics:
                    topics[key] = d.get("timestamp")
            
            for (topic, subject), ts in topics.items():
                r = self.list.rowCount(); self.list.insertRow(r)
                self.list.setItem(r, 0, QTableWidgetItem(str(topic)))
                self.list.setItem(r, 1, QTableWidgetItem(str(subject)))
                self.list.setItem(r, 2, QTableWidgetItem(ts.strftime("%Y-%m-%d") if ts else "N/A"))
                
                # Container for buttons
                btn_widget = QWidget()
                btn_layout = QHBoxLayout(btn_widget)
                btn_layout.setContentsMargins(0,0,0,0)
                
                # 📤 Add Upload Button to Row
                btn_upload = QPushButton("📤")
                btn_upload.setStyleSheet("background-color: rgba(0, 120, 212, 150); border-radius: 4px; padding: 2px;")
                btn_upload.clicked.connect(lambda _, t=topic, s=subject: self.upload_solution_page(t, s))
                
                # 🗑️ Add Delete Button
                btn_del = QPushButton("🗑️")
                btn_del.setStyleSheet("background-color: rgba(211,47,47,200); color: white; border-radius: 4px; padding: 2px;")
                btn_del.clicked.connect(lambda _, t=topic, s=subject: self.delete_topic(t, s))
                
                btn_layout.addWidget(btn_upload)
                btn_layout.addWidget(btn_del)
                self.list.setCellWidget(r, 3, btn_widget)
        except: pass
    def load_selected(self, item):
        topic = self.list.item(item.row(), 0).text()
        subject = self.list.item(item.row(), 1).text()
        self.parent_app.load_from_cloud(topic, subject)
        self.accept()

def clean_html_formatting(text):
    if text is None: return ""
    text = str(text).strip()
    if not text: return ""
    text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'</p>', '\n', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', '', text)
    text = html.unescape(text)
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()

def parse_sqlite_quiz_data(file_path):
    """
    Universally parse questions from any SQLite DB (.db, .sqlite, .sqlite3).
    Supports Testbook DB format, Simple Q DB format, and generic question schemas.
    """
    if not os.path.exists(file_path):
        return []
    try:
        conn = sqlite3.connect(file_path)
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%';")
        tables = [t[0] for t in cur.fetchall()]
        if not tables:
            conn.close()
            return []

        def table_score(tname):
            score = 0
            tlow = tname.lower()
            if 'question' in tlow: score += 10
            if 'quiz' in tlow: score += 8
            if 'mcq' in tlow: score += 8
            if 'test' in tlow: score += 5
            return score

        tables.sort(key=table_score, reverse=True)
        
        target_table = None
        best_col_match_count = 0
        target_cols = []
        
        for tname in tables:
            cur.execute(f'PRAGMA table_info("{tname}");')
            cols = [c[1] for c in cur.fetchall()]
            cols_lower = [c.lower() for c in cols]
            matches = sum(1 for c in cols_lower if any(k in c for k in ['question', 'opt', 'choice', 'ans', 'sol', 'exp']))
            if matches > best_col_match_count:
                best_col_match_count = matches
                target_table = tname
                target_cols = cols

        if not target_table:
            target_table = tables[0]
            cur.execute(f'PRAGMA table_info("{target_table}");')
            target_cols = [c[1] for c in cur.fetchall()]

        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute(f'SELECT * FROM "{target_table}"')
        rows = cur.fetchall()
        
        cols_lower = {c.lower(): c for c in target_cols}
        
        def find_col(*patterns):
            for p in patterns:
                for cl, orig in cols_lower.items():
                    if cl == p or cl.startswith(p + '_') or cl.endswith('_' + p) or p in cl:
                        return orig
            return None

        q_col = find_col('question_en', 'question', 'q', 'question_text', 'title', 'prompt')
        q_hn_col = find_col('question_hn', 'question_hi', 'question_hindi')
        
        opt_a_col = find_col('option_1_en', 'option_1', 'optiona', 'option_a', 'opta', 'opt1', 'choice_a', 'choice1', 'opt_1', 'a')
        opt_b_col = find_col('option_2_en', 'option_2', 'optionb', 'option_b', 'optb', 'opt2', 'choice_b', 'choice2', 'opt_2', 'b')
        opt_c_col = find_col('option_3_en', 'option_3', 'optionc', 'option_c', 'optc', 'opt3', 'choice_c', 'choice3', 'opt_3', 'c')
        opt_d_col = find_col('option_4_en', 'option_4', 'optiond', 'option_d', 'optd', 'opt4', 'choice_d', 'choice4', 'opt_4', 'd')
        
        ans_col = find_col('correct_option', 'answer', 'correct_ans', 'ans', 'correct', 'answer_key', 'key')
        sol_col = find_col('solution_en', 'solution', 'explanation', 'exp', 'sol', 'solution_hn')
        img_col = find_col('questionimage', 'image', 'img', 'question_image', 'img_url', 'image_url')

        quiz_list = []
        for r in rows:
            q_en = clean_html_formatting(r[q_col]) if q_col and r[q_col] else ""
            q_hn = clean_html_formatting(r[q_hn_col]) if q_hn_col and r[q_hn_col] else ""
            
            if q_en and q_hn and q_en != q_hn:
                q_text = f"{q_en}\n\n[Hindi]: {q_hn}"
            else:
                q_text = q_en or q_hn or ""
                
            opt_a = clean_html_formatting(r[opt_a_col]) if opt_a_col and r[opt_a_col] else ""
            opt_b = clean_html_formatting(r[opt_b_col]) if opt_b_col and r[opt_b_col] else ""
            opt_c = clean_html_formatting(r[opt_c_col]) if opt_c_col and r[opt_c_col] else ""
            opt_d = clean_html_formatting(r[opt_d_col]) if opt_d_col and r[opt_d_col] else ""
            
            raw_ans = str(r[ans_col]).strip() if ans_col and r[ans_col] is not None else ""
            if raw_ans in ['1', 1, '1.0']: ans = 'A'
            elif raw_ans in ['2', 2, '2.0']: ans = 'B'
            elif raw_ans in ['3', 3, '3.0']: ans = 'C'
            elif raw_ans in ['4', 4, '4.0']: ans = 'D'
            elif raw_ans.upper() in ['A', 'B', 'C', 'D']: ans = raw_ans.upper()
            else:
                if opt_a and raw_ans == opt_a: ans = 'A'
                elif opt_b and raw_ans == opt_b: ans = 'B'
                elif opt_c and raw_ans == opt_c: ans = 'C'
                elif opt_d and raw_ans == opt_d: ans = 'D'
                else: ans = raw_ans
                
            sol = clean_html_formatting(r[sol_col]) if sol_col and r[sol_col] else ""
            img = str(r[img_col]).strip() if img_col and r[img_col] else ""
            
            if q_text or opt_a or opt_b:
                quiz_list.append({
                    "question": q_text,
                    "optionA": opt_a,
                    "optionB": opt_b,
                    "optionC": opt_c,
                    "optionD": opt_d,
                    "answer": ans,
                    "solution": sol,
                    "questionImage": img
                })
        conn.close()
        return quiz_list
    except Exception as e:
        print(f"Error parsing SQLite DB: {e}")
        return []

class BulkImportDialog(QDialog):
    def __init__(self, parent=None, title="Bulk Data Import", expected_format="Question | Option A | Option B | Option C | Option D | Answer | Solution | Type", min_cols=2):
        super().__init__(parent)
        self.setWindowTitle(f"📋 {title}")
        self.resize(750, 520)
        self.min_cols = min_cols
        self.expected_format = expected_format
        
        self.setStyleSheet("""
            QDialog {
                background-color: #1a1a1a;
                color: #FFFFFF;
                font-family: 'Segoe UI', sans-serif;
                font-size: 13px;
            }
            QLabel {
                color: #E0E0E0;
                font-size: 13px;
            }
            QLabel#headerTitle {
                font-size: 16px;
                font-weight: bold;
                color: #00BCD4;
            }
            QLabel#formatHint {
                color: #B0BEC5;
                font-size: 12px;
                background-color: rgba(255, 255, 255, 12);
                border: 1px dashed rgba(255, 255, 255, 30);
                border-radius: 6px;
                padding: 6px 10px;
            }
            QLabel#statusLabel {
                font-weight: bold;
                color: #4CAF50;
                font-size: 12px;
            }
            QTextEdit {
                background-color: #242424;
                border: 1px solid rgba(255, 255, 255, 35);
                border-radius: 8px;
                padding: 8px;
                color: #FFFFFF;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 13px;
                selection-background-color: rgba(0, 120, 212, 180);
            }
            QTextEdit:focus {
                border: 1px solid #00BCD4;
            }
            QPushButton {
                background-color: rgba(255, 255, 255, 18);
                border: 1px solid rgba(255, 255, 255, 35);
                border-radius: 6px;
                padding: 7px 14px;
                color: #FFFFFF;
                font-weight: 500;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 30);
                border-color: #00BCD4;
            }
            QPushButton#btnPrimary {
                background-color: #0078D4;
                border: 1px solid #0098FF;
                font-weight: bold;
                padding: 8px 22px;
                color: #FFFFFF;
            }
            QPushButton#btnPrimary:hover {
                background-color: #1084E3;
            }
            QPushButton#btnFile {
                background-color: rgba(76, 175, 80, 160);
                border: 1px solid #4CAF50;
                font-weight: bold;
                color: #FFFFFF;
            }
            QPushButton#btnFile:hover {
                background-color: rgba(76, 175, 80, 220);
            }
            QComboBox {
                background-color: #2D2D2D;
                border: 1px solid rgba(255, 255, 255, 35);
                border-radius: 6px;
                padding: 5px 10px;
                color: #FFFFFF;
            }
            QCheckBox {
                color: #DDDDDD;
                font-size: 12px;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(12)
        
        # Header
        header_layout = QHBoxLayout()
        title_lbl = QLabel(f"📋 {title}")
        title_lbl.setObjectName("headerTitle")
        header_layout.addWidget(title_lbl)
        header_layout.addStretch()
        
        # File Import Button
        self.btn_file = QPushButton("📁 Browse Excel / CSV File")
        self.btn_file.setObjectName("btnFile")
        self.btn_file.clicked.connect(self.browse_file)
        header_layout.addWidget(self.btn_file)
        layout.addLayout(header_layout)
        
        # Format Hint
        self.format_hint = QLabel(f"💡 Expected Format: {expected_format}")
        self.format_hint.setObjectName("formatHint")
        self.format_hint.setWordWrap(True)
        layout.addWidget(self.format_hint)
        self.label = self.format_hint # Backward compatibility
        
        # Toolbar (Delimiter, Header checkbox, Clipboard, Clear)
        tb_layout = QHBoxLayout()
        tb_layout.addWidget(QLabel("Delimiter:"))
        self.combo_delim = QComboBox()
        self.combo_delim.addItems(["Auto-Detect (Excel / CSV / Pipe)", "Tab (Excel / Google Sheets)", "Pipe ( | )", "Comma ( , )", "Semicolon ( ; )"])
        self.combo_delim.currentIndexChanged.connect(self.update_preview)
        tb_layout.addWidget(self.combo_delim)
        
        self.cb_skip_header = QCheckBox("Skip Header Row")
        self.cb_skip_header.setChecked(False)
        self.cb_skip_header.stateChanged.connect(self.update_preview)
        tb_layout.addWidget(self.cb_skip_header)
        
        tb_layout.addStretch()
        
        self.btn_paste = QPushButton("📋 Paste Clipboard")
        self.btn_paste.clicked.connect(self.paste_clipboard)
        tb_layout.addWidget(self.btn_paste)
        
        self.btn_clear = QPushButton("🧹 Clear")
        self.btn_clear.clicked.connect(lambda: self.text_edit.clear())
        tb_layout.addWidget(self.btn_clear)
        
        layout.addLayout(tb_layout)
        
        # Main text edit
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText(f"Paste your data directly from Excel, Google Sheets, or text.\nExample:\n{expected_format}")
        self.text_edit.textChanged.connect(self.update_preview)
        layout.addWidget(self.text_edit)
        
        # Live status bar
        self.lbl_status = QLabel("ℹ️ Waiting for data...")
        self.lbl_status.setObjectName("statusLabel")
        layout.addWidget(self.lbl_status)
        
        # Dialog buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)
        
        self.btn_ok = QPushButton("✅ Import Data")
        self.btn_ok.setObjectName("btnPrimary")
        self.btn_ok.clicked.connect(self.accept)
        btn_layout.addWidget(self.btn_ok)
        
        layout.addLayout(btn_layout)
        
        self.cached_rows = []

    def browse_file(self):
        f, _ = QFileDialog.getOpenFileName(
            self, 
            "Select Data File", 
            os.path.join(os.path.expanduser("~"), "Downloads"), 
            "All Supported Files (*.tsv *.csv *.xlsx *.xls *.txt *.json *.db *.sqlite *.sqlite3);;SQLite DB Files (*.db *.sqlite *.sqlite3);;Excel Files (*.xlsx *.xls);;TSV Files (*.tsv);;CSV Files (*.csv);;Text Files (*.txt);;JSON Files (*.json);;All Files (*.*)"
        )
        if not f: return
        try:
            lower_f = f.lower()
            if lower_f.endswith(('.db', '.sqlite', '.sqlite3')):
                items = parse_sqlite_quiz_data(f)
                lines = []
                for item in items:
                    q = item.get("question", "").replace("\t", " ").replace("\n", " ")
                    a = item.get("optionA", "").replace("\t", " ").replace("\n", " ")
                    b = item.get("optionB", "").replace("\t", " ").replace("\n", " ")
                    c = item.get("optionC", "").replace("\t", " ").replace("\n", " ")
                    d = item.get("optionD", "").replace("\t", " ").replace("\n", " ")
                    ans = item.get("answer", "")
                    sol = item.get("solution", "").replace("\t", " ").replace("\n", " ")
                    lines.append(f"{q}\t{a}\t{b}\t{c}\t{d}\t{ans}\t{sol}\tMCQ")
                self.text_edit.setPlainText("\n".join(lines))
                self.combo_delim.setCurrentIndex(1) # Tab
            elif lower_f.endswith('.xlsx') or lower_f.endswith('.xls'):
                df = pd.read_excel(f, header=None)
                df = df.fillna('')
                lines = []
                for _, row in df.iterrows():
                    lines.append("\t".join([str(val).strip() for val in row.values]))
                self.text_edit.setPlainText("\n".join(lines))
                self.combo_delim.setCurrentIndex(1) # Tab
            elif lower_f.endswith('.tsv'):
                with open(f, 'r', encoding='utf-8', errors='replace') as fp:
                    self.text_edit.setPlainText(fp.read())
                self.combo_delim.setCurrentIndex(1) # Tab
            elif lower_f.endswith('.csv'):
                with open(f, 'r', encoding='utf-8', errors='replace') as fp:
                    self.text_edit.setPlainText(fp.read())
                self.combo_delim.setCurrentIndex(3) # Comma
            elif lower_f.endswith('.json'):
                with open(f, 'r', encoding='utf-8', errors='replace') as fp:
                    data = json.load(fp)
                if isinstance(data, list):
                    lines = []
                    for item in data:
                        if isinstance(item, dict):
                            q = item.get("question", item.get("Word", ""))
                            a = item.get("optionA", item.get("Meaning", ""))
                            b = item.get("optionB", item.get("Phrases", ""))
                            c = item.get("optionC", item.get("Story", ""))
                            d = item.get("optionD", "")
                            ans = item.get("answer", "")
                            sol = item.get("solution", "")
                            lines.append(f"{q}\t{a}\t{b}\t{c}\t{d}\t{ans}\t{sol}\tMCQ")
                    self.text_edit.setPlainText("\n".join(lines))
                    self.combo_delim.setCurrentIndex(1)
            else:
                with open(f, 'r', encoding='utf-8', errors='replace') as fp:
                    self.text_edit.setPlainText(fp.read())
            self.update_preview()
        except Exception as e:
            QMessageBox.critical(self, "File Error", f"Failed to read file: {e}")

    def paste_clipboard(self):
        cb = QApplication.clipboard()
        text = cb.text()
        if text:
            self.text_edit.setPlainText(text)
            self.update_preview()

    def get_delimiter(self, raw_text):
        mode = self.combo_delim.currentIndex()
        if mode == 1: return '\t'
        elif mode == 2: return '|'
        elif mode == 3: return ','
        elif mode == 4: return ';'
        
        # Auto-detect: count occurrences in first 5 non-empty lines
        lines = [l for l in raw_text.split('\n') if l.strip()][:5]
        sample = "\n".join(lines)
        if not sample: return '\t'
        
        tabs = sample.count('\t')
        pipes = sample.count('|')
        commas = sample.count(',')
        semis = sample.count(';')
        
        counts = [('\t', tabs), ('|', pipes), (',', commas), (';', semis)]
        counts.sort(key=lambda x: x[1], reverse=True)
        if counts[0][1] > 0:
            return counts[0][0]
        return '\t'

    def update_preview(self):
        raw_text = self.text_edit.toPlainText().strip()
        if not raw_text:
            self.lbl_status.setText("ℹ️ Waiting for data...")
            self.lbl_status.setStyleSheet("color: #888888; font-weight: bold;")
            self.cached_rows = []
            return
            
        sep = self.get_delimiter(raw_text)
        delim_name = { '\t': 'Tab (Excel/Sheets)', '|': 'Pipe (|)', ',': 'Comma (,)', ';': 'Semicolon (;)' }.get(sep, 'Auto')
        
        rows = self.parse_text_to_rows(raw_text, sep)
        self.cached_rows = rows
        
        if rows:
            self.lbl_status.setText(f"✨ Detected {len(rows)} valid row(s) | Delimiter: {delim_name}")
            self.lbl_status.setStyleSheet("color: #4CAF50; font-weight: bold;")
        else:
            self.lbl_status.setText(f"⚠️ No valid rows found with delimiter: {delim_name}. Try changing delimiter above.")
            self.lbl_status.setStyleSheet("color: #FF9800; font-weight: bold;")

    def parse_multiline_mcqs(self, raw_text):
        blocks = re.split(r'\n\s*(?:(?:\d+[\.\)]|\bQ(?:uestion)?\s*\d*[\.\:\)]))\s*', '\n' + raw_text)
        results = []
        for block in blocks:
            block = block.strip()
            if not block: continue
            
            ans_pattern = re.compile(r'(?:Ans(?:wer)?|Correct(?:\s*Ans(?:wer)?)?|Key)\s*[\:\.\-\=]?\s*[\(\[]?\s*([A-Da-d])', re.IGNORECASE)
            sol_pattern = re.compile(r'(?:Sol(?:ution)?|Exp(?:lanation)?|Note)\s*[\:\.\-\=]?\s*([^\n\r]+)', re.IGNORECASE)
            
            ans_match = ans_pattern.search(block)
            answer = ans_match.group(1).upper() if ans_match else ""
            
            sol_match = sol_pattern.search(block)
            solution = sol_match.group(1).strip() if sol_match else ""
            
            clean_block = block
            if ans_match: clean_block = clean_block[:ans_match.start()].strip()
            
            opt_pattern = re.compile(r'(?:^|\n)\s*[\(\[]?\s*([A-Da-d])[\s\.\)\:\]]+')
            opt_matches = list(opt_pattern.finditer(clean_block))
            if len(opt_matches) >= 2:
                q_text = clean_block[:opt_matches[0].start()].strip()
                q_text = re.sub(r'^\d+[\.\)\:\s]+', '', q_text).strip()
                q_text = " ".join(q_text.split())
                
                opts = {"A": "", "B": "", "C": "", "D": ""}
                for i in range(len(opt_matches)):
                    letter = opt_matches[i].group(1).upper()
                    start = opt_matches[i].end()
                    end = opt_matches[i+1].start() if i + 1 < len(opt_matches) else len(clean_block)
                    opts[letter] = " ".join(clean_block[start:end].strip().split())
                if q_text:
                    results.append([q_text, opts["A"], opts["B"], opts["C"], opts["D"], answer, solution, "MCQ"])
        return results

    def parse_text_to_rows(self, raw_text, sep):
        if not raw_text.strip(): return []
        
        # 1. Check if multiline MCQ paper format (e.g. Question on 1 line, Options on next lines)
        opt_matches = re.findall(r'(?:^|\n)\s*[\(\[]?\s*[A-Da-d][\s\.\)\:\]]+', raw_text)
        if len(opt_matches) >= 2 and ('|' not in raw_text and '\t' not in raw_text):
            parsed_mcqs = self.parse_multiline_mcqs(raw_text)
            if parsed_mcqs:
                return parsed_mcqs
        
        # 2. Standard tabular CSV / TSV / Pipe parsing with smart delimiter fallback
        delimiters_to_try = [sep]
        for d in ['\t', '|', ',', ';']:
            if d not in delimiters_to_try: delimiters_to_try.append(d)
            
        best_rows = []
        best_avg_cols = 0
        for s in delimiters_to_try:
            try:
                reader = csv.reader(io.StringIO(raw_text), delimiter=s, skipinitialspace=True)
                candidate_rows = []
                for row in reader:
                    cleaned = [col.strip() for col in row]
                    if any(cleaned):
                        candidate_rows.append(cleaned)
                if candidate_rows:
                    avg_cols = sum(len(r) for r in candidate_rows) / len(candidate_rows)
                    if avg_cols > best_avg_cols:
                        best_avg_cols = avg_cols
                        best_rows = candidate_rows
                        if avg_cols >= 2 and s == sep:
                            break
            except Exception:
                pass
                
        if not best_rows:
            # Fallback simple split
            for line in raw_text.split('\n'):
                line = line.strip()
                if not line: continue
                cleaned = [p.strip() for p in line.split(sep)]
                if any(cleaned):
                    best_rows.append(cleaned)
        
        rows = best_rows
        if rows and len(rows) > 0:
            first_row_text = " ".join(rows[0]).lower()
            header_keywords = ["question", "option", "ans", "solution", "word", "meaning", "category", "subject", "test code"]
            looks_like_header = any(k in first_row_text for k in header_keywords)
            if self.cb_skip_header.isChecked() or (looks_like_header and self.cb_skip_header.isChecked()):
                rows = rows[1:]
                
        return rows

    def get_rows(self):
        raw_text = self.text_edit.toPlainText().strip()
        if not raw_text: return []
        sep = self.get_delimiter(raw_text)
        return self.parse_text_to_rows(raw_text, sep)

    def get_text(self):
        return self.text_edit.toPlainText()

class QuizOverviewDialog(QDialog):
    def __init__(self, questions, user_answers, score, seconds_elapsed, parent=None):
        super().__init__(parent)
        self.questions = questions
        self.user_answers = user_answers
        self.score = score
        self.seconds_elapsed = seconds_elapsed
        self.selected_jump_idx = None
        self.cards = []
        
        self.setWindowTitle("📄 Laptop Quiz - A4 Questions & Solution Sheet")
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMinMaxButtonsHint)
        self.resize(820, 1040) # A4 aspect ratio
        
        self.setStyleSheet("""
            QDialog {
                background-color: #121212;
                color: #FFFFFF;
                font-family: 'Segoe UI', sans-serif;
            }
            QScrollArea {
                background-color: #161616;
                border: 1px solid #2D2D2D;
                border-radius: 8px;
            }
            QWidget#a4_container {
                background-color: #161616;
            }
            QFrame#q_sheet_card {
                background-color: #1E1E1E;
                border: 1px solid #333333;
                border-radius: 10px;
            }
            QFrame#q_sheet_card:hover {
                border: 1px solid #0078D4;
            }
            QPushButton#action_btn {
                background-color: #0078D4;
                border: 1px solid #005A9E;
                border-radius: 6px;
                padding: 4px 12px;
                color: white;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton#action_btn:hover {
                background-color: #106EBE;
            }
            QLineEdit#search_box {
                background-color: #1E1E1E;
                border: 1px solid #3A3A3A;
                border-radius: 8px;
                padding: 8px 12px;
                color: #FFFFFF;
                font-size: 13px;
            }
            QLineEdit#search_box:focus {
                border: 1px solid #00BCD4;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(12)

        # Top Summary Header (A4 Stats Banner)
        m = self.seconds_elapsed // 60
        s = self.seconds_elapsed % 60
        attempted = len(self.user_answers)
        total = len(self.questions)
        pct = (self.score / total * 100) if total else 0

        hdr_card = QFrame()
        hdr_card.setStyleSheet("background-color: #1E1E1E; border: 1px solid #333333; border-radius: 10px;")
        hdr_layout = QHBoxLayout(hdr_card)
        hdr_layout.setContentsMargins(14, 10, 14, 10)
        
        lbl_stat1 = QLabel(f"📋 Total: <b>{total}</b>")
        lbl_stat1.setStyleSheet("color: #00BCD4; font-size: 13px;")
        lbl_stat2 = QLabel(f"✍️ Attempted: <b>{attempted}/{total}</b>")
        lbl_stat2.setStyleSheet("color: #FFC107; font-size: 13px;")
        lbl_stat3 = QLabel(f"✅ Correct: <b>{self.score}</b>")
        lbl_stat3.setStyleSheet("color: #4CAF50; font-size: 13px;")
        lbl_stat4 = QLabel(f"❌ Wrong: <b>{attempted - self.score}</b>")
        lbl_stat4.setStyleSheet("color: #F44336; font-size: 13px;")
        lbl_stat5 = QLabel(f"🎯 Score: <b>{pct:.1f}%</b>")
        lbl_stat5.setStyleSheet("color: #9C27B0; font-size: 13px;")
        lbl_stat6 = QLabel(f"⏱️ Time: <b>{m:02d}:{s:02d}</b>")
        lbl_stat6.setStyleSheet("color: #FF9800; font-size: 13px;")

        for lbl in [lbl_stat1, lbl_stat2, lbl_stat3, lbl_stat4, lbl_stat5, lbl_stat6]:
            hdr_layout.addWidget(lbl)
            hdr_layout.addStretch()
            
        btn_download_html = QPushButton("📥 Download HTML")
        btn_download_html.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_download_html.setStyleSheet("background-color: #10B981; border: 1px solid #059669; border-radius: 6px; padding: 6px 14px; color: white; font-weight: bold; font-size: 12px;")
        btn_download_html.setToolTip("Download responsive HTML file to view on Mobile/Phone, Browser or Print to PDF")
        btn_download_html.clicked.connect(self.export_html)
        hdr_layout.addWidget(btn_download_html)
        layout.addWidget(hdr_card)

        # Search / Filter Bar
        from PyQt6.QtWidgets import QLineEdit
        self.search_box = QLineEdit()
        self.search_box.setObjectName("search_box")
        self.search_box.setPlaceholderText("🔍 Quick Filter / Search in Questions & Solutions...")
        self.search_box.textChanged.connect(self.filter_cards)
        layout.addWidget(self.search_box)

        # Scrollable A4 Document Container
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.container = QWidget()
        self.container.setObjectName("a4_container")
        self.cards_layout = QVBoxLayout(self.container)
        self.cards_layout.setContentsMargins(12, 12, 12, 12)
        self.cards_layout.setSpacing(14)

        letters = ["A", "B", "C", "D"]
        for idx, q in enumerate(self.questions):
            ans_raw = q.get("answer", "").strip().upper()
            correct_idx = -1
            for k, letter in enumerate(letters):
                if letter in ans_raw or ans_raw == letter:
                    correct_idx = k
                    break
            if correct_idx == -1: correct_idx = 0

            card = QFrame()
            card.setObjectName("q_sheet_card")
            c_layout = QVBoxLayout(card)
            c_layout.setContentsMargins(16, 14, 16, 14)
            c_layout.setSpacing(8)

            # Card Top Bar (Question Number, Status, Topic, Jump Button)
            top_bar = QHBoxLayout()
            lbl_qnum = QLabel(f"Q{idx + 1}")
            lbl_qnum.setStyleSheet("background-color: #00BCD4; color: #121212; font-weight: bold; font-size: 13px; border-radius: 4px; padding: 2px 8px;")
            top_bar.addWidget(lbl_qnum)

            top_name = get_question_topic(q)
            lbl_top = QLabel(top_name)
            lbl_top.setStyleSheet("background-color: #22222E; color: #CBD5E1; font-size: 11.5px; font-weight: 500; border: 1px solid #3B3B4E; border-radius: 4px; padding: 2px 6px;")
            top_bar.addWidget(lbl_top)

            sel_idx = self.user_answers.get(idx, -1)
            if sel_idx != -1:
                sel_letter = letters[sel_idx]
                if sel_idx == correct_idx:
                    lbl_status = QLabel(f"✅ Your Answer: Option {sel_letter} (Correct)")
                    lbl_status.setStyleSheet("color: #4CAF50; font-weight: bold; font-size: 13px;")
                else:
                    lbl_status = QLabel(f"❌ Your Answer: Option {sel_letter} | Correct: Option {letters[correct_idx]}")
                    lbl_status.setStyleSheet("color: #F44336; font-weight: bold; font-size: 13px;")
            else:
                lbl_status = QLabel(f"⚪ Unanswered | Correct: Option {letters[correct_idx]}")
                lbl_status.setStyleSheet("color: #AAAAAA; font-size: 13px;")
            top_bar.addWidget(lbl_status)
            top_bar.addStretch()

            btn_jump = QPushButton("Go to Q ➔")
            btn_jump.setObjectName("action_btn")
            btn_jump.clicked.connect(lambda _, q_idx=idx: self.select_and_close(q_idx))
            top_bar.addWidget(btn_jump)
            c_layout.addLayout(top_bar)

            # Question Text
            q_text = q.get("question", "").strip()
            lbl_q = QLabel(q_text)
            lbl_q.setStyleSheet("font-size: 15px; font-weight: 600; color: #FFFFFF; line-height: 1.4;")
            lbl_q.setWordWrap(True)
            c_layout.addWidget(lbl_q)

            # Options A, B, C, D
            opt_a = q.get("optionA", "").strip()
            opt_b = q.get("optionB", "").strip()
            opt_c = q.get("optionC", "").strip()
            opt_d = q.get("optionD", "").strip()
            
            if opt_a or opt_b:
                opt_box = QVBoxLayout()
                opt_box.setSpacing(4)
                for opt_i, (opt_letter, opt_val) in enumerate([("A", opt_a), ("B", opt_b), ("C", opt_c), ("D", opt_d)]):
                    if not opt_val: continue
                    lbl_opt = QLabel(f"<b>{opt_letter}.</b>  {opt_val}")
                    lbl_opt.setWordWrap(True)
                    if opt_i == correct_idx:
                        lbl_opt.setStyleSheet("background-color: rgba(76, 175, 80, 20); color: #81C784; padding: 4px 8px; border-radius: 4px; font-size: 13px;")
                    elif opt_i == sel_idx and sel_idx != correct_idx:
                        lbl_opt.setStyleSheet("background-color: rgba(244, 67, 54, 20); color: #E57373; padding: 4px 8px; border-radius: 4px; font-size: 13px;")
                    else:
                        lbl_opt.setStyleSheet("color: #CCCCCC; padding: 2px 8px; font-size: 13px;")
                    opt_box.addWidget(lbl_opt)
                c_layout.addLayout(opt_box)

            # Solution / Important Details
            sol_text = q.get("solution", "").strip()
            if sol_text:
                sol_frame = QFrame()
                sol_frame.setStyleSheet("background-color: #1A2218; border: 1px solid #2E7D32; border-radius: 8px; padding: 8px;")
                s_layout = QVBoxLayout(sol_frame)
                s_layout.setContentsMargins(10, 8, 10, 8)
                lbl_stitle = QLabel("💡 Solution / Important Details:")
                lbl_stitle.setStyleSheet("color: #81C784; font-weight: bold; font-size: 12px;")
                lbl_sbody = QLabel(sol_text)
                lbl_sbody.setStyleSheet("color: #E8F5E9; font-size: 13px; line-height: 1.4;")
                lbl_sbody.setWordWrap(True)
                s_layout.addWidget(lbl_stitle)
                s_layout.addWidget(lbl_sbody)
                c_layout.addWidget(sol_frame)

            self.cards_layout.addWidget(card)
            self.cards.append((card, f"{q_text} {sol_text} Q{idx+1}"))

        self.cards_layout.addStretch()
        self.scroll.setWidget(self.container)
        layout.addWidget(self.scroll)

        # Bottom Bar
        btn_bar = QHBoxLayout()
        lbl_hint = QLabel("💡 Tip: Click 'Go to Q ➔' to jump back, or Download HTML to view on Mobile.")
        lbl_hint.setStyleSheet("color: #888888; font-size: 12px;")
        btn_bar.addWidget(lbl_hint)
        btn_bar.addStretch()
        
        btn_download_html_bot = QPushButton("📥 Download HTML (Mobile / Web)")
        btn_download_html_bot.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_download_html_bot.setStyleSheet("background-color: #10B981; border: 1px solid #059669; border-radius: 6px; padding: 6px 14px; color: white; font-weight: bold; font-size: 12px;")
        btn_download_html_bot.clicked.connect(self.export_html)
        btn_bar.addWidget(btn_download_html_bot)
        
        btn_close = QPushButton("Close A4 Sheet")
        btn_close.setObjectName("action_btn")
        btn_close.clicked.connect(self.accept)
        btn_bar.addWidget(btn_close)
        layout.addLayout(btn_bar)

    def export_html(self):
        import html as html_lib
        from datetime import datetime
        
        total = len(self.questions)
        attempted = len(self.user_answers)
        pct = (self.score / total * 100) if total else 0
        m = self.seconds_elapsed // 60
        s = self.seconds_elapsed % 60
        letters = ["A", "B", "C", "D"]
        
        cards_html = []
        for idx, q in enumerate(self.questions):
            ans_raw = q.get("answer", "").strip().upper()
            corr_idx = 0
            for k, l in enumerate(letters):
                if l in ans_raw or ans_raw == l:
                    corr_idx = k
                    break
            
            sel_idx = self.user_answers.get(idx, -1)
            if sel_idx != -1:
                sel_letter = letters[sel_idx]
                if sel_idx == corr_idx:
                    status_badge = f'<span class="badge badge-correct">✅ Your Answer: Option {sel_letter} (Correct)</span>'
                else:
                    status_badge = f'<span class="badge badge-wrong">❌ Your Answer: Option {sel_letter} | Correct: Option {letters[corr_idx]}</span>'
            else:
                status_badge = f'<span class="badge badge-unanswered">⚪ Unanswered | Correct: Option {letters[corr_idx]}</span>'
                
            q_text = html_lib.escape(q.get("question", "").strip())
            
            opts_html = []
            for opt_i, opt_letter in enumerate(letters):
                val = q.get(f"option{opt_letter}", "").strip()
                if not val: continue
                val_esc = html_lib.escape(val)
                cls = "opt"
                if opt_i == corr_idx:
                    cls += " opt-correct"
                elif opt_i == sel_idx and sel_idx != corr_idx:
                    cls += " opt-wrong"
                opts_html.append(f'<div class="{cls}"><b>{opt_letter}.</b> {val_esc}</div>')
            
            opts_block = "".join(opts_html)
            
            sol_text = q.get("solution", "").strip()
            sol_block = ""
            if sol_text:
                sol_esc = html_lib.escape(sol_text).replace("\n", "<br>")
                sol_block = f'''
                <div class="solution-box">
                    <div class="solution-title">💡 Solution / Important Details:</div>
                    <div class="solution-content">{sol_esc}</div>
                </div>
                '''
                
            img_block = ""
            img_base64 = q.get("questionImage", "")
            if img_base64:
                img_block = f'<div class="img-container"><img src="data:image/png;base64,{img_base64}" alt="Question Image" /></div>'
                
            top_name = get_question_topic(q)
            card_html = f'''
            <div class="question-card" data-content="{html_lib.escape(q.get('question',''))} {html_lib.escape(sol_text)} {html_lib.escape(top_name)} Q{idx+1}">
                <div class="card-header">
                    <span class="q-num">Q{idx+1}</span>
                    <span class="badge" style="background:#22222E; color:#CBD5E1; border:1px solid #3B3B50;">{top_name}</span>
                    {status_badge}
                </div>
                <div class="question-text">{q_text.replace(chr(10), "<br>")}</div>
                {img_block}
                <div class="options-grid">{opts_block}</div>
                {sol_block}
            </div>
            '''
            cards_html.append(card_html)
            
        all_cards = "\n".join(cards_html)
        
        full_html = f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=5.0, user-scalable=yes">
    <title>Quiz Questions & Solutions Sheet</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Noto+Sans+Bengali:wght@400;600;700&display=swap" rel="stylesheet">
    <style>
        :root {{
            --bg: #0F0F12;
            --card-bg: #18181E;
            --border: #2B2B36;
            --text-main: #FFFFFF;
            --text-sub: #94A3B8;
            --accent: #00BCD4;
            --correct-bg: rgba(34, 197, 94, 0.15);
            --correct-border: #22C55E;
            --correct-text: #4ADE80;
            --wrong-bg: rgba(239, 68, 68, 0.15);
            --wrong-border: #EF4444;
            --wrong-text: #F87171;
            --sol-bg: #122119;
            --sol-border: #1E4620;
        }}
        * {{ 
            box-sizing: border-box; 
            margin: 0; 
            padding: 0; 
            -webkit-tap-highlight-color: transparent;
        }}
        body {{
            background-color: var(--bg);
            color: var(--text-main);
            font-family: 'Inter', 'Noto Sans Bengali', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            line-height: 1.6;
            padding: 16px 12px;
            max-width: 920px;
            margin: 0 auto;
            word-wrap: break-word;
            overflow-wrap: break-word;
        }}
        .header-card {{
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: 16px;
            padding: 18px;
            margin-bottom: 16px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.4);
        }}
        .title-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 12px;
            margin-bottom: 14px;
            border-bottom: 1px solid var(--border);
            padding-bottom: 12px;
        }}
        .title-row h1 {{
            font-size: 19px;
            color: var(--accent);
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .actions {{ display: flex; gap: 8px; }}
        .btn {{
            background: #0078D4;
            color: white;
            border: none;
            padding: 8px 16px;
            border-radius: 8px;
            font-weight: 600;
            font-size: 13px;
            cursor: pointer;
            text-decoration: none;
            display: inline-flex;
            align-items: center;
            gap: 6px;
            transition: all 0.2s;
            touch-action: manipulation;
        }}
        .btn:active {{ transform: scale(0.97); }}
        .btn-print {{ background: #10B981; }}
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(130px, 1fr));
            gap: 8px;
        }}
        .stat-badge {{
            background: #202028;
            border: 1px solid var(--border);
            padding: 10px 12px;
            border-radius: 10px;
            font-size: 13px;
            font-weight: 600;
            text-align: center;
        }}
        .search-container {{
            position: sticky;
            top: 10px;
            z-index: 100;
            margin-bottom: 16px;
        }}
        .search-box {{
            width: 100%;
            padding: 14px 18px;
            background: #1C1C24;
            border: 1px solid var(--border);
            border-radius: 12px;
            color: white;
            font-size: 15px;
            outline: none;
            box-shadow: 0 4px 15px rgba(0,0,0,0.5);
            transition: border-color 0.2s;
        }}
        .search-box:focus {{ border-color: var(--accent); }}
        .question-card {{
            background: var(--card-bg);
            border: 1px solid var(--border);
            border-radius: 14px;
            padding: 18px;
            margin-bottom: 14px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.25);
            transition: border-color 0.2s;
        }}
        .question-card:hover {{ border-color: #3B82F6; }}
        .card-header {{
            display: flex;
            align-items: center;
            gap: 10px;
            flex-wrap: wrap;
            margin-bottom: 12px;
        }}
        .q-num {{
            background: var(--accent);
            color: #0F0F12;
            font-weight: 700;
            font-size: 13px;
            padding: 3px 10px;
            border-radius: 6px;
        }}
        .badge {{
            font-size: 12px;
            font-weight: 600;
            padding: 4px 10px;
            border-radius: 6px;
        }}
        .badge-correct {{
            background: var(--correct-bg);
            color: var(--correct-text);
            border: 1px solid var(--correct-border);
        }}
        .badge-wrong {{
            background: var(--wrong-bg);
            color: var(--wrong-text);
            border: 1px solid var(--wrong-border);
        }}
        .badge-unanswered {{
            background: #272730;
            color: var(--text-sub);
            border: 1px solid #3B3B48;
        }}
        .question-text {{
            font-size: 15.5px;
            font-weight: 600;
            margin-bottom: 12px;
            color: #F8FAFC;
            line-height: 1.5;
        }}
        .img-container {{ text-align: center; margin-bottom: 12px; }}
        .img-container img {{
            max-width: 100%;
            height: auto;
            border-radius: 8px;
            border: 1px solid var(--border);
        }}
        .options-grid {{
            display: grid;
            grid-template-columns: 1fr;
            gap: 7px;
            margin-bottom: 12px;
        }}
        .opt {{
            padding: 11px 14px;
            border-radius: 8px;
            background: #1F1F27;
            border: 1px solid #2E2E3C;
            font-size: 14px;
            color: #E2E8F0;
            line-height: 1.4;
        }}
        .opt-correct {{
            background: var(--correct-bg);
            border-color: var(--correct-border);
            color: var(--correct-text);
            font-weight: 600;
        }}
        .opt-wrong {{
            background: var(--wrong-bg);
            border-color: var(--wrong-border);
            color: var(--wrong-text);
            font-weight: 600;
        }}
        .solution-box {{
            background: var(--sol-bg);
            border: 1px solid var(--sol-border);
            border-radius: 10px;
            padding: 12px 15px;
            margin-top: 10px;
        }}
        .solution-title {{
            color: #86EFAC;
            font-weight: 700;
            font-size: 13px;
            margin-bottom: 6px;
        }}
        .solution-content {{
            color: #DCFCE7;
            font-size: 13.5px;
            line-height: 1.5;
        }}
        /* Mobile Specific Responsive Tweaks */
        @media (max-width: 640px) {{
            body {{ padding: 10px 8px; font-size: 14px; }}
            .header-card {{ padding: 14px 12px; }}
            .title-row h1 {{ font-size: 17px; }}
            .stats-grid {{ grid-template-columns: repeat(2, 1fr); gap: 6px; }}
            .stat-badge {{ padding: 8px 10px; font-size: 12px; }}
            .question-card {{ padding: 14px 12px; border-radius: 12px; }}
            .question-text {{ font-size: 14.5px; }}
            .opt {{ padding: 9px 12px; font-size: 13.5px; }}
            .solution-box {{ padding: 10px 12px; }}
        }}
        @media print {{
            body {{ background: white !important; color: black !important; padding: 0; }}
            .header-card, .question-card {{
                background: white !important; color: black !important;
                border: 1px solid #ccc !important; box-shadow: none !important;
                break-inside: avoid;
            }}
            .search-container, .actions {{ display: none !important; }}
            .opt {{ background: #f8f9fa !important; color: #212529 !important; }}
            .opt-correct {{ background: #d1e7dd !important; color: #0f5132 !important; }}
            .opt-wrong {{ background: #f8d7da !important; color: #842029 !important; }}
            .solution-box {{ background: #e8f5e9 !important; color: #1b5e20 !important; border-color: #a5d6a7 !important; }}
        }}
    </style>
</head>
<body>

    <div class="header-card">
        <div class="title-row">
            <h1>📄 Laptop Quiz - Questions & Solutions Sheet</h1>
            <div class="actions">
                <button class="btn btn-print" onclick="window.print()">🖨️ Print / Save PDF</button>
            </div>
        </div>
        <div class="stats-grid">
            <div class="stat-badge" style="color: #00BCD4;">📋 Total: {total}</div>
            <div class="stat-badge" style="color: #FBBF24;">✍️ Attempted: {attempted}/{total}</div>
            <div class="stat-badge" style="color: #4ADE80;">✅ Correct: {self.score}</div>
            <div class="stat-badge" style="color: #F87171;">❌ Wrong: {attempted - self.score}</div>
            <div class="stat-badge" style="color: #C084FC;">🎯 Score: {pct:.1f}%</div>
            <div class="stat-badge" style="color: #FB923C;">⏱️ Time: {m:02d}:{s:02d}</div>
        </div>
    </div>

    <div class="search-container">
        <input type="text" id="searchBox" class="search-box" placeholder="🔍 Instant Search in Questions & Solutions..." oninput="filterQuestions()" />
    </div>

    <div id="questionsContainer">
        {all_cards}
    </div>

    <script>
        function filterQuestions() {{
            const query = document.getElementById('searchBox').value.toLowerCase().trim();
            const cards = document.querySelectorAll('.question-card');
            cards.forEach(card => {{
                const content = card.getAttribute('data-content').toLowerCase();
                if (!query || content.includes(query)) {{
                    card.style.display = 'block';
                }} else {{
                    card.style.display = 'none';
                }}
            }});
        }}
    </script>
</body>
</html>
'''
        # Ask where to save
        default_dir = os.path.join(os.path.expanduser("~"), "Downloads", "Daily quiz")
        if not os.path.exists(default_dir):
            default_dir = os.path.expanduser("~")
        default_file = os.path.join(default_dir, f"quiz_sheet_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
        
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Quiz HTML for Mobile / Web", default_file, "HTML Files (*.html)")
        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(full_html)
                
                reply = QMessageBox.question(
                    self, 
                    "HTML Downloaded Successfully!", 
                    f"✅ Quiz HTML sheet has been saved!\n\n📂 Location: {file_path}\n\n📱 You can open this file on your Mobile phone, transfer it via WhatsApp, or view it offline in any web browser.\n\nWould you like to open it now?", 
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.Yes:
                    os.startfile(file_path)
            except Exception as ex:
                QMessageBox.critical(self, "Error Saving HTML", f"Failed to save HTML file: {ex}")

    def filter_cards(self, text):
        query = text.strip().lower()
        for card, content in self.cards:
            card.setVisible(query in content.lower() if query else True)

    def select_and_close(self, idx):
        self.selected_jump_idx = idx
        self.accept()

def split_english_bengali(text):
    """Splits a bilingual question/option/solution string into English and Bengali parts."""
    if not text:
        return "", ""
    lines = text.strip().split("\n")
    eng_lines = []
    bn_lines = []
    for line in lines:
        l = line.strip()
        if not l:
            continue
        if "/" in l:
            parts = [p.strip() for p in l.split("/") if p.strip()]
            for p in parts:
                if re.search(r'[\u0980-\u09FF]', p):
                    bn_lines.append(p)
                else:
                    eng_lines.append(p)
        else:
            if re.search(r'[\u0980-\u09FF]', l):
                bn_lines.append(l)
            else:
                eng_lines.append(l)
    return " ".join(eng_lines).strip(), " ".join(bn_lines).strip()

def build_question_speech_text(q, q_idx, lang_mode):
    """Constructs the full speech string for a given question and language mode."""
    q_eng, q_bn = split_english_bengali(q.get("question", ""))
    letters = ["A", "B", "C", "D"]
    opt_eng = []
    opt_bn = []
    for i, l in enumerate(letters):
        oe, ob = split_english_bengali(q.get(f"option{l}", ""))
        if oe: opt_eng.append(f"Option {i+1}: {oe}")
        if ob: opt_bn.append(f"অপশন {i+1}: {ob}")
    
    if lang_mode == 'en':
        voice = "en-IN-NeerjaNeural"
        speech = f"Question {q_idx + 1}. {q_eng or q.get('question', '')}. " + ". ".join(opt_eng)
    elif lang_mode == 'bn':
        voice = "bn-IN-TanishaaNeural"
        speech = f"প্রশ্ন {q_idx + 1}. {q_bn or q_eng or q.get('question', '')}. " + ". ".join(opt_bn)
    else: # both
        voice = "bn-IN-TanishaaNeural"
        speech_parts = []
        if q_eng:
            speech_parts.append(f"Question {q_idx + 1}. {q_eng}. " + ". ".join(opt_eng))
        if q_bn:
            speech_parts.append(f"বাংলায় প্রশ্ন {q_idx + 1}. {q_bn}. " + ". ".join(opt_bn))
        speech = " ... ".join(speech_parts) if speech_parts else q.get("question", "")
    
    return speech, voice

def generate_edge_tts_audio_sync(text, voice, rate):
    """Generates audio bytes via edge-tts in a clean isolated async runner."""
    if not EDGE_TTS_AVAILABLE or not text:
        return b''
    clean = re.sub(r'[\*\#\_]', '', text).strip()
    if not clean:
        return b''
    # Auto-switch voice for Bengali characters
    if re.search(r'[\u0980-\u09FF]', clean) and 'en-' in voice:
        voice = "bn-IN-TanishaaNeural"

    async def _fetch():
        comm = edge_tts.Communicate(clean, voice, rate=rate)
        data = b''
        async for chunk in comm.stream():
            if chunk['type'] == 'audio':
                data += chunk['data']
        return data

    try:
        return asyncio.run(_fetch())
    except Exception as e:
        return b''

class BackgroundAudioPreloader(QThread):
    def __init__(self, questions, start_idx, lang_mode, rate, cache_dict):
        super().__init__()
        self.questions = questions
        self.start_idx = start_idx
        self.lang_mode = lang_mode
        self.rate = rate
        self.cache_dict = cache_dict
        self.is_cancelled = False

    def run(self):
        if not EDGE_TTS_AVAILABLE:
            return
        
        # Preload upcoming 4 questions in advance into memory
        for idx in range(self.start_idx, min(len(self.questions), self.start_idx + 4)):
            if self.is_cancelled:
                break
            cache_key = (idx, self.lang_mode)
            if cache_key in self.cache_dict:
                continue
            
            try:
                q = self.questions[idx]
                speech_text, voice = build_question_speech_text(q, idx, self.lang_mode)
                audio_bytes = generate_edge_tts_audio_sync(speech_text, voice, self.rate)
                if audio_bytes and not self.is_cancelled:
                    self.cache_dict[cache_key] = audio_bytes
            except Exception:
                pass

class NeuralSpeechWorker(QThread):
    def __init__(self, text, voice="en-IN-NeerjaNeural", rate="+25%", fallback_speaker=None, on_ready_callback=None):
        super().__init__()
        self.text = text
        self.voice = voice
        self.rate = rate
        self.fallback_speaker = fallback_speaker
        self.on_ready_callback = on_ready_callback
        self.is_cancelled = False

    def run(self):
        if self.is_cancelled:
            return
        
        # 1. Ultra-Natural Neural AI Human Voice (Edge Neural Studio TTS)
        if EDGE_TTS_AVAILABLE:
            try:
                audio_bytes = generate_edge_tts_audio_sync(self.text, self.voice, self.rate)
                if audio_bytes and not self.is_cancelled:
                    if self.on_ready_callback:
                        self.on_ready_callback(audio_bytes)
                    if not pygame.mixer.get_init():
                        pygame.mixer.init()
                    pygame.mixer.music.stop()
                    sound_io = io.BytesIO(audio_bytes)
                    pygame.mixer.music.load(sound_io)
                    pygame.mixer.music.play()
                    while pygame.mixer.music.get_busy() and not self.is_cancelled:
                        pygame.time.Clock().tick(10)
                    return
            except Exception as e:
                print("Neural AI TTS fallback to SAPI:", e)

        # 2. Offline fallback (Windows SAPI5)
        if not self.is_cancelled and self.fallback_speaker:
            try:
                clean = re.sub(r'[\*\#\_]', '', self.text)
                self.fallback_speaker.Speak(clean, 3)
            except Exception:
                pass

    def stop(self):
        self.is_cancelled = True
        if EDGE_TTS_AVAILABLE:
            try:
                if pygame.mixer.get_init():
                    pygame.mixer.stop()
            except Exception:
                pass
        if self.fallback_speaker:
            try:
                self.fallback_speaker.Speak("", 2)
            except Exception:
                pass

def play_system_sfx(name):
    """Plays synthesized high-tech audio chime / feedback without external files."""
    if not pygame:
        return
    try:
        if not pygame.mixer.get_init():
            pygame.mixer.init(frequency=44100, size=-16, channels=2, buffer=512)
        import numpy as np
        sample_rate = 44100
        def make_tone(freq, dur=0.1, vol=0.25):
            n_samples = int(sample_rate * dur)
            t = np.linspace(0, dur, n_samples, False)
            env = np.sin(np.pi * np.linspace(0, 1, n_samples)) ** 2
            wave = np.sin(2 * np.pi * freq * t) * env * vol
            stereo = np.column_stack((wave, wave))
            return pygame.sndarray.make_sound((stereo * 32767).astype(np.int16))
            
        if name == "mic_on":
            s1 = make_tone(587, 0.08, 0.22)
            s2 = make_tone(880, 0.12, 0.22)
            s1.play()
            pygame.time.delay(65)
            s2.play()
        elif name == "mic_off":
            s1 = make_tone(784, 0.07, 0.18)
            s2 = make_tone(440, 0.1, 0.18)
            s1.play()
            pygame.time.delay(55)
            s2.play()
        elif name == "correct":
            s1 = make_tone(659, 0.07, 0.25)
            s2 = make_tone(880, 0.14, 0.28)
            s1.play()
            pygame.time.delay(60)
            s2.play()
        elif name == "wrong":
            s = make_tone(220, 0.18, 0.2)
            s.play()
    except Exception:
        pass

def match_option_by_content(text, question_dict):
    """Intelligently matches user spoken text with the actual wording of options."""
    if not text or not question_dict:
        return None
    t = text.lower().strip()
    t = re.sub(r'[^\w\s]', ' ', t)
    letters = ["A", "B", "C", "D"]
    for idx, letter in enumerate(letters):
        opt_val = question_dict.get(f"option{letter}", "").strip().lower()
        if not opt_val:
            continue
        eng_part, bn_part = split_english_bengali(opt_val)
        for part in [opt_val, eng_part, bn_part]:
            if not part:
                continue
            clean = re.sub(r'[^\w\s]', ' ', part).strip()
            if not clean:
                continue
            # Direct inclusion if length >= 3
            if len(clean) >= 3 and (clean in t or t in clean):
                return idx
            # Match any significant word (length >= 4)
            for w in clean.split():
                if len(w) >= 4 and w in t.split():
                    return idx
    return None

def parse_voice_command(text, current_question=None):
    """Parses spoken user utterance into MCQ choice index or rich conversational commands."""
    if not text:
        return None, None
    t = text.lower().strip()
    t = re.sub(r'[^\w\s]', ' ', t)
    words = t.split()
    if not words or len(words) > 7:
        return None, None

    # 1. NAVIGATION COMMANDS
    if any(p in t for p in ['next', 'agla', 'skip', 'porer', 'chalo', 'age bado', 'next question', 'porer proshno']):
        return ('next', None)
    if any(p in t for p in ['previous', 'pichla', 'back', 'agerta', 'prev', 'pichle']):
        return ('prev', None)

    # 2. MARK / REVIEW / CLEAR
    if any(p in t for p in ['mark', 'review', 'star', 'tag']):
        return ('mark', None)
    if any(p in t for p in ['clear', 'reset', 'remove', 'muche dao']):
        return ('clear', None)

    # 3. SOLUTION / DETAILS
    if any(p in t for p in ['detail', 'details', 'solution', 'explain', 'explanation', 'somadhan', 'bistarito']):
        return ('details', None)

    # 4. SPEED CONTROLS
    if any(p in t for p in ['faster', 'speed up', 'fast', 'tez', 'jaldi']):
        return ('speed_up', None)
    if any(p in t for p in ['slower', 'slow down', 'slow', 'dheere', 'dhire', 'aste']):
        return ('speed_down', None)

    # 5. LANGUAGE SWITCHING BY VOICE
    if 'bangla' in t or 'bengali' in t:
        return ('lang_bn', None)
    if 'english' in t:
        return ('lang_en', None)
    if 'both' in t:
        return ('lang_both', None)

    # 6. EXPLICIT OPTION PHRASES ("option A", "option B", "option C", "option D")
    if 'option a' in t or 'option 1' in t or 'ans a' in t or 'answer a' in t or 'a wala' in t or 'a number' in t:
        return ('option', 0)
    if 'option b' in t or 'option 2' in t or 'ans b' in t or 'answer b' in t or 'b wala' in t or 'b number' in t:
        return ('option', 1)
    if 'option c' in t or 'option 3' in t or 'ans c' in t or 'answer c' in t or 'c wala' in t or 'c number' in t:
        return ('option', 2)
    if 'option d' in t or 'option 4' in t or 'ans d' in t or 'answer d' in t or 'd wala' in t or 'd number' in t:
        return ('option', 3)

    # 7. WORD-BASED MATCHING (Exact whole words only)
    words_set = set(words)
    if words_set.intersection({'a', '1', 'one', 'first', 'ek', 'prothom', 'apple'}):
        return ('option', 0)
    if words_set.intersection({'b', '2', 'two', 'second', 'do', 'dui', 'ditiyo', 'ball', 'boy', 'be'}):
        return ('option', 1)
    if words_set.intersection({'c', '3', 'three', 'third', 'teen', 'tritiyo', 'cat', 'see', 'sea'}):
        return ('option', 2)
    if words_set.intersection({'d', '4', 'four', 'fourth', 'char', 'choturtho', 'dog', 'delhi'}):
        return ('option', 3)

    # 8. CONTENT-BASED SMART OPTION MATCHING (Speaks the name of the answer!)
    if current_question:
        idx = match_option_by_content(text, current_question)
        if idx is not None:
            return ('option', idx)

    return (None, None)

def extract_voice_features(raw_audio_bytes, sample_rate=16000):
    """Extracts lightweight acoustic biometric features from raw PCM audio."""
    if not raw_audio_bytes or len(raw_audio_bytes) < 512:
        return None
    try:
        import numpy as np
        samples = np.frombuffer(raw_audio_bytes, dtype=np.int16).astype(np.float32)
        if len(samples) < 512:
            return None
        samples = samples - np.mean(samples)
        rms = np.sqrt(np.mean(samples**2))
        if rms < 150: # Ignore silence / low sound
            return None
        zcr = np.mean(np.abs(np.diff(np.sign(samples)))) / 2.0
        fft_vals = np.abs(np.fft.rfft(samples * np.hanning(len(samples))))
        freqs = np.fft.rfftfreq(len(samples), 1.0 / sample_rate)
        vocal_mask = (freqs >= 80) & (freqs <= 3500)
        vocal_fft = fft_vals[vocal_mask]
        vocal_freqs = freqs[vocal_mask]
        if len(vocal_fft) == 0 or np.sum(vocal_fft) == 0:
            return None
        spectral_centroid = float(np.sum(vocal_freqs * vocal_fft) / np.sum(vocal_fft))
        dominant_pitch = float(vocal_freqs[np.argmax(vocal_fft)])
        band_low = float(np.sum(fft_vals[(freqs >= 80) & (freqs < 500)]))
        band_mid = float(np.sum(fft_vals[(freqs >= 500) & (freqs < 1500)]))
        band_high = float(np.sum(fft_vals[(freqs >= 1500) & (freqs <= 3500)]))
        total_band = band_low + band_mid + band_high + 1e-6
        return {
            "pitch": dominant_pitch,
            "centroid": spectral_centroid,
            "zcr": float(zcr),
            "band_low_ratio": band_low / total_band,
            "band_mid_ratio": band_mid / total_band,
            "band_high_ratio": band_high / total_band
        }
    except Exception:
        return None

def calculate_speaker_similarity(feat1, feat2):
    """Calculates biometric similarity score [0.0 to 1.0] between two voiceprints."""
    if not feat1 or not feat2:
        return 1.0
    try:
        pitch_diff = abs(feat1["pitch"] - feat2["pitch"])
        pitch_sim = max(0.0, 1.0 - (pitch_diff / 200.0))
        centroid_diff = abs(feat1["centroid"] - feat2["centroid"])
        centroid_sim = max(0.0, 1.0 - (centroid_diff / 600.0))
        ratio_diff = (abs(feat1["band_low_ratio"] - feat2["band_low_ratio"]) +
                      abs(feat1["band_mid_ratio"] - feat2["band_mid_ratio"]) +
                      abs(feat1["band_high_ratio"] - feat2["band_high_ratio"])) / 2.0
        ratio_sim = max(0.0, 1.0 - ratio_diff)
        return (0.35 * pitch_sim) + (0.35 * centroid_sim) + (0.30 * ratio_sim)
    except Exception:
        return 1.0

class SpeakerVoiceprintManager:
    def __init__(self, profile_path="speaker_profile.json"):
        self.profile_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), profile_path)
        self.profile = self.load_profile()

    def load_profile(self):
        try:
            if os.path.exists(self.profile_path):
                with open(self.profile_path, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return None

    def save_profile(self, features):
        if not features:
            return
        try:
            with open(self.profile_path, "w", encoding="utf-8") as f:
                json.dump(features, f, indent=2)
            self.profile = features
        except Exception as e:
            print("Failed to save speaker profile:", e)

    def verify_or_learn(self, raw_audio_bytes, sample_rate=16000):
        feats = extract_voice_features(raw_audio_bytes, sample_rate)
        if not feats:
            return True # Can't extract (e.g. too short), allow processing
            
        if not self.profile:
            # Auto-learn on first successful user command
            self.save_profile(feats)
            print("[Voice Biometrics] Auto-learned user's voice signature!")
            return True
            
        sim = calculate_speaker_similarity(self.profile, feats)
        
        if sim >= 0.50:
            # Match confirmed! Adaptively blend profile (92% existing + 8% new)
            updated = {
                k: float(0.92 * self.profile[k] + 0.08 * feats[k])
                for k in feats
            }
            self.save_profile(updated)
            return True
        else:
            print(f"[Voice Biometrics] Rejected background speaker (Confidence: {sim*100:.1f}%)")
            return False

class VoiceAnswerListener(QObject):
    command_detected = pyqtSignal(str, object, str)  # cmd_type, value, raw_text
    heard_text = pyqtSignal(str)                     # raw recognized text
    partial_text = pyqtSignal(str)                   # live real-time partial text
    status_changed = pyqtSignal(str)                 # 'listening', 'processing', 'idle', 'error'
    
    def __init__(self, lang_mode="both", is_speaking_fn=None, get_q_fn=None, parent=None):
        super().__init__(parent)
        self.lang_mode = lang_mode
        self.is_speaking_fn = is_speaking_fn
        self.get_q_fn = get_q_fn
        self.is_active = False
        self.stop_fn = None
        self.sd_stream = None
        self.worker_thread = None
        self.audio_queue = queue.Queue()
        self.voiceprint_mgr = SpeakerVoiceprintManager()
        self.vosk_model = None
        
        # Initialize Neural Engine
        if VOSK_AVAILABLE and vosk_module:
            try:
                self.vosk_model = vosk_module.Model(lang="en-us")
            except Exception as e:
                print("Vosk model init error:", e)
                self.vosk_model = None

    def start(self):
        self.stop()
        self.is_active = True
        
        # Priority 1: High-Speed Neural Real-Time Stream (Vosk + SoundDevice)
        if VOSK_AVAILABLE and SOUNDDEVICE_AVAILABLE and self.vosk_model and sounddevice_module:
            try:
                grammar = [
                    "option a", "option b", "option c", "option d",
                    "option one", "option two", "option three", "option four",
                    "option 1", "option 2", "option 3", "option 4",
                    "a", "b", "c", "d", "one", "two", "three", "four", "first", "second", "third", "fourth",
                    "next", "previous", "back", "details", "solution", "explain", "mark", "clear",
                    "faster", "slower", "english", "bengali", "[unk]"
                ]
                try:
                    self.kaldi_rec = vosk_module.KaldiRecognizer(self.vosk_model, 16000, json.dumps(grammar))
                except Exception:
                    self.kaldi_rec = vosk_module.KaldiRecognizer(self.vosk_model, 16000)
                    
                self.audio_queue = queue.Queue()
                
                def sd_callback(indata, frames, time_info, status):
                    if not self.is_active:
                        return
                    if self.is_speaking_fn and self.is_speaking_fn():
                        return
                    self.audio_queue.put(bytes(indata))

                self.sd_stream = sounddevice_module.RawInputStream(
                    samplerate=16000,
                    blocksize=3200,
                    dtype='int16',
                    channels=1,
                    callback=sd_callback
                )
                self.sd_stream.start()
                self.status_changed.emit("listening")
                
                self.worker_thread = threading.Thread(target=self._process_neural_stream, daemon=True)
                self.worker_thread.start()
                print("[Voice Assistant] Neural Real-Time Streaming Engine Active! 🚀")
                return
            except Exception as ex:
                print("[Voice Assistant] SoundDevice stream init failed, falling back to SpeechRecognition:", ex)

        # Priority 2: Google SpeechRecognition Fallback
        if SPEECH_RECOG_AVAILABLE and speech_recognition:
            try:
                self.recognizer = speech_recognition.Recognizer()
                init_thresh = 200
                if self.voiceprint_mgr.profile:
                    init_thresh = self.voiceprint_mgr.profile.get("calibrated_energy_threshold", 200)
                self.recognizer.energy_threshold = init_thresh
                self.recognizer.dynamic_energy_threshold = True
                self.recognizer.dynamic_energy_adjustment_damping = 0.15
                self.recognizer.dynamic_energy_ratio = 1.8
                self.recognizer.pause_threshold = 0.5
                self.recognizer.non_speaking_duration = 0.4
                self.recognizer.phrase_threshold = 0.10
                self.microphone = speech_recognition.Microphone()
                
                self.status_changed.emit("listening")
                recog_lang = "bn-IN" if self.lang_mode == "bn" else "en-IN"
                
                def audio_callback(recognizer, audio):
                    if not self.is_active:
                        return
                    if self.is_speaking_fn and self.is_speaking_fn():
                        return
                    self.status_changed.emit("processing")
                    try:
                        text = recognizer.recognize_google(audio, language=recog_lang)
                        if text and self.is_active:
                            print(f"[Voice Assistant] Heard: '{text}'")
                            self.heard_text.emit(text)
                            current_q = self.get_q_fn() if self.get_q_fn else None
                            cmd, val = parse_voice_command(text, current_question=current_q)
                            if cmd and self.is_active:
                                self.command_detected.emit(cmd, val, text)
                    except speech_recognition.UnknownValueError:
                        pass
                    except Exception as ex:
                        print("[Voice Assistant] Speech recognition error:", ex)
                    finally:
                        if self.is_active:
                            self.status_changed.emit("listening")

                self.stop_fn = self.recognizer.listen_in_background(
                    self.microphone,
                    audio_callback,
                    phrase_time_limit=4.0
                )
            except Exception as ex:
                print("Mic fallback error:", ex)
                self.status_changed.emit("error")
        else:
            self.status_changed.emit("error")

    def _process_neural_stream(self):
        while self.is_active:
            try:
                data = self.audio_queue.get(timeout=0.1)
            except queue.Empty:
                continue
            if not self.is_active:
                break
                
            if self.kaldi_rec.AcceptWaveform(data):
                res = json.loads(self.kaldi_rec.Result())
                text = res.get("text", "").strip()
                if text and text != "[unk]":
                    print(f"[Neural Voice Engine] Final Heard: '{text}'")
                    self.heard_text.emit(text)
                    current_q = self.get_q_fn() if self.get_q_fn else None
                    cmd, val = parse_voice_command(text, current_question=current_q)
                    if cmd and self.is_active:
                        self.command_detected.emit(cmd, val, text)
            else:
                part = json.loads(self.kaldi_rec.PartialResult())
                partial_str = part.get("partial", "").strip()
                if partial_str and partial_str != "[unk]":
                    self.partial_text.emit(partial_str)

    def stop(self):
        self.is_active = False
        if self.sd_stream:
            try:
                self.sd_stream.stop()
                self.sd_stream.close()
            except Exception:
                pass
            self.sd_stream = None
        if self.stop_fn:
            try:
                self.stop_fn(wait_for_stop=False)
            except Exception:
                pass
            self.stop_fn = None
        self.status_changed.emit("idle")
        self.status_changed.emit("idle")

class QuizPlayerDialog(QDialog):
    def __init__(self, questions, parent=None):
        super().__init__(parent)
        self.questions = questions
        self.current_idx = 0
        self.score = 0
        self.seconds_elapsed = 0
        self.user_answers = {}
        self.marked_for_review = set()
        
        # Real Human Neural AI Voice System (Bilingual: English + Bengali + Zero Latency Pre-cache)
        self.voice_enabled = False
        self.language_mode = 'both'  # 'both', 'en', 'bn'
        self.voice_speed_percent = 25  # +25% for brisk natural pacing
        self.voice_en = "en-IN-NeerjaNeural"    # Studio English Neural Voice
        self.voice_bn = "bn-IN-TanishaaNeural"  # Studio Bengali Neural Voice
        self.voice_thread = None
        self.preloader_thread = None
        self.audio_cache = {}  # (q_idx, lang_mode) -> audio_bytes for instant playback
        
        # Microphone Voice-Answer System [Shift]
        self.mic_enabled = False
        self.mic_thread = None
        
        self.speaker = None
        if WIN32_SAPI_AVAILABLE and win32com_client:
            try:
                self.speaker = win32com_client.Dispatch("SAPI.SpVoice")
            except Exception as ex:
                print("SAPI Voice fallback init error:", ex)
        
        # Start preloading initial questions right away
        self.start_preload(0)
        
        self.setWindowTitle("🏆 Laptop CBT Exam Portal - Test Player")
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMinMaxButtonsHint)
        self.resize(1160, 750)
        
        self.setStyleSheet("""
            QDialog {
                background-color: #0F0F11;
                color: #FFFFFF;
                font-family: 'Segoe UI', sans-serif;
            }
            QScrollArea {
                background-color: #141416;
                border: none;
            }
            QWidget#scrollWidget {
                background-color: #141416;
            }
            QFrame#q_box {
                background-color: #1A1A1E;
                border: 1px solid #2C2C34;
                border-radius: 12px;
            }
            QFrame#side_panel {
                background-color: #17171C;
                border: 1px solid #282830;
                border-radius: 12px;
            }
            QLabel#lbl_q_title {
                color: #00BCD4;
                font-size: 17px;
                font-weight: bold;
            }
            QLabel#lbl_question {
                color: #FFFFFF;
                font-size: 16px;
                font-weight: 600;
                line-height: 1.5;
            }
            QPushButton#opt_btn {
                background-color: #202026;
                border: 1px solid #333340;
                border-radius: 10px;
                padding: 12px 18px;
                color: #E2E8F0;
                text-align: left;
                font-size: 14px;
                font-weight: 500;
            }
            QPushButton#opt_btn:hover {
                background-color: #2A2A34;
                border: 1px solid #0078D4;
                color: #FFFFFF;
            }
            QPushButton#opt_btn:disabled {
                color: #CBD5E1;
            }
            QPushButton#save_next_btn {
                background-color: #0078D4;
                border: 1px solid #005A9E;
                border-radius: 8px;
                padding: 10px 22px;
                color: #FFFFFF;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton#save_next_btn:hover {
                background-color: #106EBE;
            }
            QPushButton#prev_btn {
                background-color: #282830;
                border: 1px solid #3A3A46;
                border-radius: 8px;
                padding: 10px 20px;
                color: #FFFFFF;
                font-weight: 600;
                font-size: 13px;
            }
            QPushButton#prev_btn:hover {
                background-color: #363642;
            }
            QPushButton#mark_btn {
                background-color: #6A1B9A;
                border: 1px solid #8E24AA;
                border-radius: 8px;
                padding: 10px 18px;
                color: #FFFFFF;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton#mark_btn:hover {
                background-color: #7B1FA2;
            }
            QPushButton#clear_btn {
                background-color: #374151;
                border: 1px solid #4B5563;
                border-radius: 8px;
                padding: 10px 18px;
                color: #F3F4F6;
                font-weight: 600;
                font-size: 13px;
            }
            QPushButton#clear_btn:hover {
                background-color: #4B5563;
            }
            QPushButton#submit_test_btn {
                background-color: #2E7D32;
                border: 1px solid #388E3C;
                border-radius: 8px;
                padding: 12px;
                color: #FFFFFF;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton#submit_test_btn:hover {
                background-color: #388E3C;
            }
            QPushButton#overview_btn {
                background-color: #1E293B;
                border: 1px solid #00BCD4;
                border-radius: 8px;
                padding: 6px 14px;
                color: #00BCD4;
                font-weight: bold;
                font-size: 13px;
            }
            QPushButton#overview_btn:hover {
                background-color: #00BCD4;
                color: #0F0F11;
            }
            QFrame#solCard {
                background-color: #132218;
                border: 1px solid #2E7D32;
                border-radius: 10px;
            }
            QLabel#solTitle {
                color: #81C784;
                font-weight: bold;
                font-size: 14px;
            }
            QLabel#solText {
                color: #E8F5E9;
                font-size: 13px;
            }
            QScrollArea#grid_scroll {
                background-color: #17171C;
                border: none;
            }
        """)

        # Main Layout
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(14, 12, 14, 12)
        self.main_layout.setSpacing(10)

        # 1. Top Exam Header Bar (PW / Testbook Style)
        self.header_layout = QHBoxLayout()
        lbl_exam_title = QLabel("🏆 CBT MOCK TEST PORTAL")
        lbl_exam_title.setStyleSheet("font-size: 16px; font-weight: bold; color: #00BCD4; letter-spacing: 0.5px;")
        
        lbl_subject = QLabel("•  General Awareness / Laptop Quiz")
        lbl_subject.setStyleSheet("font-size: 13px; color: #94A3B8; font-weight: 500;")

        self.lbl_timer = QLabel("⏱️ Time: 00:00")
        self.lbl_timer.setStyleSheet("font-size: 15px; font-weight: bold; color: #FF9800; background: #261E14; border: 1px solid #FF9800; border-radius: 6px; padding: 4px 10px;")

        lbl_marking = QLabel("🎯 Marks: +1.00, -0.25")
        lbl_marking.setStyleSheet("font-size: 13px; color: #4CAF50; background: #132218; border: 1px solid #2E7D32; border-radius: 6px; padding: 4px 10px;")

        # Voice Narration Toggle Button
        self.btn_voice = QPushButton("🔇 Voice: OFF [V]")
        self.btn_voice.setObjectName("overview_btn")
        self.btn_voice.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_voice.setStyleSheet("background-color: #2D2D38; border: 1px solid #454556; color: #CBD5E1; font-weight: bold; font-size: 12px;")
        self.btn_voice.setToolTip("Toggle Voice Narration [V]\nSpeed Controls: Numpad '+' (Fast) / '-' (Slow)\nRead Solution: Press 'Ctrl'")
        self.btn_voice.clicked.connect(self.toggle_voice_mode)

        # Language Mode Switcher (Both / English / Bangla)
        self.btn_lang = QPushButton("🌐 Lang: Both [L]")
        self.btn_lang.setObjectName("overview_btn")
        self.btn_lang.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_lang.setStyleSheet("background-color: #1E293B; border: 1px solid #38BDF8; color: #38BDF8; font-weight: bold; font-size: 12px;")
        self.btn_lang.setToolTip("Switch Voice Language: Both / English / Bangla (Press 'L')")
        self.btn_lang.clicked.connect(self.cycle_language_mode)

        # Microphone Voice-Answer Mode Toggle Button [S / Shift]
        self.btn_mic = QPushButton("🎙️ Mic: OFF [S / Shift]")
        self.btn_mic.setObjectName("overview_btn")
        self.btn_mic.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_mic.setStyleSheet("background-color: #2D2D38; border: 1px solid #454556; color: #CBD5E1; font-weight: bold; font-size: 12px;")
        self.btn_mic.setToolTip("Toggle Microphone Speech Answer [Press 'S' or 'Shift']\nSpeak: 'Option A', 'B', '1', '2', '3', '4', 'Next', etc.")
        self.btn_mic.clicked.connect(self.toggle_mic_mode)

        self.btn_overview = QPushButton("📊 Questions Chart (A4)")
        self.btn_overview.setObjectName("overview_btn")
        self.btn_overview.setToolTip("View full A4 summary, table of all questions, answers, and important details")
        self.btn_overview.clicked.connect(self.open_overview)

        self.btn_player_html = QPushButton("📥 HTML")
        self.btn_player_html.setObjectName("overview_btn")
        self.btn_player_html.setStyleSheet("background-color: #10B981; border: 1px solid #059669; color: #FFFFFF; font-weight: bold;")
        self.btn_player_html.setToolTip("Download responsive HTML file for Mobile Phone / Web viewing")
        self.btn_player_html.clicked.connect(self.export_html)

        self.header_layout.addWidget(lbl_exam_title)
        self.header_layout.addWidget(lbl_subject)
        self.header_layout.addStretch()
        self.header_layout.addWidget(lbl_marking)
        self.header_layout.addSpacing(6)
        self.header_layout.addWidget(self.btn_voice)
        self.header_layout.addSpacing(6)
        self.header_layout.addWidget(self.btn_lang)
        self.header_layout.addSpacing(6)
        self.header_layout.addWidget(self.btn_mic)
        self.header_layout.addSpacing(8)
        self.header_layout.addWidget(self.lbl_timer)
        self.header_layout.addSpacing(8)
        self.header_layout.addWidget(self.btn_overview)
        self.header_layout.addSpacing(6)
        self.header_layout.addWidget(self.btn_player_html)
        self.main_layout.addLayout(self.header_layout)

        # 2. Split Body (Left: Question Arena, Right: Side Palette)
        self.split_layout = QHBoxLayout()
        self.split_layout.setSpacing(12)

        # Left: Main Question Arena (~72% Width)
        self.left_widget = QWidget()
        self.left_layout = QVBoxLayout(self.left_widget)
        self.left_layout.setContentsMargins(0, 0, 0, 0)
        self.left_layout.setSpacing(10)

        # Arena Subheader
        arena_hdr = QHBoxLayout()
        self.lbl_q_title = QLabel("Question 1")
        self.lbl_q_title.setObjectName("lbl_q_title")
        
        lbl_hint_keys = QLabel("⌨️ Shortcuts: [1, 2, 3, 4] to Answer  |  [Space/Enter] Next  |  [M] Mark  |  [C] Clear")
        lbl_hint_keys.setStyleSheet("color: #64748B; font-size: 11px; font-weight: 500;")

        arena_hdr.addWidget(self.lbl_q_title)
        arena_hdr.addStretch()
        arena_hdr.addWidget(lbl_hint_keys)
        self.left_layout.addLayout(arena_hdr)

        # Scroll Area for Question & Options
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_widget = QWidget()
        self.scroll_widget.setObjectName("scrollWidget")
        self.scroll_layout = QVBoxLayout(self.scroll_widget)
        self.scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.scroll_layout.setSpacing(12)

        # Question Card Box
        self.q_card = QFrame()
        self.q_card.setObjectName("q_box")
        self.q_card_layout = QVBoxLayout(self.q_card)
        self.q_card_layout.setContentsMargins(18, 16, 18, 16)
        self.q_card_layout.setSpacing(10)
        
        self.lbl_question = QLabel("Question Text")
        self.lbl_question.setObjectName("lbl_question")
        self.lbl_question.setWordWrap(True)
        self.q_card_layout.addWidget(self.lbl_question)

        self.lbl_image = QLabel()
        self.lbl_image.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_image.hide()
        self.q_card_layout.addWidget(self.lbl_image)

        self.scroll_layout.addWidget(self.q_card)

        # Live Microphone Voice-Answer Status Banner
        self.lbl_mic_banner = QLabel("🎙️ Voice Assistant Active: Say 'Option A', 'B', '1', '2', or 'Next' ...")
        self.lbl_mic_banner.setStyleSheet("""
            background-color: #1E1B2E;
            border: 1px solid #7C3AED;
            border-radius: 8px;
            color: #DDD6FE;
            font-size: 13px;
            font-weight: 600;
            padding: 8px 14px;
        """)
        self.lbl_mic_banner.hide()
        self.scroll_layout.addWidget(self.lbl_mic_banner)

        # Option Buttons (A, B, C, D) with [1, 2, 3, 4] Badges
        self.opt_buttons = []
        for i, letter in enumerate(["A", "B", "C", "D"]):
            btn = QPushButton()
            btn.setObjectName("opt_btn")
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.clicked.connect(lambda checked, idx=i: self.check_answer(idx))
            self.opt_buttons.append(btn)
            self.scroll_layout.addWidget(btn)

        # Solution / Explanation Panel
        self.sol_card = QFrame()
        self.sol_card.setObjectName("solCard")
        self.sol_layout = QVBoxLayout(self.sol_card)
        self.sol_layout.setContentsMargins(14, 12, 14, 12)
        self.lbl_sol_title = QLabel("💡 Explanation / Solution:")
        self.lbl_sol_title.setObjectName("solTitle")
        self.lbl_solution = QLabel("Solution details go here...")
        self.lbl_solution.setObjectName("solText")
        self.lbl_solution.setWordWrap(True)
        self.sol_layout.addWidget(self.lbl_sol_title)
        self.sol_layout.addWidget(self.lbl_solution)
        self.sol_card.hide()
        self.scroll_layout.addWidget(self.sol_card)

        self.scroll_area.setWidget(self.scroll_widget)
        self.left_layout.addWidget(self.scroll_area)

        # Bottom Action Bar (CBT Standard)
        self.action_bar = QHBoxLayout()
        self.btn_mark = QPushButton("⭐ Mark for Review")
        self.btn_mark.setObjectName("mark_btn")
        self.btn_mark.setToolTip("Mark this question to review later (Shortcut: M)")
        self.btn_mark.clicked.connect(self.toggle_mark_review)

        self.btn_clear = QPushButton("🧹 Clear Response")
        self.btn_clear.setObjectName("clear_btn")
        self.btn_clear.setToolTip("Clear selected answer (Shortcut: C or Backspace)")
        self.btn_clear.clicked.connect(self.clear_response)

        self.btn_prev = QPushButton("⬅️ Previous")
        self.btn_prev.setObjectName("prev_btn")
        self.btn_prev.setToolTip("Go to previous question (Shortcut: Left Arrow)")
        self.btn_prev.clicked.connect(self.prev_question)

        self.btn_next = QPushButton("Save & Next ➡️")
        self.btn_next.setObjectName("save_next_btn")
        self.btn_next.setToolTip("Save answer and move to next question (Shortcut: Space or Enter)")
        self.btn_next.clicked.connect(self.next_question)

        self.action_bar.addWidget(self.btn_mark)
        self.action_bar.addWidget(self.btn_clear)
        self.action_bar.addStretch()
        self.action_bar.addWidget(self.btn_prev)
        self.action_bar.addWidget(self.btn_next)
        self.left_layout.addLayout(self.action_bar)

        self.split_layout.addWidget(self.left_widget, 72)

        # Right: Testbook / PhysicsWallah Style Side Palette (~28% Width)
        self.side_frame = QFrame()
        self.side_frame.setObjectName("side_panel")
        self.side_layout = QVBoxLayout(self.side_frame)
        self.side_layout.setContentsMargins(12, 12, 12, 12)
        self.side_layout.setSpacing(10)

        # Palette Title
        lbl_palette_title = QLabel("📌 Question Palette")
        lbl_palette_title.setStyleSheet("font-size: 15px; font-weight: bold; color: #FFFFFF;")
        self.side_layout.addWidget(lbl_palette_title)

        # Legend Status Badges Grid
        self.lbl_legend_ans = QLabel("🟢 Answered: 0")
        self.lbl_legend_ans.setStyleSheet("color: #4CAF50; font-size: 12px; font-weight: bold;")
        self.lbl_legend_wrong = QLabel("🔴 Wrong/Unsaved: 0")
        self.lbl_legend_wrong.setStyleSheet("color: #F44336; font-size: 12px; font-weight: bold;")
        self.lbl_legend_marked = QLabel("🟣 Marked: 0")
        self.lbl_legend_marked.setStyleSheet("color: #BA68C8; font-size: 12px; font-weight: bold;")
        self.lbl_legend_unvis = QLabel(f"⚪ Not Visited: {len(self.questions)}")
        self.lbl_legend_unvis.setStyleSheet("color: #94A3B8; font-size: 12px;")

        legend_grid = QHBoxLayout()
        v1 = QVBoxLayout(); v1.addWidget(self.lbl_legend_ans); v1.addWidget(self.lbl_legend_marked)
        v2 = QVBoxLayout(); v2.addWidget(self.lbl_legend_wrong); v2.addWidget(self.lbl_legend_unvis)
        legend_grid.addLayout(v1); legend_grid.addLayout(v2)
        self.side_layout.addLayout(legend_grid)

        # Separator
        line = QFrame(); line.setFrameShape(QFrame.Shape.HLine); line.setStyleSheet("color: #2C2C34;")
        self.side_layout.addWidget(line)

        # Grid of Number Buttons
        self.grid_scroll = QScrollArea()
        self.grid_scroll.setObjectName("grid_scroll")
        self.grid_scroll.setWidgetResizable(True)
        self.grid_widget = QWidget()
        self.grid_layout = QVBoxLayout(self.grid_widget)
        self.grid_layout.setContentsMargins(4, 4, 4, 4)
        self.grid_layout.setSpacing(6)

        # Place buttons in rows of 5
        self.palette_btns = []
        current_row_layout = None
        for i in range(len(self.questions)):
            if i % 5 == 0:
                current_row_layout = QHBoxLayout()
                current_row_layout.setSpacing(6)
                self.grid_layout.addLayout(current_row_layout)

            q_btn = QPushButton(str(i + 1))
            q_btn.setFixedSize(38, 34)
            q_btn.setCursor(Qt.CursorShape.PointingHandCursor)
            q_btn.clicked.connect(lambda _, idx=i: self.jump_to_question(idx))
            current_row_layout.addWidget(q_btn)
            self.palette_btns.append(q_btn)

        # Pad last row if needed
        if current_row_layout and len(self.questions) % 5 != 0:
            current_row_layout.addStretch()

        self.grid_layout.addStretch()
        self.grid_scroll.setWidget(self.grid_widget)
        self.side_layout.addWidget(self.grid_scroll)

        # Submit Test Button at bottom of Side Palette
        self.btn_submit_test = QPushButton("🏁 Submit Test")
        self.btn_submit_test.setObjectName("submit_test_btn")
        self.btn_submit_test.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_submit_test.clicked.connect(self.confirm_submit)
        self.side_layout.addWidget(self.btn_submit_test)

        self.split_layout.addWidget(self.side_frame, 28)
        self.main_layout.addLayout(self.split_layout)

        # Timer setup
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_timer)
        self.timer.start(1000)

        # Load first question
        self.load_question()

    def start_preload(self, start_idx):
        if not EDGE_TTS_AVAILABLE:
            return
        rate_str = f"+{self.voice_speed_percent}%" if self.voice_speed_percent >= 0 else f"{self.voice_speed_percent}%"
        if self.preloader_thread and self.preloader_thread.isRunning():
            self.preloader_thread.is_cancelled = True
        self.preloader_thread = BackgroundAudioPreloader(self.questions, start_idx, self.language_mode, rate_str, self.audio_cache)
        self.preloader_thread.start()

    def stop_speech(self):
        if self.voice_thread:
            try:
                self.voice_thread.stop()
            except Exception:
                pass
        if EDGE_TTS_AVAILABLE:
            try:
                if pygame.mixer.get_init():
                    pygame.mixer.music.stop()
            except Exception:
                pass
        if self.speaker:
            try:
                self.speaker.Speak("", 2) # SVSFPurgeBeforeSpeak
            except Exception:
                pass

    def speak_text(self, text, voice=None, purge=True, on_ready=None):
        if not self.voice_enabled:
            return
        if purge:
            self.stop_speech()
        if not voice:
            if re.search(r'[\u0980-\u09FF]', text):
                voice = self.voice_bn
            else:
                voice = self.voice_en if self.language_mode == 'en' else self.voice_bn
        rate_str = f"+{self.voice_speed_percent}%" if self.voice_speed_percent >= 0 else f"{self.voice_speed_percent}%"
        self.voice_thread = NeuralSpeechWorker(text, voice=voice, rate=rate_str, fallback_speaker=self.speaker, on_ready_callback=on_ready)
        self.voice_thread.start()

    def toggle_voice_mode(self):
        self.voice_enabled = not self.voice_enabled
        self.stop_speech()
        self.update_voice_btn_ui()
        if self.voice_enabled:
            self.start_preload(self.current_idx)
            if self.current_idx not in self.user_answers:
                self.narrate_current_question()

    def cycle_language_mode(self):
        if self.language_mode == 'both':
            self.language_mode = 'en'
        elif self.language_mode == 'en':
            self.language_mode = 'bn'
        else:
            self.language_mode = 'both'
        self.update_lang_btn_ui()
        self.start_preload(self.current_idx) # pre-cache in new language!
        if self.voice_enabled:
            self.stop_speech()
            if self.language_mode == 'en':
                self.speak_text("Language set to English.", voice=self.voice_en, purge=True)
            elif self.language_mode == 'bn':
                self.speak_text("ভাষা বাংলা সেট করা হয়েছে।", voice=self.voice_bn, purge=True)
            else:
                self.speak_text("Language set to English and Bangla.", voice=self.voice_bn, purge=True)

    def update_lang_btn_ui(self):
        if self.language_mode == 'both':
            self.btn_lang.setText("🌐 Lang: Both [L]")
            self.btn_lang.setStyleSheet("background-color: #1E293B; border: 1px solid #38BDF8; color: #38BDF8; font-weight: bold; font-size: 12px;")
        elif self.language_mode == 'en':
            self.btn_lang.setText("🇬🇧 Lang: Eng [L]")
            self.btn_lang.setStyleSheet("background-color: #1E3A5F; border: 1px solid #60A5FA; color: #93C5FD; font-weight: bold; font-size: 12px;")
        else:
            self.btn_lang.setText("🇧🇩 Lang: বাংলা [L]")
            self.btn_lang.setStyleSheet("background-color: #132E20; border: 1px solid #34D399; color: #6EE7B7; font-weight: bold; font-size: 12px;")

    def update_voice_btn_ui(self):
        if self.voice_enabled:
            speed_mult = 1.0 + (self.voice_speed_percent / 100.0)
            self.btn_voice.setText(f"🔊 AI Voice: ON ({speed_mult:.2f}x) [V]")
            self.btn_voice.setStyleSheet("background-color: #059669; border: 1px solid #10B981; color: #FFFFFF; font-weight: bold; font-size: 12px;")
        else:
            self.btn_voice.setText("🔇 AI Voice: OFF [V]")
            self.btn_voice.setStyleSheet("background-color: #2D2D38; border: 1px solid #454556; color: #CBD5E1; font-weight: bold; font-size: 12px;")

    def increase_voice_speed(self):
        if self.voice_speed_percent < 80:
            self.voice_speed_percent += 15
            self.audio_cache.clear()
            self.update_voice_btn_ui()
            self.start_preload(self.current_idx)
            if self.voice_enabled:
                self.speak_text("Speed increased.", purge=True)

    def decrease_voice_speed(self):
        if self.voice_speed_percent > -30:
            self.voice_speed_percent -= 15
            self.audio_cache.clear()
            self.update_voice_btn_ui()
            self.start_preload(self.current_idx)
            if self.voice_enabled:
                self.speak_text("Speed decreased.", purge=True)

    def narrate_current_question(self):
        if not self.voice_enabled or self.current_idx >= len(self.questions):
            return
            
        cache_key = (self.current_idx, self.language_mode)
        
        # 1. Zero-Latency Instant Playback from Pre-cached Audio (< 1ms!)
        if cache_key in self.audio_cache and EDGE_TTS_AVAILABLE:
            self.stop_speech()
            try:
                if not pygame.mixer.get_init():
                    pygame.mixer.init()
                pygame.mixer.music.stop()
                sound_io = io.BytesIO(self.audio_cache[cache_key])
                pygame.mixer.music.load(sound_io)
                pygame.mixer.music.play()
                self.start_preload(self.current_idx + 1)
                return
            except Exception:
                pass

        # 2. If not yet cached, fetch and stream immediately
        q = self.questions[self.current_idx]
        speech, voice = build_question_speech_text(q, self.current_idx, self.language_mode)
        
        def save_to_cache(data):
            self.audio_cache[cache_key] = data

        self.speak_text(speech, voice=voice, purge=True, on_ready=save_to_cache)
        self.start_preload(self.current_idx + 1)

    def narrate_explanation(self):
        if not self.voice_enabled:
            self.voice_enabled = True
            self.update_voice_btn_ui()
        q = self.questions[self.current_idx]
        sol = q.get("solution", "").strip() or "No explanation provided for this question."
        s_eng, s_bn = split_english_bengali(sol)
        
        if self.language_mode == 'en':
            self.speak_text(f"Important Details. {s_eng or sol}", voice=self.voice_en, purge=True)
        elif self.language_mode == 'bn':
            self.speak_text(f"গুরুত্বপূর্ণ তথ্য ও সমাধান। {s_bn or sol}", voice=self.voice_bn, purge=True)
        else:
            self.speak_text(f"Important Details and Solution. {sol}", voice=self.voice_bn, purge=True)

    def is_speech_active(self):
        """Returns True if AI Voice is currently playing audio."""
        if self.voice_thread and self.voice_thread.isRunning():
            return True
        if EDGE_TTS_AVAILABLE and pygame and pygame.mixer.get_init():
            try:
                if pygame.mixer.get_busy():
                    return True
            except Exception:
                pass
        return False

    def toggle_mic_mode(self):
        if not SPEECH_RECOG_AVAILABLE and not VOSK_AVAILABLE:
            QMessageBox.warning(self, "Mic Speech Recognition", "Speech Recognition engine is not available.")
            return
        self.mic_enabled = not self.mic_enabled
        if self.mic_enabled:
            self.btn_mic.setText("🎙️ Mic: LISTENING... [S / Shift]")
            self.btn_mic.setStyleSheet("background-color: #B91C1C; border: 2px solid #EF4444; color: #FFFFFF; font-weight: bold; font-size: 12px;")
            self.lbl_mic_banner.setText("🎙️ ılı.lıllılı.ıllı Neural AI Listening: Speak Option A/B, 1, 2, 'Next' ...")
            self.lbl_mic_banner.setStyleSheet("""
                background-color: #1E1B2E;
                border: 1px solid #8B5CF6;
                border-radius: 8px;
                color: #EDE9FE;
                font-size: 13px;
                font-weight: 600;
                padding: 8px 14px;
            """)
            self.lbl_mic_banner.show()
            self.start_mic_listener()
        else:
            self.btn_mic.setText("🎙️ Mic: OFF [S / Shift]")
            self.btn_mic.setStyleSheet("background-color: #2D2D38; border: 1px solid #454556; color: #CBD5E1; font-weight: bold; font-size: 12px;")
            self.lbl_mic_banner.hide()
            self.stop_mic_listener()

    def get_current_question(self):
        """Returns the current question dictionary."""
        if 0 <= self.current_idx < len(self.questions):
            return self.questions[self.current_idx]
        return None

    def start_mic_listener(self):
        self.stop_mic_listener()
        play_system_sfx("mic_on")
        self.mic_thread = VoiceAnswerListener(
            lang_mode=self.language_mode,
            is_speaking_fn=self.is_speech_active,
            get_q_fn=self.get_current_question,
            parent=self
        )
        self.mic_thread.command_detected.connect(self.on_voice_command)
        self.mic_thread.heard_text.connect(self.on_voice_heard)
        self.mic_thread.partial_text.connect(self.on_voice_partial)
        self.mic_thread.status_changed.connect(self.on_mic_status)
        self.mic_thread.start()

    def stop_mic_listener(self):
        if self.mic_thread:
            play_system_sfx("mic_off")
            self.mic_thread.stop()
            self.mic_thread = None

    def on_voice_partial(self, partial_text):
        if not self.mic_enabled or not partial_text:
            return
        self.lbl_mic_banner.setText(f"🎙️ Live: \"{partial_text}...\"")
        self.lbl_mic_banner.setStyleSheet("""
            background-color: #1E293B;
            border: 1px solid #38BDF8;
            border-radius: 8px;
            color: #7DD3FC;
            font-size: 13px;
            font-weight: 600;
            padding: 8px 14px;
        """)

    def on_mic_status(self, status):
        if not self.mic_enabled:
            return
        if status == "listening":
            self.btn_mic.setText("🎙️ Mic: LISTENING... [S / Shift]")
            self.btn_mic.setStyleSheet("background-color: #B91C1C; border: 2px solid #EF4444; color: #FFFFFF; font-weight: bold; font-size: 12px;")
            self.lbl_mic_banner.setText("🎙️ ılı.lıllılı.ıllı AI Listening: Speak Answer (A/B/C/D or name), 'Next', 'Details'...")
            self.lbl_mic_banner.setStyleSheet("""
                background-color: #1E1B2E;
                border: 1px solid #8B5CF6;
                border-radius: 8px;
                color: #EDE9FE;
                font-size: 13px;
                font-weight: 600;
                padding: 8px 14px;
            """)
        elif status == "processing":
            self.btn_mic.setText("⏳ Mic: Processing... [S / Shift]")
            self.btn_mic.setStyleSheet("background-color: #D97706; border: 2px solid #F59E0B; color: #FFFFFF; font-weight: bold; font-size: 12px;")
        elif status == "error":
            self.btn_mic.setText("⚠️ Mic: Error [S / Shift]")
            self.btn_mic.setStyleSheet("background-color: #4B5563; border: 1px solid #6B7280; color: #D1D5DB; font-weight: bold; font-size: 12px;")

    def on_voice_heard(self, text):
        if not self.mic_enabled:
            return
        self.lbl_mic_banner.setText(f"👂 Heard: \"{text}\"")
        self.lbl_mic_banner.setStyleSheet("""
            background-color: #1E293B;
            border: 1px solid #38BDF8;
            border-radius: 8px;
            color: #7DD3FC;
            font-size: 13px;
            font-weight: 600;
            padding: 8px 14px;
        """)

    def on_voice_command(self, cmd_type, value, raw_text):
        if not self.mic_enabled:
            return
        letters = ["A", "B", "C", "D"]
        if cmd_type == "option":
            if self.current_idx not in self.user_answers:
                self.lbl_mic_banner.setText(f"🎯 Heard: \"{raw_text}\" ➔ Selecting Option {letters[value]} ✓")
                self.lbl_mic_banner.setStyleSheet("""
                    background-color: #064E3B;
                    border: 1px solid #10B981;
                    border-radius: 8px;
                    color: #A7F3D0;
                    font-size: 13px;
                    font-weight: bold;
                    padding: 8px 14px;
                """)
                self.check_answer(value)
        elif cmd_type == "next":
            self.lbl_mic_banner.setText(f"➡️ Heard: \"{raw_text}\" ➔ Next Question ✓")
            self.lbl_mic_banner.setStyleSheet("""
                background-color: #1E3A5F;
                border: 1px solid #60A5FA;
                border-radius: 8px;
                color: #BFDBFE;
                font-size: 13px;
                font-weight: bold;
                padding: 8px 14px;
            """)
            self.next_question()
        elif cmd_type == "prev":
            self.lbl_mic_banner.setText(f"⬅️ Heard: \"{raw_text}\" ➔ Previous Question ✓")
            self.lbl_mic_banner.setStyleSheet("""
                background-color: #1E3A5F;
                border: 1px solid #60A5FA;
                border-radius: 8px;
                color: #BFDBFE;
                font-size: 13px;
                font-weight: bold;
                padding: 8px 14px;
            """)
            self.prev_question()
        elif cmd_type == "mark":
            self.lbl_mic_banner.setText(f"⭐ Heard: \"{raw_text}\" ➔ Marked for Review ✓")
            self.lbl_mic_banner.setStyleSheet("""
                background-color: #581C87;
                border: 1px solid #C084FC;
                border-radius: 8px;
                color: #F3E8FF;
                font-size: 13px;
                font-weight: bold;
                padding: 8px 14px;
            """)
            self.toggle_mark_review()
        elif cmd_type == "clear":
            self.lbl_mic_banner.setText(f"🧹 Heard: \"{raw_text}\" ➔ Cleared Response ✓")
            self.lbl_mic_banner.setStyleSheet("""
                background-color: #374151;
                border: 1px solid #9CA3AF;
                border-radius: 8px;
                color: #F3F4F6;
                font-size: 13px;
                font-weight: bold;
                padding: 8px 14px;
            """)
            self.clear_response()
        elif cmd_type == "details":
            self.lbl_mic_banner.setText(f"💡 Heard: \"{raw_text}\" ➔ Reading Important Details ✓")
            self.lbl_mic_banner.setStyleSheet("""
                background-color: #4C1D95;
                border: 1px solid #A855F7;
                border-radius: 8px;
                color: #E9D5FF;
                font-size: 13px;
                font-weight: bold;
                padding: 8px 14px;
            """)
            self.narrate_explanation()
        elif cmd_type == "speed_up":
            self.lbl_mic_banner.setText(f"⏩ Heard: \"{raw_text}\" ➔ Speed Increased ✓")
            self.increase_voice_speed()
        elif cmd_type == "speed_down":
            self.lbl_mic_banner.setText(f"⏪ Heard: \"{raw_text}\" ➔ Speed Decreased ✓")
            self.decrease_voice_speed()
        elif cmd_type == "lang_bn":
            self.language_mode = 'bn'
            self.update_lang_btn_ui()
            self.lbl_mic_banner.setText(f"🇧🇩 Heard: \"{raw_text}\" ➔ Language: Bangla ✓")
            self.start_preload(self.current_idx)
            self.speak_text("ভাষা বাংলা সেট করা হয়েছে।", voice=self.voice_bn, purge=True)
        elif cmd_type == "lang_en":
            self.language_mode = 'en'
            self.update_lang_btn_ui()
            self.lbl_mic_banner.setText(f"🇬🇧 Heard: \"{raw_text}\" ➔ Language: English ✓")
            self.start_preload(self.current_idx)
            self.speak_text("Language set to English.", voice=self.voice_en, purge=True)
        elif cmd_type == "lang_both":
            self.language_mode = 'both'
            self.update_lang_btn_ui()
            self.lbl_mic_banner.setText(f"🌐 Heard: \"{raw_text}\" ➔ Language: Both ✓")
            self.start_preload(self.current_idx)
            self.speak_text("Language set to English and Bangla.", voice=self.voice_bn, purge=True)

    def closeEvent(self, event):
        self.stop_speech()
        self.stop_mic_listener()
        if self.preloader_thread and self.preloader_thread.isRunning():
            self.preloader_thread.is_cancelled = True
        self.audio_cache.clear()
        super().closeEvent(event)

    def keyPressEvent(self, event):
        key = event.key()
        text = event.text().upper()

        # Keyboard Option Shortcuts: 1, 2, 3, 4 or A, B, C, D
        if key in (Qt.Key.Key_1, Qt.Key.Key_A):
            self.check_answer(0)
        elif key in (Qt.Key.Key_2, Qt.Key.Key_B):
            self.check_answer(1)
        elif key in (Qt.Key.Key_3, Qt.Key.Key_C):
            self.check_answer(2)
        elif key in (Qt.Key.Key_4, Qt.Key.Key_D):
            self.check_answer(3)
        elif key in (Qt.Key.Key_Control, Qt.Key.Key_AltGr) or (event.modifiers() & Qt.KeyboardModifier.ControlModifier and key == Qt.Key.Key_Control):
            self.narrate_explanation()
        elif text == 'S' or key == Qt.Key.Key_Shift or (event.modifiers() & Qt.KeyboardModifier.ShiftModifier and key == Qt.Key.Key_Shift):
            self.toggle_mic_mode()
        elif text == 'V':
            self.toggle_voice_mode()
        elif text == 'L':
            self.cycle_language_mode()
        elif key in (Qt.Key.Key_Plus, Qt.Key.Key_Equal):
            self.increase_voice_speed()
        elif key in (Qt.Key.Key_Minus, Qt.Key.Key_Underscore):
            self.decrease_voice_speed()
        elif key in (Qt.Key.Key_Right, Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Space):
            self.next_question()
        elif key == Qt.Key.Key_Left:
            self.prev_question()
        elif text == 'M':
            self.toggle_mark_review()
        elif text == 'C' or key in (Qt.Key.Key_Backspace, Qt.Key.Key_Delete):
            self.clear_response()
        else:
            super().keyPressEvent(event)

    def toggle_mark_review(self):
        if self.current_idx in self.marked_for_review:
            self.marked_for_review.remove(self.current_idx)
            self.btn_mark.setText("⭐ Mark for Review")
            self.btn_mark.setStyleSheet("background-color: #6A1B9A; border: 1px solid #8E24AA; color: #FFFFFF;")
        else:
            self.marked_for_review.add(self.current_idx)
            self.btn_mark.setText("⭐ Marked ✓")
            self.btn_mark.setStyleSheet("background-color: #9C27B0; border: 2px solid #E1BEE7; color: #FFFFFF; font-weight: bold;")
        self.update_palette()

    def clear_response(self):
        if self.current_idx in self.user_answers:
            del self.user_answers[self.current_idx]
            # Recalculate score
            self.recalc_score()
            self.load_question()

    def recalc_score(self):
        letters = ["A", "B", "C", "D"]
        self.score = 0
        for idx, sel_idx in self.user_answers.items():
            q = self.questions[idx]
            ans_raw = q.get("answer", "").strip().upper()
            corr = 0
            for k, l in enumerate(letters):
                if l in ans_raw or ans_raw == l: corr = k; break
            if sel_idx == corr:
                self.score += 1

    def update_palette(self):
        letters = ["A", "B", "C", "D"]
        ans_count = 0
        wrong_count = 0
        marked_count = len(self.marked_for_review)
        
        for idx, btn in enumerate(self.palette_btns):
            is_current = (idx == self.current_idx)
            is_marked = idx in self.marked_for_review
            ans_selected = self.user_answers.get(idx, None)
            
            if ans_selected is not None:
                q = self.questions[idx]
                ans_raw = q.get("answer", "").strip().upper()
                correct_idx = -1
                for k, letter in enumerate(letters):
                    if letter in ans_raw or ans_raw == letter:
                        correct_idx = k
                        break
                if correct_idx == -1: correct_idx = 0
                
                is_correct = (ans_selected == correct_idx)
                if is_correct:
                    ans_count += 1
                    bg = "#1B5E20" # Green
                    border = "#4CAF50" if not is_current else "#00BCD4"
                else:
                    wrong_count += 1
                    bg = "#B71C1C" # Red
                    border = "#F44336" if not is_current else "#00BCD4"
                    
                if is_marked:
                    border = "#BA68C8"
                    border_w = "2px"
                else:
                    border_w = "3px" if is_current else "1px"
                btn.setStyleSheet(f"background-color: {bg}; border: {border_w} solid {border}; border-radius: 6px; color: #FFFFFF; font-weight: bold;")
            elif is_marked:
                bg = "#6A1B9A" # Purple
                border = "#CE93D8" if not is_current else "#00BCD4"
                border_w = "3px" if is_current else "1px"
                btn.setStyleSheet(f"background-color: {bg}; border: {border_w} solid {border}; border-radius: 6px; color: #FFFFFF; font-weight: bold;")
            else:
                if is_current:
                    btn.setStyleSheet("background-color: #262630; border: 3px solid #00BCD4; border-radius: 6px; color: #00BCD4; font-weight: bold;")
                else:
                    btn.setStyleSheet("background-color: #1E1E24; border: 1px solid #333340; border-radius: 6px; color: #94A3B8; font-weight: normal;")
                    
        # Update legend counters
        self.lbl_legend_ans.setText(f"🟢 Answered: {ans_count}")
        self.lbl_legend_wrong.setText(f"🔴 Wrong: {wrong_count}")
        self.lbl_legend_marked.setText(f"🟣 Marked: {marked_count}")
        not_vis = max(0, len(self.questions) - len(self.user_answers) - (marked_count - len(set(self.user_answers.keys()).intersection(self.marked_for_review))))
        self.lbl_legend_unvis.setText(f"⚪ Unattempted: {not_vis}")

        # Ensure active button is visible in scroll area
        if 0 <= self.current_idx < len(self.palette_btns):
            self.grid_scroll.ensureWidgetVisible(self.palette_btns[self.current_idx])

    def export_html(self):
        d = QuizOverviewDialog(self.questions, self.user_answers, self.score, self.seconds_elapsed, self)
        d.export_html()

    def open_overview(self):
        self.stop_speech()
        d = QuizOverviewDialog(self.questions, self.user_answers, self.score, self.seconds_elapsed, self)
        if d.exec() == QDialog.DialogCode.Accepted:
            if d.selected_jump_idx is not None:
                self.jump_to_question(d.selected_jump_idx)

    def jump_to_question(self, idx):
        if 0 <= idx < len(self.questions):
            self.stop_speech()
            self.current_idx = idx
            self.load_question()

    def prev_question(self):
        if self.current_idx > 0:
            self.stop_speech()
            self.current_idx -= 1
            self.load_question()

    def next_question(self):
        self.stop_speech()
        if self.current_idx < len(self.questions) - 1:
            self.current_idx += 1
            self.load_question()
        else:
            self.confirm_submit()

    def confirm_submit(self):
        self.stop_speech()
        unanswered = len(self.questions) - len(self.user_answers)
        msg = f"Are you sure you want to submit the test?\n\n• Total Questions: {len(self.questions)}\n• Attempted: {len(self.user_answers)}\n• Unattempted: {unanswered}"
        res = QMessageBox.question(self, "Submit Test", msg, QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if res == QMessageBox.StandardButton.Yes:
            self.show_results()

    def update_timer(self):
        self.seconds_elapsed += 1
        m = self.seconds_elapsed // 60
        s = self.seconds_elapsed % 60
        self.lbl_timer.setText(f"⏱️ Time: {m:02d}:{s:02d}")

    def load_question(self):
        if self.current_idx >= len(self.questions):
            self.show_results()
            return

        self.update_palette()
        self.btn_prev.setEnabled(self.current_idx > 0)
        q = self.questions[self.current_idx]
        top_name = get_question_topic(q)
        self.lbl_q_title.setText(f"Question {self.current_idx + 1} of {len(self.questions)}   •   {top_name}")

        # Update Mark for Review button state
        if self.current_idx in self.marked_for_review:
            self.btn_mark.setText("⭐ Marked ✓")
            self.btn_mark.setStyleSheet("background-color: #9C27B0; border: 2px solid #E1BEE7; color: #FFFFFF; font-weight: bold;")
        else:
            self.btn_mark.setText("⭐ Mark for Review")
            self.btn_mark.setStyleSheet("background-color: #6A1B9A; border: 1px solid #8E24AA; color: #FFFFFF;")

        self.lbl_question.setText(q.get("question", ""))

        # Options text with [1, 2, 3, 4] keyboard cues
        keys = ["1", "2", "3", "4"]
        letters = ["A", "B", "C", "D"]
        for i in range(4):
            opt_key = f"option{letters[i]}"
            opt_text = q.get(opt_key, "")
            self.opt_buttons[i].setText(f" [ {keys[i]} ]   {letters[i]}.  {opt_text}")

        # Image render
        img_base64 = q.get("questionImage", "")
        if img_base64:
            try:
                img_data = base64.b64decode(img_base64)
                pixmap = QPixmap()
                pixmap.loadFromData(img_data)
                scaled = pixmap.scaled(700, 250, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
                self.lbl_image.setPixmap(scaled)
                self.lbl_image.show()
            except Exception as e:
                self.lbl_image.hide()
        else:
            self.lbl_image.hide()

        # Check if already answered
        ans_raw = q.get("answer", "").strip().upper()
        correct_idx = -1
        for idx, letter in enumerate(letters):
            if letter in ans_raw or ans_raw == letter:
                correct_idx = idx
                break
        if correct_idx == -1: correct_idx = 0

        if self.current_idx in self.user_answers:
            selected_idx = self.user_answers[self.current_idx]
            for btn in self.opt_buttons:
                btn.setEnabled(False)
                btn.setStyleSheet("")
            
            # Correct button
            self.opt_buttons[correct_idx].setStyleSheet("""
                QPushButton#opt_btn {
                    background-color: #1B5E20;
                    border: 2px solid #4CAF50;
                    color: #FFFFFF;
                    font-weight: bold;
                }
            """)
            if selected_idx != correct_idx and selected_idx != -1:
                self.opt_buttons[selected_idx].setStyleSheet("""
                    QPushButton#opt_btn {
                        background-color: #B71C1C;
                        border: 2px solid #F44336;
                        color: #FFFFFF;
                        font-weight: bold;
                    }
                """)
            
            sol_text = q.get("solution", "").strip() or "No explanation provided for this question."
            self.lbl_solution.setText(sol_text)
            self.sol_card.show()
            self.btn_next.setEnabled(True)
        else:
            # Not answered yet
            for btn in self.opt_buttons:
                btn.setEnabled(True)
                btn.setStyleSheet("")
            self.sol_card.hide()
            self.btn_next.setEnabled(True)
            if self.voice_enabled:
                self.narrate_current_question()

        # Update Mic Banner state
        if self.mic_enabled:
            self.lbl_mic_banner.setText("🎙️ ılı.lıllılı.ıllı AI Listening: Speak Answer (A/B/C/D or name), 'Next', 'Details'...")
            self.lbl_mic_banner.setStyleSheet("""
                background-color: #1E1B2E;
                border: 1px solid #8B5CF6;
                border-radius: 8px;
                color: #EDE9FE;
                font-size: 13px;
                font-weight: 600;
                padding: 8px 14px;
            """)
            self.lbl_mic_banner.show()
        else:
            self.lbl_mic_banner.hide()

    def check_answer(self, selected_idx):
        if self.current_idx in self.user_answers:
            return
            
        self.user_answers[self.current_idx] = selected_idx
        q = self.questions[self.current_idx]
        ans_raw = q.get("answer", "").strip().upper()
        
        # Determine correct index
        letters = ["A", "B", "C", "D"]
        correct_idx = -1
        for idx, letter in enumerate(letters):
            if letter in ans_raw or ans_raw == letter:
                correct_idx = idx
                break
        if correct_idx == -1:
            correct_idx = 0 # default fallback

        # Style selected and correct buttons
        for btn in self.opt_buttons:
            btn.setEnabled(False)

        # Style correct answer
        self.opt_buttons[correct_idx].setStyleSheet("""
            QPushButton#opt_btn {
                background-color: #1B5E20;
                border: 2px solid #4CAF50;
                color: #FFFFFF;
                font-weight: bold;
            }
        """)

        is_correct = (selected_idx == correct_idx)
        if is_correct:
            self.score += 1
            play_system_sfx("correct")
        else:
            play_system_sfx("wrong")
            # Style incorrect selected answer
            self.opt_buttons[selected_idx].setStyleSheet("""
                QPushButton#opt_btn {
                    background-color: #B71C1C;
                    border: 2px solid #F44336;
                    color: #FFFFFF;
                    font-weight: bold;
                }
            """)

        # Show Explanation
        sol_text = q.get("solution", "").strip()
        if not sol_text:
            sol_text = "No explanation provided for this question."
        self.lbl_solution.setText(sol_text)
        self.sol_card.show()
        
        self.update_palette()
        self.btn_next.setEnabled(True)

        # Voice feedback with Important Details & Solution narration
        if self.voice_enabled:
            corr_letter = letters[correct_idx]
            sel_letter = letters[selected_idx]
            sol_text = q.get("solution", "").strip()
            s_eng, s_bn = split_english_bengali(sol_text)
            
            if self.language_mode == 'en':
                prefix = f"Correct! Option {sel_letter} is right." if is_correct else f"Incorrect. Correct answer is Option {corr_letter}."
                detail = f" Important Details: {s_eng or sol_text}" if (s_eng or sol_text) else ""
                self.speak_text(prefix + detail, voice=self.voice_en, purge=True)
            elif self.language_mode == 'bn':
                prefix = f"সঠিক উত্তর! অপশন {sel_letter} সঠিক।" if is_correct else f"ভুল উত্তর। সঠিক উত্তর হলো অপশন {corr_letter}।"
                detail = f" গুরুত্বপূর্ণ তথ্য: {s_bn or sol_text}" if (s_bn or sol_text) else ""
                self.speak_text(prefix + detail, voice=self.voice_bn, purge=True)
            else: # both
                prefix = f"Correct! Option {sel_letter} is right." if is_correct else f"Incorrect. Correct answer is Option {corr_letter}."
                detail = f" Important Details: {sol_text}" if sol_text else ""
                self.speak_text(prefix + detail, voice=self.voice_bn, purge=True)

    def show_results(self):
        self.timer.stop()
        self.stop_speech()
        self.stop_mic_listener()
        if self.preloader_thread and self.preloader_thread.isRunning():
            self.preloader_thread.is_cancelled = True
        self.audio_cache.clear()
        
        # Clear dialog layouts and replace with score screen
        for i in reversed(range(self.main_layout.count())):
            item = self.main_layout.takeAt(i)
            if item.widget():
                item.widget().deleteLater()

        total = len(self.questions)
        attempted = len(self.user_answers)
        correct = self.score
        wrong = attempted - correct
        unattempted = max(0, total - attempted)
        pct = (correct / total) * 100 if total else 0
        m = self.seconds_elapsed // 60
        s = self.seconds_elapsed % 60
        avg_time = (self.seconds_elapsed / attempted) if attempted else 0
        net_marks = max(0.0, (correct * 1.0) - (wrong * 0.25))

        # Performance Tier
        if pct >= 85:
            perf_text = "🌟 Outstanding Performance!"
            perf_color = "#4ADE80"
def get_question_topic(q, default="General Awareness"):
    for k in ["topic", "subject", "category", "section", "source", "chapter"]:
        if q.get(k) and str(q.get(k)).strip():
            return str(q.get(k)).strip()
    
    text = (q.get("question", "") + " " + q.get("solution", "")).lower()
    if any(w in text for w in ["wetland", "ramsar", "forest", "national park", "wildlife", "climate", "tiger", "environment", "ecology"]):
        return "🌿 Environment & Ecology"
    elif any(w in text for w in ["prime minister", "minister", "president", "governor", "election", "bjp", "congress", "parliament", "chief minister", "pm", "constitution", "article", "lok sabha", "rajya sabha"]):
        return "🏛️ Polity & Governance"
    elif any(w in text for w in ["award", "prize", "padma", "bharat ratna", "oscar", "grammy", "nobel", "academy award"]):
        return "🏆 Awards & Honors"
    elif any(w in text for w in ["sports", "cricket", "olympic", "football", "world cup", "badminton", "chess", "tennis", "hockey", "medal"]):
        return "⚽ Sports & Games"
    elif any(w in text for w in ["satellite", "isro", "nasa", "drdo", "missile", "ai", "physics", "chemistry", "biology", "space", "technology", "quantum"]):
        return "🔬 Science & Tech"
    elif any(w in text for w in ["gdp", "rbi", "budget", "inflation", "bank", "economic", "trade", "gst", "rupee", "repo rate"]):
        return "📈 Economy & Banking"
    elif any(w in text for w in ["summit", "g20", "un ", "united nations", "treaty", "ambassador", "bilateral", "global", "international"]):
        return "🌐 Global Affairs"
    elif any(w in text for w in ["history", "dynasty", "war", "battle", "freedom", "mughal", "british", "revolt", "ancient", "medieval"]):
        return "📜 History & Heritage"
    elif any(w in text for w in ["river", "mountain", "ocean", "capital", "strait", "lake", "border", "island", "plateau"]):
        return "🗺️ Geography & Places"
    
    return default

    def show_results(self):
        self.timer.stop()
        
        # Clear dialog layouts and replace with score screen
        for i in reversed(range(self.main_layout.count())):
            item = self.main_layout.takeAt(i)
            if item.widget():
                item.widget().deleteLater()

        total = len(self.questions)
        attempted = len(self.user_answers)
        correct = self.score
        wrong = attempted - correct
        unattempted = max(0, total - attempted)
        pct = (correct / total) * 100 if total else 0
        m = self.seconds_elapsed // 60
        s = self.seconds_elapsed % 60
        avg_time = (self.seconds_elapsed / attempted) if attempted else 0
        net_marks = max(0.0, (correct * 1.0) - (wrong * 0.25))

        # Topic Breakdown Calculation
        letters = ["A", "B", "C", "D"]
        topic_stats = {}
        for idx, q in enumerate(self.questions):
            top_name = get_question_topic(q)
            if top_name not in topic_stats:
                topic_stats[top_name] = {"total": 0, "correct": 0, "attempted": 0}
            topic_stats[top_name]["total"] += 1
            
            sel_idx = self.user_answers.get(idx, -1)
            ans_raw = q.get("answer", "").strip().upper()
            corr_idx = 0
            for k, l in enumerate(letters):
                if l in ans_raw or ans_raw == l: corr_idx = k; break
            
            if sel_idx != -1:
                topic_stats[top_name]["attempted"] += 1
                if sel_idx == corr_idx:
                    topic_stats[top_name]["correct"] += 1

        # Performance Tier
        if pct >= 85:
            perf_text = "🌟 Outstanding Performance!"
            perf_color = "#4ADE80"
        elif pct >= 65:
            perf_text = "🎯 Very Good Score!"
            perf_color = "#38BDF8"
        elif pct >= 40:
            perf_text = "👍 Good Effort, Keep Practicing!"
            perf_color = "#FBBF24"
        else:
            perf_text = "📚 Needs Revision & Practice"
            perf_color = "#F87171"

        card = QFrame()
        card.setObjectName("card")
        card.setStyleSheet("""
            QFrame#card {
                background-color: #17171C;
                border: 1px solid #2C2C36;
                border-radius: 16px;
            }
        """)
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(28, 20, 28, 20)
        card_layout.setSpacing(14)

        # Header Title
        lbl_congrats = QLabel("🎉 Test Submitted Successfully!")
        lbl_congrats.setStyleSheet("font-size: 22px; font-weight: bold; color: #00BCD4;")
        lbl_congrats.setAlignment(Qt.AlignmentFlag.AlignCenter)
        
        lbl_sub = QLabel("Laptop CBT Exam Series • Performance Analysis & Scorecard")
        lbl_sub.setStyleSheet("font-size: 13px; color: #94A3B8;")
        lbl_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)

        card_layout.addWidget(lbl_congrats)
        card_layout.addWidget(lbl_sub)

        # Hero Scorecard Frame
        hero_frame = QFrame()
        hero_frame.setStyleSheet("background-color: #1F1F27; border: 1px solid #333340; border-radius: 12px; padding: 12px;")
        hero_layout = QHBoxLayout(hero_frame)
        hero_layout.setContentsMargins(18, 12, 18, 12)
        
        # Left side: Marks & Accuracy
        v_score = QVBoxLayout()
        lbl_score_num = QLabel(f"<span style='font-size: 32px; font-weight: bold; color: #4ADE80;'>{correct}</span> <span style='font-size: 16px; color: #94A3B8;'>/ {total} Correct</span>")
        lbl_acc = QLabel(f"🎯 Accuracy: <b>{pct:.1f}%</b>  |  Net Score: <b>{net_marks:.2f} / {total}</b>")
        lbl_acc.setStyleSheet("font-size: 13.5px; color: #CBD5E1;")
        v_score.addWidget(lbl_score_num)
        v_score.addWidget(lbl_acc)
        hero_layout.addLayout(v_score)
        hero_layout.addStretch()

        # Right side: Tier & Time
        v_tier = QVBoxLayout()
        lbl_perf = QLabel(perf_text)
        lbl_perf.setStyleSheet(f"font-size: 15px; font-weight: bold; color: {perf_color};")
        lbl_time_stat = QLabel(f"⏱️ Total Time: <b>{m:02d}:{s:02d}</b> (Avg {avg_time:.1f}s/Q)")
        lbl_time_stat.setStyleSheet("font-size: 13px; color: #FF9800;")
        v_tier.addWidget(lbl_perf)
        v_tier.addWidget(lbl_time_stat)
        hero_layout.addLayout(v_tier)

        card_layout.addWidget(hero_frame)

        # 6 Key Metrics Grid
        grid_metrics = QGridLayout()
        grid_metrics.setSpacing(8)

        stats_data = [
            ("📋 Total Questions", f"{total}", "#00BCD4"),
            ("✍️ Attempted", f"{attempted} / {total}", "#FFC107"),
            ("🟢 Correct Answers", f"{correct} (+{correct*1.0:.2f})", "#4CAF50"),
            ("🔴 Incorrect Answers", f"{wrong} (-{wrong*0.25:.2f})", "#F44336"),
            ("⚪ Unattempted", f"{unattempted}", "#94A3B8"),
            ("⏱️ Time Elapsed", f"{m:02d}:{s:02d}", "#FF9800"),
        ]

        for i, (label, val, col) in enumerate(stats_data):
            r, c = divmod(i, 3)
            box = QFrame()
            box.setStyleSheet("background-color: #202028; border: 1px solid #2E2E3C; border-radius: 8px;")
            b_layout = QVBoxLayout(box)
            b_layout.setContentsMargins(10, 8, 10, 8)
            lbl_title = QLabel(label)
            lbl_title.setStyleSheet("font-size: 11.5px; color: #94A3B8; font-weight: 500;")
            lbl_val = QLabel(val)
            lbl_val.setStyleSheet(f"font-size: 14px; font-weight: bold; color: {col};")
            b_layout.addWidget(lbl_title)
            b_layout.addWidget(lbl_val)
            grid_metrics.addWidget(box, r, c)

        card_layout.addLayout(grid_metrics)

        # 📚 Topic & Subject Performance Breakdown Section
        lbl_topics_hdr = QLabel("📚 Question Topics & Domain Analysis (Source Breakdown):")
        lbl_topics_hdr.setStyleSheet("font-size: 13px; font-weight: bold; color: #00BCD4; margin-top: 4px;")
        card_layout.addWidget(lbl_topics_hdr)

        topic_scroll = QScrollArea()
        topic_scroll.setFixedHeight(95)
        topic_scroll.setWidgetResizable(True)
        topic_scroll.setStyleSheet("background-color: #1A1A22; border: 1px solid #2C2C38; border-radius: 8px;")
        
        topic_widget = QWidget()
        topic_layout = QHBoxLayout(topic_widget)
        topic_layout.setContentsMargins(8, 6, 8, 6)
        topic_layout.setSpacing(8)

        for top_name, t_data in topic_stats.items():
            t_tot = t_data["total"]
            t_cor = t_data["correct"]
            t_pct = (t_cor / t_tot * 100) if t_tot else 0
            
            t_box = QFrame()
            t_box.setStyleSheet("background-color: #22222D; border: 1px solid #363646; border-radius: 6px; padding: 6px 10px;")
            tb_layout = QVBoxLayout(t_box)
            tb_layout.setContentsMargins(6, 4, 6, 4)
            tb_layout.setSpacing(2)
            
            lbl_t_title = QLabel(top_name)
            lbl_t_title.setStyleSheet("color: #FFFFFF; font-weight: bold; font-size: 12px;")
            lbl_t_score = QLabel(f"Score: <b>{t_cor}/{t_tot}</b> ({t_pct:.0f}%)")
            lbl_t_score.setStyleSheet(f"color: {'#4CAF50' if t_pct>=70 else '#FFC107' if t_pct>=40 else '#F44336'}; font-size: 11px;")
            
            tb_layout.addWidget(lbl_t_title)
            tb_layout.addWidget(lbl_t_score)
            topic_layout.addWidget(t_box)

        topic_layout.addStretch()
        topic_scroll.setWidget(topic_widget)
        card_layout.addWidget(topic_scroll)

        # Action Buttons Row
        btn_row = QHBoxLayout()
        btn_row.setSpacing(10)

        btn_overview_sheet = QPushButton("📊 View Detailed A4 Sheet & Solutions")
        btn_overview_sheet.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_overview_sheet.setStyleSheet("background-color: #0078D4; border: 1px solid #005A9E; border-radius: 8px; padding: 10px 16px; color: white; font-weight: bold; font-size: 13px;")
        btn_overview_sheet.clicked.connect(self.open_overview)

        btn_download_html_res = QPushButton("📥 Download HTML (Mobile / Web)")
        btn_download_html_res.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_download_html_res.setStyleSheet("background-color: #10B981; border: 1px solid #059669; border-radius: 8px; padding: 10px 16px; color: white; font-weight: bold; font-size: 13px;")
        btn_download_html_res.clicked.connect(self.export_html)

        self.btn_save_report = QPushButton("💾 Save .txt Report")
        self.btn_save_report.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_save_report.setStyleSheet("background-color: #D97706; border: 1px solid #B45309; border-radius: 8px; padding: 10px 14px; color: white; font-weight: bold; font-size: 13px;")
        self.btn_save_report.clicked.connect(self.save_test_report)

        btn_close = QPushButton("Close")
        btn_close.setCursor(Qt.CursorShape.PointingHandCursor)
        btn_close.setStyleSheet("background-color: #374151; border: 1px solid #4B5563; border-radius: 8px; padding: 10px 16px; color: white; font-weight: bold; font-size: 13px;")
        btn_close.clicked.connect(self.accept)

        btn_row.addWidget(btn_overview_sheet)
        btn_row.addWidget(btn_download_html_res)
        btn_row.addWidget(self.btn_save_report)
        btn_row.addWidget(btn_close)
        card_layout.addLayout(btn_row)

        self.main_layout.addWidget(card)

    def save_test_report(self):
        try:
            from datetime import datetime
            import os
            
            target_dir = os.path.join(os.path.expanduser("~"), "Downloads", "Daily quiz")
            if not os.path.exists(target_dir):
                os.makedirs(target_dir, exist_ok=True)
                
            now = datetime.now()
            timestamp = now.strftime("%Y%m%d_%H%M%S")
            filename = f"quiz_result_{timestamp}.txt"
            full_path = os.path.join(target_dir, filename)
            
            pct = (self.score / len(self.questions)) * 100 if self.questions else 0
            m = self.seconds_elapsed // 60
            s = self.seconds_elapsed % 60
            
            report_lines = [
                "==================================================",
                "🏆 LAPTOP QUIZ PRACTICE REPORT",
                "==================================================",
                f"Date & Time : {now.strftime('%Y-%m-%d %H:%M:%S')}",
                f"Total Questions : {len(self.questions)}",
                f"Correct Answers : {self.score}",
                f"Accuracy : {pct:.1f}%",
                f"Time Taken : {m:02d}:{s:02d}",
                "==================================================",
                "\nDetailed Question Log:\n"
            ]
            
            letters = ["A", "B", "C", "D"]
            for idx, q in enumerate(self.questions):
                selected_idx = self.user_answers.get(idx, -1)
                
                # Determine correct index
                ans_raw = q.get("answer", "").strip().upper()
                correct_idx = -1
                for k, letter in enumerate(letters):
                    if letter in ans_raw or ans_raw == letter:
                        correct_idx = k
                        break
                if correct_idx == -1: correct_idx = 0
                
                status = "CORRECT" if selected_idx == correct_idx else "INCORRECT"
                sel_text = letters[selected_idx] if selected_idx != -1 else "No Answer"
                corr_text = letters[correct_idx]
                
                report_lines.append(f"Question {idx+1}: {q.get('question', '')}")
                report_lines.append(f"  Options:")
                report_lines.append(f"    A. {q.get('optionA', '')}")
                report_lines.append(f"    B. {q.get('optionB', '')}")
                report_lines.append(f"    C. {q.get('optionC', '')}")
                report_lines.append(f"    D. {q.get('optionD', '')}")
                report_lines.append(f"  Your Answer    : {sel_text}")
                report_lines.append(f"  Correct Answer : {corr_text}")
                report_lines.append(f"  Result         : {status}")
                report_lines.append(f"  Explanation    : {q.get('solution', '')}")
                report_lines.append("-" * 50)
                
            with open(full_path, 'w', encoding='utf-8') as f:
                f.write("\n".join(report_lines))
                
            # Also save the actual Quiz questions as a JSON file!
            json_filename = f"quiz_data_{timestamp}.json"
            json_full_path = os.path.join(target_dir, json_filename)
            with open(json_full_path, 'w', encoding='utf-8') as f:
                json.dump(self.questions, f, indent=4, ensure_ascii=False)
                
            QMessageBox.information(self, "Quiz & Report Saved", f"Saved successfully in folder:\n{target_dir}\n\n1. Report: {filename}\n2. Quiz Data: {json_filename}")
            self.btn_save_report.setText("Saved! ✓")
            self.btn_save_report.setEnabled(False)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save quiz & report: {e}")

global_floating_widgets = []

class RoutineWidget(QWidget):
    widget_data_received = pyqtSignal(dict)
    
    def __init__(self, db, userId, parent=None):
        super().__init__(parent, Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.Tool)
        self.db = db
        self.userId = userId
        self.drag_pos = QPoint()
        self.current_days = []
        self.selected_day_offset = 0 # 0 = Today, 1 = Tomorrow, etc.
        
        # Mouse Resizing & Dragging States
        self.setMouseTracking(True)
        self.resize_margin = 12
        self.resize_dir = None
        
        # Timer States
        self.timer_seconds_left = 0
        self.timer_total_seconds = 0
        self.timer_running = False
        self.timer_subject_name = ""
        self.timer_slot_type = ""
        self.timer_index = 0
        self.timer_day_name = ""
        self.timer_active_sound = False
        
        self.setWindowTitle("Study Routine Widget")
        self.setMinimumSize(250, 400)
        self.resize(270, 480)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        self.setStyleSheet("""
            QFrame#widgetFrame {
                background-color: rgba(20, 20, 20, 120);
                border: 2px solid rgba(0, 188, 212, 180);
                border-radius: 16px;
            }
            QLabel {
                color: #E2E2E2;
                font-family: 'Segoe UI', -apple-system, sans-serif;
            }
            QLabel#headerTitle {
                font-size: 14px;
                font-weight: bold;
                color: #00E5FF;
            }
            QLabel#slotHeader {
                font-size: 10px;
                font-weight: bold;
                color: #FF9800;
                margin-top: 5px;
                letter-spacing: 0.5px;
            }
            QPushButton#navBtn {
                background-color: rgba(255, 255, 255, 15);
                border: 1px solid rgba(255, 255, 255, 10);
                border-radius: 6px;
                color: #E2E2E2;
                font-weight: bold;
                font-size: 11px;
                padding: 3px 8px;
            }
            QPushButton#navBtn:hover {
                background-color: rgba(0, 188, 212, 80);
                border: 1px solid #00BCD4;
                color: #FFFFFF;
            }
            QPushButton#closeBtn {
                background: transparent;
                border: none;
                color: #94A3B8;
                font-size: 16px;
                font-weight: bold;
            }
            QPushButton#closeBtn:hover {
                color: #FF5722;
            }
            
            /* Task Card Styling (Solid dark block blocks background window text bleed-through) */
            QFrame#taskCard {
                background-color: #202020;
                border: 1px solid rgba(255, 255, 255, 12);
                border-radius: 12px;
            }
            QFrame#taskCard:hover {
                background-color: #2D2D30;
                border: 1px solid rgba(0, 188, 212, 140);
            }
            QLabel#taskLabel {
                font-size: 12px;
                font-weight: bold;
                color: #FFFFFF;
                line-height: 1.3;
            }
            QPushButton#checkCircle {
                background-color: rgba(255, 255, 255, 8);
                border: 2px solid rgba(255, 255, 255, 25);
                border-radius: 10px;
                min-width: 20px;
                max-width: 20px;
                min-height: 20px;
                max-height: 20px;
            }
            QPushButton#checkCircle:hover {
                border: 2px solid #10B981;
                background-color: rgba(16, 185, 129, 30);
            }
            QPushButton#checkCircleChecked {
                background-color: #4CAF50;
                border: 2px solid #4CAF50;
                border-radius: 10px;
                min-width: 20px;
                max-width: 20px;
                min-height: 20px;
                max-height: 20px;
                color: white;
                font-weight: bold;
                font-size: 10px;
            }
            QPushButton#timerIconBtn {
                background-color: rgba(255, 255, 255, 12);
                border: 1px solid rgba(255, 255, 255, 10);
                border-radius: 8px;
                font-size: 13px;
                min-width: 28px;
                max-width: 28px;
                min-height: 28px;
                max-height: 28px;
                color: #E2E2E2;
            }
            QPushButton#timerIconBtn:hover {
                background-color: rgba(0, 188, 212, 60);
                border: 1px solid rgba(0, 188, 212, 120);
            }
            
            /* Timer View Styles */
            QLabel#timerLabel {
                font-size: 34px;
                font-weight: bold;
                color: #00E5FF;
                letter-spacing: 1px;
            }
            QProgressBar {
                background-color: rgba(255, 255, 255, 10);
                border: 1px solid rgba(255, 255, 255, 12);
                border-radius: 6px;
                text-align: center;
                color: transparent;
                height: 12px;
            }
            QProgressBar::chunk {
                background-color: #00BCD4;
                border-radius: 5px;
            }
            QPushButton#controlBtn {
                background-color: #00BCD4;
                border: none;
                border-radius: 8px;
                padding: 8px 16px;
                font-weight: bold;
                color: white;
                font-size: 12px;
            }
            QPushButton#controlBtn:hover {
                background-color: #00ACC1;
            }
        """)

        # Main Layout inside a rounded frame
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6) # 6px transparent resize margin
        self.setMouseTracking(True)
        self.frame = QFrame()
        self.frame.setMouseTracking(True)
        self.frame.setObjectName("widgetFrame")
        self.frame_layout = QVBoxLayout(self.frame)
        self.frame_layout.setContentsMargins(12, 12, 12, 12)
        self.frame_layout.setSpacing(6)
        
        # Title Bar (Always Visible)
        title_bar = QHBoxLayout()
        self.lbl_day = QLabel("📅 TODAY")
        self.lbl_day.setObjectName("headerTitle")
        title_bar.addWidget(self.lbl_day)
        title_bar.addStretch()
        
        self.btn_prev = QPushButton("◀")
        self.btn_prev.setObjectName("navBtn")
        self.btn_prev.clicked.connect(self.prev_day)
        
        self.btn_next_day = QPushButton("▶")
        self.btn_next_day.setObjectName("navBtn")
        self.btn_next_day.clicked.connect(self.next_day)
        
        self.btn_close = QPushButton("✕")
        self.btn_close.setObjectName("closeBtn")
        self.btn_close.clicked.connect(self.close)
        
        title_bar.addWidget(self.btn_prev)
        title_bar.addWidget(self.btn_next_day)
        title_bar.addWidget(self.btn_close)
        self.frame_layout.addLayout(title_bar)
        
        # Separator line
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("background-color: rgba(255, 255, 255, 15); height: 1px; border: none;")
        self.frame_layout.addWidget(sep)

        # STACKED WIDGET (Toggle between list and timer views)
        self.stacked_widget = QStackedWidget()
        self.frame_layout.addWidget(self.stacked_widget)
        
        # SCREEN 0: Tasks list scroll area
        self.init_tasks_list_screen()
        
        # SCREEN 1: Study Timer screen
        self.init_study_timer_screen()
        
        layout.addWidget(self.frame)
        
        # Custom SizeGrip in Bottom-Right Corner (Fully Responsive)
        self.size_grip_layout = QHBoxLayout()
        self.size_grip_layout.addStretch()
        self.size_grip = QSizeGrip(self)
        self.size_grip.setStyleSheet("width: 14px; height: 14px; background: transparent;")
        self.size_grip_layout.addWidget(self.size_grip)
        self.frame_layout.addLayout(self.size_grip_layout)
        
        # Local updates QTimer
        self.widget_timer = QTimer(self)
        self.widget_timer.timeout.connect(self.tick_timer)
        
        # Border alert QTimer
        self.flash_timer = QTimer(self)
        self.flash_timer.timeout.connect(self.flash_border)
        self.flash_state = False

        # Connect internal signal for thread-safe UI updates
        self.widget_data_received.connect(self.update_widget_ui)
        self.start_listener()

    def init_tasks_list_screen(self):
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setStyleSheet("background: transparent; border: none;")
        self.scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        
        self.scroll_content = QWidget()
        self.scroll_layout = QVBoxLayout(self.scroll_content)
        self.scroll_layout.setContentsMargins(0, 0, 0, 0)
        self.scroll_layout.setSpacing(8)
        
        self.scroll.setWidget(self.scroll_content)
        self.stacked_widget.addWidget(self.scroll)

    def init_study_timer_screen(self):
        self.timer_panel = QWidget()
        tpl = QVBoxLayout(self.timer_panel)
        tpl.setContentsMargins(0, 10, 0, 10)
        tpl.setSpacing(12)
        
        # Back button
        self.btn_back_to_list = QPushButton("◀ Back to Routine")
        self.btn_back_to_list.setObjectName("navBtn")
        self.btn_back_to_list.clicked.connect(self.back_to_list)
        tpl.addWidget(self.btn_back_to_list, alignment=Qt.AlignmentFlag.AlignLeft)
        
        # Active subject label
        self.lbl_timer_subject = QLabel("Subject")
        self.lbl_timer_subject.setStyleSheet("font-size: 13px; font-weight: bold; color: #EAEAEA;")
        self.lbl_timer_subject.setWordWrap(True)
        tpl.addWidget(self.lbl_timer_subject)
        
        tpl.addStretch()
        
        # Countdown text
        self.lbl_countdown = QLabel("00:00:00")
        self.lbl_countdown.setObjectName("timerLabel")
        self.lbl_countdown.setAlignment(Qt.AlignmentFlag.AlignCenter)
        tpl.addWidget(self.lbl_countdown)
        
        # Progress Bar
        self.progress_bar = QProgressBar()
        tpl.addWidget(self.progress_bar)
        
        # Controls Row
        c_row = QHBoxLayout()
        self.btn_play_pause = QPushButton("▶ Start")
        self.btn_play_pause.setObjectName("controlBtn")
        self.btn_play_pause.clicked.connect(self.toggle_play_pause)
        
        self.btn_reset_timer = QPushButton("🔄 Reset")
        self.btn_reset_timer.setStyleSheet("""
            QPushButton {
                background-color: rgba(255, 255, 255, 12);
                border: 1px solid rgba(255, 255, 255, 20);
                border-radius: 6px;
                padding: 6px 14px;
                color: #DDD;
            }
            QPushButton:hover {
                background-color: rgba(255, 255, 255, 20);
            }
        """)
        self.btn_reset_timer.clicked.connect(self.reset_subject_timer)
        
        c_row.addWidget(self.btn_play_pause)
        c_row.addWidget(self.btn_reset_timer)
        tpl.addLayout(c_row)
        
        tpl.addStretch()
        
        # Direct task check toggle
        self.btn_mark_complete_timer = QPushButton("✅ Mark Completed")
        self.btn_mark_complete_timer.setObjectName("controlBtn")
        self.btn_mark_complete_timer.setStyleSheet("""
            QPushButton {
                background-color: rgba(76, 175, 80, 160);
                border: 1px solid #4CAF50;
                border-radius: 6px;
                padding: 8px;
                font-weight: bold;
                color: white;
            }
            QPushButton:hover {
                background-color: rgba(76, 175, 80, 220);
            }
        """)
        self.btn_mark_complete_timer.clicked.connect(self.mark_complete_from_timer)
        tpl.addWidget(self.btn_mark_complete_timer)
        
        self.stacked_widget.addWidget(self.timer_panel)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            pos = event.position().toPoint()
            self.resize_dir = self.get_resize_direction(pos)
            if self.resize_dir:
                self.drag_start_geo = self.geometry()
                self.drag_start_pos = event.globalPosition().toPoint()
            else:
                # Drag the window
                self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()

    def mouseReleaseEvent(self, event):
        self.resize_dir = None
        self.drag_pos = QPoint()
        event.accept()

    def mouseMoveEvent(self, event):
        pos = event.position().toPoint()
        global_pos = event.globalPosition().toPoint()
        
        # If not holding left button, change cursor shape based on borders
        if event.buttons() != Qt.MouseButton.LeftButton:
            r_dir = self.get_resize_direction(pos)
            if r_dir == "bottom-right":
                self.setCursor(Qt.CursorShape.SizeBDiagCursor)
            elif r_dir == "bottom":
                self.setCursor(Qt.CursorShape.SizeVerCursor)
            elif r_dir == "right":
                self.setCursor(Qt.CursorShape.SizeHorCursor)
            else:
                self.setCursor(Qt.CursorShape.ArrowCursor)
            event.accept()
            return

        # If resizing
        if self.resize_dir:
            delta = global_pos - self.drag_start_pos
            if self.resize_dir == "right":
                new_width = max(self.minimumWidth(), self.drag_start_geo.width() + delta.x())
                self.resize(new_width, self.height())
            elif self.resize_dir == "bottom":
                new_height = max(self.minimumHeight(), self.drag_start_geo.height() + delta.y())
                self.resize(self.width(), new_height)
            elif self.resize_dir == "bottom-right":
                new_width = max(self.minimumWidth(), self.drag_start_geo.width() + delta.x())
                new_height = max(self.minimumHeight(), self.drag_start_geo.height() + delta.y())
                self.resize(new_width, new_height)
            event.accept()
        else:
            # If dragging window
            if not self.drag_pos.isNull():
                self.move(global_pos - self.drag_pos)
                event.accept()

    def get_resize_direction(self, pos):
        w = self.width()
        h = self.height()
        m = self.resize_margin
        
        # Check bottom-right corner
        if pos.x() >= w - m and pos.y() >= h - m:
            return "bottom-right"
        # Check right border
        elif pos.x() >= w - m:
            return "right"
        # Check bottom border
        elif pos.y() >= h - m:
            return "bottom"
        return None

    def start_listener(self):
        try:
            doc_ref = self.db.collection("user_routines").document(self.userId)
            def on_snapshot(doc_snapshot, changes, read_time):
                for doc in doc_snapshot:
                    if doc.exists:
                        data = doc.to_dict()
                        self.widget_data_received.emit(data)
                        return
                self.widget_data_received.emit({})
            self.listener = doc_ref.on_snapshot(on_snapshot)
        except Exception as e:
            print(f"Widget Firestore Listener Error: {e}")

    def closeEvent(self, event):
        if hasattr(self, 'listener') and self.listener:
            try: self.listener.unsubscribe()
            except: pass
        self.widget_timer.stop()
        self.flash_timer.stop()
        self.timer_active_sound = False
        
        # Remove from global references to let Python free memory
        global global_floating_widgets
        if self in global_floating_widgets:
            global_floating_widgets.remove(self)
        
        # Quit application if no other windows are visible
        visible_windows = [w for w in QApplication.topLevelWidgets() if w.isVisible() and w != self]
        if not visible_windows:
            QApplication.quit()
            
        event.accept()

    def prev_day(self):
        self.selected_day_offset = (self.selected_day_offset - 1) % 7
        if hasattr(self, 'last_data'):
            self.update_widget_ui(self.last_data)

    def next_day(self):
        self.selected_day_offset = (self.selected_day_offset + 1) % 7
        if hasattr(self, 'last_data'):
            self.update_widget_ui(self.last_data)

    def get_week_key_widget(self):
        import datetime
        now = datetime.datetime.now()
        jan1 = datetime.datetime(now.year, 1, 1)
        day_of_year = now.timetuple().tm_yday
        offset = (jan1.weekday() + 1) % 7
        week = (day_of_year + offset - 1) // 7 + 1
        return f"{week}-{now.year}"

    def update_widget_ui(self, data):
        self.last_data = data
        if not data:
            self.lbl_day.setText("No Data ⚪")
            # Clear layout
            for i in reversed(range(self.scroll_layout.count())):
                item = self.scroll_layout.takeAt(i)
                if item.widget(): item.widget().deleteLater()
            return
            
        days_list = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        
        from datetime import datetime
        current_wday = datetime.today().weekday()
        target_wday = (current_wday + self.selected_day_offset) % 7
        target_day_name = days_list[target_wday]
        
        prefix = "📅 TODAY" if self.selected_day_offset == 0 else f"📅 {target_day_name[:3].upper()}"
        self.lbl_day.setText(prefix)
        
        self.current_days = data.get("days", [])
        
        # Map raw days data by uppercase abbreviated dayName (e.g. 'MON', 'TUE')
        days_map = {d.get("dayName", "").strip().upper(): d for d in self.current_days}
        
        abbrev_map = {
            "monday": "MON", "tuesday": "TUE", "wednesday": "WED", 
            "thursday": "THU", "friday": "FRI", "saturday": "SAT", "sunday": "SUN"
        }
        abbrev_day = abbrev_map.get(target_day_name.lower(), "MON")
        d_data = days_map.get(abbrev_day, {})
        
        # Clear layout first
        for i in reversed(range(self.scroll_layout.count())):
            item = self.scroll_layout.takeAt(i)
            if item.widget(): item.widget().deleteLater()
            
        is_holiday = d_data.get("isHoliday", False)
        if is_holiday:
            lbl_hol = QLabel("🌴 Today is a Holiday!")
            lbl_hol.setStyleSheet("color: #FF5722; font-weight: bold; font-size: 13px; text-align: center; margin-top: 20px;")
            lbl_hol.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.scroll_layout.addWidget(lbl_hol)
            self.scroll_layout.addStretch()
            return
            
        completed_list = d_data.get("completedSubjects", [])
        
        # Collect all subjects for the day
        day_subjects = d_data.get("daySubjects", [])
        night_subjects = d_data.get("nightSubjects", [])
        
        all_subjects = []
        for i, sub in enumerate(day_subjects):
            if sub.strip():
                all_subjects.append((sub, "day", i))
        for i, sub in enumerate(night_subjects):
            if sub.strip():
                all_subjects.append((sub, "night", i))
                
        if not all_subjects:
            lbl_free = QLabel("Free Day / No Scheduled Subjects 📚")
            lbl_free.setStyleSheet("color: #888; font-style: italic; text-align: center; margin-top: 20px; font-size: 11px;")
            lbl_free.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.scroll_layout.addWidget(lbl_free)
        else:
            for sub, slot_type, idx in all_subjects:
                card = self.create_task_card(sub, completed_list, target_day_name, slot_type, idx)
                self.scroll_layout.addWidget(card)
                
        self.scroll_layout.addStretch()

    def create_task_card(self, sub_name, completed_list, day_name, slot_type, index):
        card_frame = QFrame()
        card_frame.setObjectName("taskCard")
        
        cl = QHBoxLayout(card_frame)
        cl.setContentsMargins(10, 10, 10, 10)
        cl.setSpacing(8)
        
        abbrev_map = {
            "monday": "MON", "tuesday": "TUE", "wednesday": "WED", 
            "thursday": "THU", "friday": "FRI", "saturday": "SAT", "sunday": "SUN"
        }
        day_abbrev = abbrev_map.get(day_name.lower(), "MON")
        week_key = self.get_week_key_widget()
        slot_key = f"{day_abbrev}|{'D' if slot_type == 'day' else 'N'}|{index}|{week_key}"
        
        is_completed = slot_key in completed_list
        
        # 1. Custom Checkbox Button
        btn_check = QPushButton()
        if is_completed:
            btn_check.setObjectName("checkCircleChecked")
            btn_check.setText("✓")
        else:
            btn_check.setObjectName("checkCircle")
            
        btn_check.clicked.connect(lambda checked, d=day_name, s=slot_type, idx=index: self.toggle_subject_status_by_key(d, s, idx))
        cl.addWidget(btn_check, alignment=Qt.AlignmentFlag.AlignVCenter)
        
        # 2. Subject Info Label
        lbl_info = QLabel(sub_name)
        lbl_info.setObjectName("taskLabel")
        lbl_info.setWordWrap(True)
        lbl_info.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        if is_completed:
            lbl_info.setStyleSheet("color: #888888; text-decoration: line-through;")
        cl.addWidget(lbl_info, alignment=Qt.AlignmentFlag.AlignVCenter)
        
        # 3. Timer Icon Button
        btn_timer = QPushButton("⏱️")
        btn_timer.setObjectName("timerIconBtn")
        btn_timer.setToolTip("Start Countdown Study Timer")
        btn_timer.clicked.connect(lambda checked, name=sub_name, s=slot_type, idx=index, d=day_name: self.open_subject_timer(name, s, idx, d))
        cl.addWidget(btn_timer, alignment=Qt.AlignmentFlag.AlignVCenter)
        
        return card_frame

    def parse_duration(self, text):
        # Extract hours: (2h) or (1.5h)
        match_h = re.search(r'\((\d+(\.\d+)?)\s*h\)', text, re.IGNORECASE)
        if match_h:
            return int(float(match_h.group(1)) * 3600)
            
        # Extract minutes: (45m)
        match_m = re.search(r'\((\d+)\s*m\)', text, re.IGNORECASE)
        if match_m:
            return int(match_m.group(1)) * 60
            
        # Default fallback: 1 Hour (3600 seconds)
        return 3600

    def open_subject_timer(self, subject_name, slot_type, index, day_name):
        self.timer_subject_name = subject_name
        self.timer_slot_type = slot_type
        self.timer_index = index
        self.timer_day_name = day_name
        self.timer_active_sound = False
        
        self.lbl_timer_subject.setText(subject_name)
        
        # Parse duration
        parsed_secs = self.parse_duration(subject_name)
        self.timer_total_seconds = parsed_secs
        self.timer_seconds_left = parsed_secs
        
        self.update_timer_display()
        self.progress_bar.setValue(100)
        
        # Style buttons reset
        self.btn_play_pause.setText("▶ Start")
        self.timer_running = False
        self.widget_timer.stop()
        self.flash_timer.stop()
        self.reset_widget_border()
        
        # Switch stacked screen to 1 (Timer View)
        self.stacked_widget.setCurrentIndex(1)

    def back_to_list(self):
        # Stop everything and return
        self.widget_timer.stop()
        self.flash_timer.stop()
        self.reset_widget_border()
        self.timer_active_sound = False
        self.stacked_widget.setCurrentIndex(0)

    def update_timer_display(self):
        h = self.timer_seconds_left // 3600
        m = (self.timer_seconds_left % 3600) // 60
        s = self.timer_seconds_left % 60
        self.lbl_countdown.setText(f"{h:02d}:{m:02d}:{s:02d}")
        
        # Update progress bar
        if self.timer_total_seconds > 0:
            pct = int((self.timer_seconds_left / self.timer_total_seconds) * 100)
            self.progress_bar.setValue(pct)
        else:
            self.progress_bar.setValue(0)

    def toggle_play_pause(self):
        if self.timer_running:
            self.widget_timer.stop()
            self.btn_play_pause.setText("▶ Start")
            self.timer_running = False
        else:
            if self.timer_seconds_left <= 0:
                self.reset_subject_timer()
            self.widget_timer.start(1000)
            self.btn_play_pause.setText("⏸ Pause")
            self.timer_running = True

    def reset_subject_timer(self):
        self.widget_timer.stop()
        self.flash_timer.stop()
        self.reset_widget_border()
        self.timer_active_sound = False
        self.timer_seconds_left = self.timer_total_seconds
        self.update_timer_display()
        self.btn_play_pause.setText("▶ Start")
        self.timer_running = False

    def tick_timer(self):
        if self.timer_seconds_left > 0:
            self.timer_seconds_left -= 1
            self.update_timer_display()
            
            if self.timer_seconds_left == 0:
                self.widget_timer.stop()
                self.btn_play_pause.setText("▶ Start")
                self.timer_running = False
                
                # Sound alarm & flash alert
                self.timer_active_sound = True
                self.play_alarm_sequence(5)
                self.flash_timer.start(500)
        else:
            self.widget_timer.stop()

    def play_alarm_sequence(self, count=4):
        if count <= 0 or not self.timer_active_sound: return
        import winsound
        try: winsound.Beep(1200, 350)
        except: pass
        QTimer.singleShot(600, lambda: self.play_alarm_sequence(count - 1))

    def flash_border(self):
        if self.flash_state:
            self.frame.setStyleSheet("""
                QFrame#widgetFrame {
                    background-color: rgba(20, 20, 20, 120);
                    border: 2px solid #FF5722;
                    border-radius: 16px;
                }
            """)
            self.flash_state = False
        else:
            self.reset_widget_border()
            self.flash_state = True

    def reset_widget_border(self):
        self.frame.setStyleSheet("""
            QFrame#widgetFrame {
                background-color: rgba(20, 20, 20, 120);
                border: 2px solid rgba(0, 188, 212, 180);
                border-radius: 16px;
            }
        """)

    def mark_complete_from_timer(self):
        self.toggle_subject_status_by_key(self.timer_day_name, self.timer_slot_type, self.timer_index)
        # Play completion beep
        import winsound
        try: winsound.Beep(1500, 200)
        except: pass
        self.back_to_list()

    def toggle_subject_status_by_key(self, day_name, slot_type, index):
        abbrev_map = {
            "monday": "MON", "tuesday": "TUE", "wednesday": "WED", 
            "thursday": "THU", "friday": "FRI", "saturday": "SAT", "sunday": "SUN"
        }
        day_abbrev = abbrev_map.get(day_name.lower(), "MON")
        week_key = self.get_week_key_widget()
        slot_key = f"{day_abbrev}|{'D' if slot_type == 'day' else 'N'}|{index}|{week_key}"

        try:
            updated_days = []
            for d in self.current_days:
                d_copy = dict(d)
                if d_copy.get("dayName", "").strip().upper() == day_abbrev:
                    comp = list(d_copy.get("completedSubjects", []))
                    if slot_key in comp:
                        comp.remove(slot_key)
                    else:
                        comp.append(slot_key)
                    d_copy["completedSubjects"] = comp
                updated_days.append(d_copy)

            doc_ref = self.db.collection("user_routines").document(self.userId)
            doc_ref.set({"days": updated_days}, merge=True)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to update task status from widget: {e}")

class MainWindow(QMainWindow):
    routine_data_received = pyqtSignal(dict)
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Copy Pro - Smart MCQ Extractor")
        self.resize(1150, 650)
        
        try:
            self.ocr_engine = WindowsOCR()
        except:
            self.ocr_engine = None
        self.parser = MCQParser()
        self.overlay = None
        self.current_target = None
        self.drag_pos = QPoint()
        self.db = None
        self.init_firebase()
        
        # Connect real-time routine signal
        self.routine_data_received.connect(self.update_routine_ui)
        
        self.init_ui()

    def init_firebase(self):
        if not FIREBASE_AVAILABLE: return
        kp = "serviceAccountKey.json"
        if os.path.exists(kp):
            try:
                cred = credentials.Certificate(kp)
                firebase_admin.initialize_app(cred)
                self.db = firestore.client()
            except: pass

    def init_ui(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #121212; }
            QWidget#centralWidget {
                background-color: #1a1a1a;
                border: 1px solid rgba(255, 255, 255, 25);
                border-radius: 12px;
            }
            QWidget { color: #FFFFFF; font-family: 'Segoe UI', sans-serif; font-size: 13px; }
            QTableWidget {
                background-color: rgba(30, 30, 30, 150);
                gridline-color: rgba(255, 255, 255, 20);
                border: 1px solid rgba(255, 255, 255, 30);
                border-radius: 10px;
                selection-background-color: rgba(0, 120, 212, 100);
            }
            QHeaderView::section {
                background-color: rgba(45, 45, 45, 180);
                color: #EAEAEA; border: none; padding: 5px; font-weight: bold;
            }
            QPushButton {
                background-color: rgba(255, 255, 255, 20);
                border: 1px solid rgba(255, 255, 255, 30);
                border-radius: 8px; padding: 6px 12px; color: white;
            }
            QPushButton:hover { background-color: rgba(255, 255, 255, 40); }
            QPushButton#accent { background-color: rgba(0, 120, 212, 150); }
            QPushButton#firebase { background-color: rgba(255, 152, 0, 150); }
            QPushButton#success { background-color: rgba(76, 175, 80, 150); }
            QTextEdit, QComboBox {
                background-color: rgba(40, 40, 40, 180);
                border: 1px solid rgba(255, 255, 255, 30);
                border-radius: 8px; padding: 5px; color: white;
            }
            QTabWidget::pane { border: 1px solid rgba(255, 255, 255, 30); background: transparent; border-radius: 10px; }
            QTabBar::tab {
                background: rgba(45, 45, 45, 150); padding: 10px 20px; border-top-left-radius: 10px; border-top-right-radius: 10px;
                margin-right: 2px; color: #BBB;
            }
            QTabBar::tab:selected { background: rgba(0, 120, 212, 180); color: white; font-weight: bold; }
        """)
        self.central_widget = QWidget()
        self.central_widget.setObjectName("centralWidget")
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(15, 15, 15, 15)
        self.main_layout.setSpacing(10)

        # 🇧🇩 Bengali OCR Toggle
        options_bar = QHBoxLayout()
        self.cb_bengali_ocr = QCheckBox("🇧🇩 Bengali OCR (Bengali + English)")
        self.cb_bengali_ocr.setStyleSheet("QCheckBox { font-weight: bold; color: #00BCD4; font-size: 13px; }")
        self.cb_bengali_ocr.stateChanged.connect(self.update_ocr_language)
        options_bar.addWidget(self.cb_bengali_ocr)
        options_bar.addStretch()
        self.main_layout.addLayout(options_bar)

        # 📑 Tab Widget
        self.tabs = QTabWidget()
        self.main_layout.addWidget(self.tabs)

        # --- TAB 1: MCQ EXTRACTOR ---
        self.extractor_widget = QWidget()
        self.extractor_layout = QVBoxLayout(self.extractor_widget)
        self.tabs.addTab(self.extractor_widget, "🔍 MCQ Extractor")
        
        self.toolbar_widget = QWidget()
        self.toolbar_layout = QHBoxLayout(self.toolbar_widget)

        self.toolbar_layout.setContentsMargins(0,0,0,0)
        self.toolbar_layout.setSpacing(6)
        
        self.btn_new_row = QPushButton("➕")
        self.btn_new_row.clicked.connect(self.add_blank_row)
        self.btn_cap_q = QPushButton("🔍 Q"); self.btn_cap_q.setObjectName("accent")
        self.btn_cap_q.clicked.connect(lambda: self.start_capture("Question"))
        self.btn_cap_o = QPushButton("📄 All Opt"); self.btn_cap_o.clicked.connect(lambda: self.start_capture("Options"))
        self.btn_a = QPushButton("A"); self.btn_a.clicked.connect(lambda: self.quick_set_answer("A"))
        self.btn_b = QPushButton("B"); self.btn_b.clicked.connect(lambda: self.quick_set_answer("B"))
        self.btn_c = QPushButton("C"); self.btn_c.clicked.connect(lambda: self.quick_set_answer("C"))
        self.btn_d = QPushButton("D"); self.btn_d.clicked.connect(lambda: self.quick_set_answer("D"))
        self.btn_cap_a = QPushButton("✅ Ans"); self.btn_cap_a.setObjectName("firebase")
        self.btn_cap_a.clicked.connect(lambda: self.start_capture("Answer"))
        self.btn_cap_s = QPushButton("💡 Sol"); self.btn_cap_s.clicked.connect(lambda: self.start_capture("Solution"))
        self.btn_cap_img = QPushButton("🖼️ Cap Img"); self.btn_cap_img.clicked.connect(lambda: self.start_capture("Image"))
        self.btn_upload_img = QPushButton("📤 Upload Q-Img"); self.btn_upload_img.clicked.connect(lambda: self.upload_question_image())
        self.btn_bulk = QPushButton("📋 Bulk"); self.btn_bulk.clicked.connect(self.open_bulk_import)
        self.btn_export_app = QPushButton("📤 Simple Q"); self.btn_export_app.setObjectName("accent")
        self.btn_export_app.clicked.connect(self.export_to_simple_q)
        self.btn_sync_cloud = QPushButton("🔥 Cloud Sync"); self.btn_sync_cloud.setObjectName("firebase")
        self.btn_sync_cloud.clicked.connect(self.sync_to_firebase)
        
        self.btn_history = QPushButton("☁️ History"); self.btn_history.setObjectName("accent")
        self.btn_history.clicked.connect(self.open_history)
        
        self.btn_cap_t = QPushButton("📋 Type"); self.btn_cap_t.clicked.connect(lambda: self.start_capture("Type"))
        self.btn_mini = QPushButton("📌 Mini"); self.btn_mini.setCheckable(True); self.btn_mini.clicked.connect(self.toggle_mini_mode)
        
        self.btn_add_to_mock = QPushButton("📥 Add to Mock"); self.btn_add_to_mock.setObjectName("success")
        self.btn_add_to_mock.clicked.connect(self.transfer_to_mock)
        
        self.btn_clear_extractor = QPushButton("🧹 Clear"); self.btn_clear_extractor.clicked.connect(lambda: self.table.setRowCount(0))
        
        self.combo_subject = QComboBox()
        self.combo_subject.addItems(["Mathematics", "Science", "History", "English 📚", "General Knowledge", "Technology", "Other"])
        self.combo_subject.currentIndexChanged.connect(self.update_subject_theme)
        self.btn_export = QPushButton("💾 Export"); self.btn_export.setObjectName("success"); self.btn_export.clicked.connect(self.export_excel)
        
        for b in [self.btn_new_row, self.btn_cap_q, self.btn_cap_o, self.btn_a, self.btn_b, self.btn_c, self.btn_d, self.btn_cap_a, self.btn_cap_s, self.btn_cap_img, self.btn_upload_img, self.btn_bulk, self.btn_export_app, self.btn_sync_cloud, self.btn_history, self.btn_add_to_mock, self.btn_clear_extractor, self.btn_cap_t, self.btn_mini]:
            self.toolbar_layout.addWidget(b)
        self.toolbar_layout.addWidget(self.combo_subject)
        self.toolbar_layout.addStretch()
        self.toolbar_layout.addWidget(self.btn_export)
        self.extractor_layout.addWidget(self.toolbar_widget)
        
        self.table_container = QWidget()
        tl = QVBoxLayout(self.table_container); tl.setContentsMargins(0,0,0,0)
        self.table = QTableWidget(0, 11)
        self.table.setHorizontalHeaderLabels(["Question", "Option A", "Option B", "Option C", "Option D", "Correct Answer", "Solution", "Type", "Image", "Action", "doc_id"])
        self.table.setColumnHidden(10, True) # 🤫 Hide doc_id column
        self.table.setColumnWidth(8, 70)     # 🖼️ Set width for Image status
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        self.table.setWordWrap(True); self.table.verticalHeader().setDefaultSectionSize(35)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        tl.addWidget(self.table)
        self.editor_label = QLabel("📝 Quick Editor:"); self.editor_label.setStyleSheet("font-weight: bold; margin-top: 5px;")
        tl.addWidget(self.editor_label)
        self.detail_editor = QTextEdit(); self.detail_editor.setMaximumHeight(80)
        self.detail_editor.textChanged.connect(self.sync_editor_to_table)
        tl.addWidget(self.detail_editor)
        self.extractor_layout.addWidget(self.table_container)
        self.table.currentCellChanged.connect(self.sync_table_to_editor)
        
        # --- TAB 2: MOCK CREATOR ---
        self.init_mock_tab()
        
        # --- TAB 3: DAILY PRACTICE ---
        self.init_daily_tab()
        
        # --- TAB 4: LAPTOP QUIZ ---
        self.init_quiz_play_tab()
        
        # --- TAB 5: MY ROUTINE ---
        self.init_routine_tab()
        
        self.add_blank_row()

    def init_mock_tab(self):
        self.mock_widget = QWidget()
        self.mock_layout = QVBoxLayout(self.mock_widget)
        self.tabs.addTab(self.mock_widget, "🏆 Mock Creator")

        # Mock Header
        header = QHBoxLayout()
        self.mock_test_name = QTextEdit()
        self.mock_test_name.setPlaceholderText("Enter Mock Test Name (e.g. SSC CGL Mock 01)")
        self.mock_test_name.setMaximumHeight(40)
        header.addWidget(QLabel("Test Name:"))
        header.addWidget(self.mock_test_name)
        
        self.btn_sync_mock = QPushButton("🔥 Sync Mock Set")
        self.btn_sync_mock.setObjectName("firebase")
        self.btn_sync_mock.clicked.connect(self.sync_mock_to_cloud)
        header.addWidget(self.btn_sync_mock)

        self.btn_bulk_mock = QPushButton("📋 Bulk Import (Mock)")
        self.btn_bulk_mock.clicked.connect(self.open_bulk_mock_import)
        header.addWidget(self.btn_bulk_mock)

        self.btn_clear_mock = QPushButton("🧹 Clear All")
        self.btn_clear_mock.clicked.connect(lambda: self.mock_table.setRowCount(0))
        header.addWidget(self.btn_clear_mock)

        self.btn_mock_history = QPushButton("☁️ Mock History"); self.btn_mock_history.setObjectName("accent")
        self.btn_mock_history.clicked.connect(self.open_history)
        header.addWidget(self.btn_mock_history)
        
        self.btn_mock_export = QPushButton("💾 Export Excel"); self.btn_mock_export.setObjectName("success")
        self.btn_mock_export.clicked.connect(self.export_mock_excel)
        header.addWidget(self.btn_mock_export)
        
        self.mock_layout.addLayout(header)

        # 🔍 Mock OCR Toolbar
        self.mock_toolbar = QHBoxLayout()
        self.btn_mock_q = QPushButton("🔍 Q"); self.btn_mock_q.setObjectName("accent")
        self.btn_mock_q.clicked.connect(lambda: self.start_capture("Question"))
        self.btn_mock_o = QPushButton("📄 All Opt"); self.btn_mock_o.clicked.connect(lambda: self.start_capture("Options"))
        self.btn_mock_a = QPushButton("A"); self.btn_mock_a.clicked.connect(lambda: self.quick_set_answer("A"))
        self.btn_mock_b = QPushButton("B"); self.btn_mock_b.clicked.connect(lambda: self.quick_set_answer("B"))
        self.btn_mock_c = QPushButton("C"); self.btn_mock_c.clicked.connect(lambda: self.quick_set_answer("C"))
        self.btn_mock_d = QPushButton("D"); self.btn_mock_d.clicked.connect(lambda: self.quick_set_answer("D"))
        self.btn_mock_ans = QPushButton("✅ Ans"); self.btn_mock_ans.setObjectName("firebase")
        self.btn_mock_ans.clicked.connect(lambda: self.start_capture("Answer"))
        self.btn_mock_sol = QPushButton("💡 Sol"); self.btn_mock_sol.clicked.connect(lambda: self.start_capture("Solution"))
        
        self.btn_mock_img = QPushButton("🖼️ Cap Q-Img"); self.btn_mock_img.setObjectName("firebase")
        self.btn_mock_img.clicked.connect(lambda: self.start_capture("MockImage"))
        
        self.btn_mock_upload = QPushButton("📤 Upload Q-Img"); self.btn_mock_upload.clicked.connect(lambda: self.upload_question_image())
        
        for btn in [self.btn_mock_q, self.btn_mock_o, self.btn_mock_a, self.btn_mock_b, self.btn_mock_c, self.btn_mock_d, self.btn_mock_ans, self.btn_mock_sol, self.btn_mock_img, self.btn_mock_upload]:
            self.mock_toolbar.addWidget(btn)
        self.mock_toolbar.addStretch()
        self.mock_layout.addLayout(self.mock_toolbar)

        # Mock Table
        self.mock_table = QTableWidget(0, 10)
        self.mock_table.setHorizontalHeaderLabels(["Question", "A", "B", "C", "D", "Ans", "Solution", "Image", "Category", "Action"])
        self.mock_table.setColumnWidth(7, 70) # 🖼️ Image
        self.mock_table.setColumnWidth(8, 110) # 📂 Category
        self.mock_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.mock_layout.addWidget(self.mock_table)
        
        # Add Row Button
        self.btn_add_mock_row = QPushButton("➕ Add Mock Question")
        self.btn_add_mock_row.clicked.connect(self.add_mock_row)
        self.mock_layout.addWidget(self.btn_add_mock_row)

    def init_daily_tab(self):
        self.daily_widget = QWidget()
        self.daily_layout = QVBoxLayout(self.daily_widget)
        self.tabs.addTab(self.daily_widget, "📅 Daily Practice")
        
        # Header
        header = QHBoxLayout()
        title = QLabel("📅 Daily Practice Vocabulary Manager")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #FF9800;")
        header.addWidget(title); header.addStretch()
        
        self.btn_sync_daily = QPushButton("🔥 Sync Daily Practice"); self.btn_sync_daily.setObjectName("firebase")
        self.btn_sync_daily.clicked.connect(self.sync_daily_to_cloud)
        header.addWidget(self.btn_sync_daily)
        
        self.btn_bulk_daily = QPushButton("📋 Bulk Import (Vocab)")
        self.btn_bulk_daily.clicked.connect(self.open_bulk_daily_import)
        header.addWidget(self.btn_bulk_daily)
        
        self.btn_clear_daily = QPushButton("🧹 Clear All")
        self.btn_clear_daily.clicked.connect(lambda: self.daily_table.setRowCount(0))
        header.addWidget(self.btn_clear_daily)
        
        self.daily_layout.addLayout(header)

        # OCR Toolbar for Daily Practice
        self.daily_toolbar = QHBoxLayout()
        self.btn_daily_word = QPushButton("🏷️ Word"); self.btn_daily_word.setObjectName("accent")
        self.btn_daily_word.clicked.connect(lambda: self.start_capture("Word"))
        
        self.btn_daily_mean = QPushButton("📖 Meaning"); self.btn_daily_mean.clicked.connect(lambda: self.start_capture("Meaning"))
        self.btn_daily_phrases = QPushButton("💬 Phrases"); self.btn_daily_phrases.clicked.connect(lambda: self.start_capture("Phrases"))
        self.btn_daily_story = QPushButton("📜 Story"); self.btn_daily_story.clicked.connect(lambda: self.start_capture("Story"))
        
        for btn in [self.btn_daily_word, self.btn_daily_mean, self.btn_daily_phrases, self.btn_daily_story]:
            self.daily_toolbar.addWidget(btn)
        self.daily_toolbar.addStretch()
        self.daily_layout.addLayout(self.daily_toolbar)

        # Table
        self.daily_table = QTableWidget(0, 7)
        self.daily_table.setHorizontalHeaderLabels(["Word", "Meaning", "Phrases", "Story", "Quiz (JSON)", "Order", "Action"])
        self.daily_table.setColumnWidth(0, 150)
        self.daily_table.setColumnWidth(1, 250)
        self.daily_table.setColumnWidth(4, 200) # Quiz JSON
        self.daily_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.daily_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.daily_layout.addWidget(self.daily_table)
        
        self.btn_add_daily_row = QPushButton("➕ Add New Word")
        self.btn_add_daily_row.clicked.connect(self.add_daily_row)
        self.daily_layout.addWidget(self.btn_add_daily_row)

    def add_daily_row(self):
        r = self.daily_table.rowCount(); self.daily_table.insertRow(r)
        for i in range(6): self.daily_table.setItem(r, i, QTableWidgetItem(""))
        self.daily_table.setItem(r, 5, QTableWidgetItem(str(r + 1))) # Initial order
        
        del_btn = QPushButton("🗑️")
        del_btn.clicked.connect(lambda _, row=r: self.daily_table.removeRow(self.daily_table.currentRow() if self.daily_table.currentRow() >=0 else row))
        self.daily_table.setCellWidget(r, 6, del_btn)

    def init_quiz_play_tab(self):
        self.quiz_play_widget = QWidget()
        self.quiz_play_layout = QVBoxLayout(self.quiz_play_widget)
        self.tabs.addTab(self.quiz_play_widget, "🎮 Laptop Quiz")

        # Header
        header = QHBoxLayout()
        title = QLabel("🎮 Local Laptop Quiz Player & Creator")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #00BCD4;")
        header.addWidget(title)
        header.addStretch()

        self.btn_play_quiz = QPushButton("🎮 Play Quiz")
        self.btn_play_quiz.setObjectName("success")
        self.btn_play_quiz.clicked.connect(self.play_local_quiz)
        header.addWidget(self.btn_play_quiz)

        self.btn_save_quiz = QPushButton("💾 Save Local")
        self.btn_save_quiz.clicked.connect(self.save_local_quiz)
        header.addWidget(self.btn_save_quiz)

        self.btn_load_quiz = QPushButton("📂 Load Local")
        self.btn_load_quiz.clicked.connect(self.load_local_quiz)
        header.addWidget(self.btn_load_quiz)

        self.btn_sync_quiz_cloud = QPushButton("🔥 Cloud Sync")
        self.btn_sync_quiz_cloud.setObjectName("firebase")
        self.btn_sync_quiz_cloud.clicked.connect(self.sync_quiz_to_firebase)
        header.addWidget(self.btn_sync_quiz_cloud)

        self.btn_quiz_cloud_history = QPushButton("☁️ Cloud History")
        self.btn_quiz_cloud_history.setObjectName("accent")
        self.btn_quiz_cloud_history.clicked.connect(self.open_history)
        header.addWidget(self.btn_quiz_cloud_history)

        self.btn_bulk_quiz = QPushButton("📋 Bulk Import")
        self.btn_bulk_quiz.clicked.connect(self.open_bulk_quiz_import)
        header.addWidget(self.btn_bulk_quiz)

        self.btn_clear_quiz = QPushButton("🧹 Clear All")
        self.btn_clear_quiz.clicked.connect(lambda: self.quiz_play_table.setRowCount(0))
        header.addWidget(self.btn_clear_quiz)

        self.quiz_play_layout.addLayout(header)

        # OCR Toolbar for Quiz Play
        self.quiz_toolbar = QHBoxLayout()
        self.btn_quiz_q = QPushButton("🔵 Q"); self.btn_quiz_q.setObjectName("accent")
        self.btn_quiz_q.clicked.connect(lambda: self.start_capture("QuizQuestion"))

        self.btn_quiz_opt = QPushButton("📄 All Opt"); self.btn_quiz_opt.clicked.connect(lambda: self.start_capture("QuizOptions"))
        self.btn_quiz_a = QPushButton("A"); self.btn_quiz_a.clicked.connect(lambda: self.start_capture("QuizOptionA"))
        self.btn_quiz_b = QPushButton("B"); self.btn_quiz_b.clicked.connect(lambda: self.start_capture("QuizOptionB"))
        self.btn_quiz_c = QPushButton("C"); self.btn_quiz_c.clicked.connect(lambda: self.start_capture("QuizOptionC"))
        self.btn_quiz_d = QPushButton("D"); self.btn_quiz_d.clicked.connect(lambda: self.start_capture("QuizOptionD"))

        self.btn_quiz_ans = QPushButton("✅ Ans"); self.btn_quiz_ans.setObjectName("firebase")
        self.btn_quiz_ans.clicked.connect(lambda: self.start_capture("QuizAnswer"))

        self.btn_quiz_sol = QPushButton("💡 Sol"); self.btn_quiz_sol.clicked.connect(lambda: self.start_capture("QuizSolution"))

        self.btn_quiz_img = QPushButton("🖼️ Cap Q-Img"); self.btn_quiz_img.setObjectName("firebase")
        self.btn_quiz_img.clicked.connect(lambda: self.start_capture("QuizImage"))

        self.btn_quiz_upload = QPushButton("📤 Upload Q-Img"); self.btn_quiz_upload.clicked.connect(lambda: self.upload_quiz_question_image())

        for btn in [self.btn_quiz_q, self.btn_quiz_opt, self.btn_quiz_a, self.btn_quiz_b, self.btn_quiz_c, self.btn_quiz_d, self.btn_quiz_ans, self.btn_quiz_sol, self.btn_quiz_img, self.btn_quiz_upload]:
            self.quiz_toolbar.addWidget(btn)
        self.quiz_toolbar.addStretch()
        self.quiz_play_layout.addLayout(self.quiz_toolbar)

        # Table
        self.quiz_play_table = QTableWidget(0, 9)
        self.quiz_play_table.setHorizontalHeaderLabels(["Question", "A", "B", "C", "D", "Ans", "Solution", "Image", "Action"])
        self.quiz_play_table.setColumnWidth(7, 70) # Image
        self.quiz_play_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        self.quiz_play_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.quiz_play_table.setWordWrap(True)
        self.quiz_play_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.quiz_play_layout.addWidget(self.quiz_play_table)

        # Add Row Button
        self.btn_add_quiz_play_row = QPushButton("➕ Add Quiz Question")
        self.btn_add_quiz_play_row.clicked.connect(self.add_quiz_play_row)
        self.quiz_play_layout.addWidget(self.btn_add_quiz_play_row)
        
        self.add_quiz_play_row()

    def add_quiz_play_row(self):
        r = self.quiz_play_table.rowCount(); self.quiz_play_table.insertRow(r)
        for i in range(8): self.quiz_play_table.setItem(r, i, QTableWidgetItem(""))
        del_btn = QPushButton("🗑️")
        del_btn.clicked.connect(lambda _, row=r: self.quiz_play_table.removeRow(self.quiz_play_table.currentRow() if self.quiz_play_table.currentRow() >=0 else row))
        self.quiz_play_table.setCellWidget(r, 8, del_btn)

    def upload_quiz_question_image(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Question Image", "", "Images (*.png *.jpg *.jpeg)")
        if not file_path: return
        try:
            base64_str = self.process_image_to_base64(file_path)
            r = self.quiz_play_table.currentRow()
            if r < 0: r = self.quiz_play_table.rowCount() - 1
            if r >= 0:
                item = QTableWidgetItem("🖼️ Ready")
                item.setData(Qt.ItemDataRole.UserRole, base64_str)
                item.setToolTip("Image Uploaded Successfully")
                self.quiz_play_table.setItem(r, 7, item)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to upload image: {e}")

    def save_local_quiz(self):
        if self.quiz_play_table.rowCount() == 0:
            QMessageBox.warning(self, "No Data", "There are no questions in the quiz table to save.")
            return
        
        import os
        from datetime import datetime
        target_dir = os.path.join(os.path.expanduser("~"), "Downloads", "Daily quiz")
        if not os.path.exists(target_dir):
            os.makedirs(target_dir, exist_ok=True)

        # Collect only non-empty rows
        quiz_list = []
        for r in range(self.quiz_play_table.rowCount()):
            q_text = self.quiz_play_table.item(r, 0).text().strip() if self.quiz_play_table.item(r, 0) else ""
            if not q_text:
                continue
                
            img_item = self.quiz_play_table.item(r, 7)
            img_data = img_item.data(Qt.ItemDataRole.UserRole) if img_item else ""
            if not img_data and img_item: img_data = img_item.text()
            
            q_data = {
                "question": q_text,
                "optionA": self.quiz_play_table.item(r, 1).text().strip() if self.quiz_play_table.item(r, 1) else "",
                "optionB": self.quiz_play_table.item(r, 2).text().strip() if self.quiz_play_table.item(r, 2) else "",
                "optionC": self.quiz_play_table.item(r, 3).text().strip() if self.quiz_play_table.item(r, 3) else "",
                "optionD": self.quiz_play_table.item(r, 4).text().strip() if self.quiz_play_table.item(r, 4) else "",
                "answer": self.quiz_play_table.item(r, 5).text().strip() if self.quiz_play_table.item(r, 5) else "",
                "solution": self.quiz_play_table.item(r, 6).text().strip() if self.quiz_play_table.item(r, 6) else "",
                "questionImage": img_data
            }
            quiz_list.append(q_data)
            
        if not quiz_list:
            QMessageBox.warning(self, "No Valid Data", "The quiz table is empty or all questions are blank. Please add or import questions first.")
            return

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_file = os.path.join(target_dir, f"quiz_{timestamp}.json")
        file_path, selected_filter = QFileDialog.getSaveFileName(
            self, 
            "Save Quiz Locally", 
            default_file, 
            "JSON Files (*.json);;Excel Files (*.xlsx)"
        )
        if not file_path: return
        
        try:
            if file_path.endswith('.xlsx'):
                df = pd.DataFrame(quiz_list)
                df.to_excel(file_path, index=False)
            else:
                if not file_path.endswith('.json'): file_path += '.json'
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(quiz_list, f, indent=4, ensure_ascii=False)
                    
            QMessageBox.information(self, "Quiz Saved", f"✅ Successfully saved {len(quiz_list)} questions at:\n{file_path}")
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save quiz: {e}")

    def load_local_quiz(self):
        import os
        user_downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        target_dir = os.path.join(user_downloads, "Daily quiz") if os.path.exists(os.path.join(user_downloads, "Daily quiz")) else user_downloads

        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "Select Quiz File to Load", 
            target_dir, 
            "Supported Quiz Files (*.json *.xlsx *.xls *.tsv *.csv *.txt *.db *.sqlite *.sqlite3);;SQLite DB Files (*.db *.sqlite *.sqlite3);;JSON Files (*.json);;Excel Files (*.xlsx *.xls);;TSV Files (*.tsv);;CSV Files (*.csv);;Text Files (*.txt);;All Files (*.*)"
        )
        if not file_path: return
        
        try:
            lower_path = file_path.lower()
            quiz_list = []
            if lower_path.endswith(('.db', '.sqlite', '.sqlite3')):
                quiz_list = parse_sqlite_quiz_data(file_path)
            elif lower_path.endswith('.json'):
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                if isinstance(data, list):
                    raw_items = data
                elif isinstance(data, dict):
                    raw_items = [data]
                else:
                    raw_items = []
                for item in raw_items:
                    q = str(item.get("question", item.get("Question", ""))).strip()
                    if q or any(str(item.get(k, "")).strip() for k in ["optionA", "optionB", "A", "B"]):
                        quiz_list.append({
                            "question": q,
                            "optionA": str(item.get("optionA", item.get("Option A", item.get("A", "")))).strip(),
                            "optionB": str(item.get("optionB", item.get("Option B", item.get("B", "")))).strip(),
                            "optionC": str(item.get("optionC", item.get("Option C", item.get("C", "")))).strip(),
                            "optionD": str(item.get("optionD", item.get("Option D", item.get("D", "")))).strip(),
                            "answer": str(item.get("answer", item.get("Correct Answer", item.get("Ans", "")))).strip(),
                            "solution": str(item.get("solution", item.get("Solution", item.get("Sol", "")))).strip(),
                            "questionImage": item.get("questionImage", item.get("Image", ""))
                        })
            elif lower_path.endswith('.xlsx') or lower_path.endswith('.xls'):
                df = pd.read_excel(file_path).fillna('')
                for _, row in df.iterrows():
                    vals = [str(v).strip() for v in row.values]
                    if vals and vals[0].lower() in ['question', 'q', 'questions', 'mcq']:
                        continue
                    if any(vals):
                        q_data = {
                            "question": vals[0] if len(vals) > 0 else "",
                            "optionA": vals[1] if len(vals) > 1 else "",
                            "optionB": vals[2] if len(vals) > 2 else "",
                            "optionC": vals[3] if len(vals) > 3 else "",
                            "optionD": vals[4] if len(vals) > 4 else "",
                            "answer": vals[5] if len(vals) > 5 else "",
                            "solution": vals[6] if len(vals) > 6 else "",
                            "questionImage": vals[7] if len(vals) > 7 else ""
                        }
                        quiz_list.append(q_data)
            else: # TSV, CSV, TXT
                with open(file_path, 'r', encoding='utf-8', errors='replace') as fp:
                    raw_text = fp.read()
                d = BulkImportDialog(self)
                rows = d.parse_text_to_rows(raw_text, '\t')
                for parts in rows:
                    if not any(parts): continue
                    if parts and parts[0].lower() in ['question', 'q', 'questions', 'mcq']:
                        continue
                    q_data = {
                        "question": parts[0] if len(parts) > 0 else "",
                        "optionA": parts[1] if len(parts) > 1 else "",
                        "optionB": parts[2] if len(parts) > 2 else "",
                        "optionC": parts[3] if len(parts) > 3 else "",
                        "optionD": parts[4] if len(parts) > 4 else "",
                        "answer": parts[5] if len(parts) > 5 else "",
                        "solution": parts[6] if len(parts) > 6 else "",
                        "questionImage": parts[7] if len(parts) > 7 else ""
                    }
                    quiz_list.append(q_data)
            
            if not quiz_list:
                QMessageBox.warning(self, "Empty File", "No valid questions were found in this file (or the file contains only empty rows).")
                return

            self.quiz_play_table.setRowCount(0)
            for r, q_data in enumerate(quiz_list):
                self.quiz_play_table.insertRow(r)
                
                self.quiz_play_table.setItem(r, 0, QTableWidgetItem(str(q_data.get("question", ""))))
                self.quiz_play_table.setItem(r, 1, QTableWidgetItem(str(q_data.get("optionA", ""))))
                self.quiz_play_table.setItem(r, 2, QTableWidgetItem(str(q_data.get("optionB", ""))))
                self.quiz_play_table.setItem(r, 3, QTableWidgetItem(str(q_data.get("optionC", ""))))
                self.quiz_play_table.setItem(r, 4, QTableWidgetItem(str(q_data.get("optionD", ""))))
                self.quiz_play_table.setItem(r, 5, QTableWidgetItem(str(q_data.get("answer", ""))))
                self.quiz_play_table.setItem(r, 6, QTableWidgetItem(str(q_data.get("solution", ""))))
                
                img_data = q_data.get("questionImage", "")
                if img_data:
                    item = QTableWidgetItem("🖼️ Ready")
                    item.setData(Qt.ItemDataRole.UserRole, str(img_data))
                    item.setToolTip("Image Restored Successfully")
                    self.quiz_play_table.setItem(r, 7, item)
                else:
                    self.quiz_play_table.setItem(r, 7, QTableWidgetItem(""))
                
                del_btn = QPushButton("🗑️")
                del_btn.clicked.connect(lambda _, btn=del_btn: self.quiz_play_table.removeRow(self.quiz_play_table.indexAt(btn.pos()).row() if self.quiz_play_table.indexAt(btn.pos()).row() >= 0 else self.quiz_play_table.currentRow()))
                self.quiz_play_table.setCellWidget(r, 8, del_btn)
            
            self.quiz_play_table.viewport().update()
            QMessageBox.information(self, "Success", f"✅ Successfully loaded {len(quiz_list)} questions into Laptop Quiz!")
        except Exception as e:
            QMessageBox.critical(self, "Load Error", f"Failed to load quiz file: {e}")

    def sync_quiz_to_firebase(self):
        if not self.db:
            QMessageBox.warning(self, "Firebase", "Firebase database is not connected.")
            return
        if self.quiz_play_table.rowCount() == 0:
            QMessageBox.warning(self, "No Data", "Quiz table is empty. Add questions before syncing.")
            return
            
        d = SyncMetadataDialog("Laptop Quiz", self)
        if d.exec() == QDialog.DialogCode.Accepted:
            meta = d.get_data()
            meta_subject = meta["subject"].split(' ')[0]
            try:
                batch = self.db.batch()
                ts = firestore.SERVER_TIMESTAMP
                synced_count = 0
                for r in range(self.quiz_play_table.rowCount()):
                    q_text = self.quiz_play_table.item(r, 0).text().strip() if self.quiz_play_table.item(r, 0) else ""
                    if not q_text: continue
                    
                    img_item = self.quiz_play_table.item(r, 7)
                    img_data = img_item.data(Qt.ItemDataRole.UserRole) if img_item else ""
                    if not img_data and img_item: img_data = img_item.text()
                    
                    data = {
                        "question": q_text,
                        "optionA": self.quiz_play_table.item(r, 1).text() if self.quiz_play_table.item(r, 1) else "",
                        "optionB": self.quiz_play_table.item(r, 2).text() if self.quiz_play_table.item(r, 2) else "",
                        "optionC": self.quiz_play_table.item(r, 3).text() if self.quiz_play_table.item(r, 3) else "",
                        "optionD": self.quiz_play_table.item(r, 4).text() if self.quiz_play_table.item(r, 4) else "",
                        "answer": self.quiz_play_table.item(r, 5).text() if self.quiz_play_table.item(r, 5) else "",
                        "solution": self.quiz_play_table.item(r, 6).text() if self.quiz_play_table.item(r, 6) else "",
                        "questionImage": img_data,
                        "type": meta["type"],
                        "subject": meta_subject,
                        "topic": meta["topic"],
                        "test_code": meta["test_code"],
                        "is_complete": meta["is_complete"],
                        "timestamp": ts
                    }
                    content_id = hashlib.md5(f"{meta_subject}_{meta['topic']}_{q_text}".encode()).hexdigest()
                    batch.set(self.db.collection("quizzes").document(content_id), data, merge=True)
                    synced_count += 1
                batch.commit()
                QMessageBox.information(self, "Cloud Sync Successful", f"✅ Successfully synced {synced_count} questions of '{meta['topic']}' to Firebase Cloud!")
            except Exception as e:
                QMessageBox.critical(self, "Sync Error", f"Failed to sync to Firebase: {e}")

    def open_bulk_quiz_import(self):
        d = BulkImportDialog(self, title="Laptop Quiz Bulk Import", expected_format="Question | Option A | Option B | Option C | Option D | Answer | Solution", min_cols=1)
        if d.exec() == QDialog.DialogCode.Accepted:
            rows = d.get_rows()
            if not rows:
                QMessageBox.warning(self, "No Data", "No valid rows found to import.")
                return
            
            imported_count = 0
            for parts in rows:
                if not any(parts): continue
                r = self.quiz_play_table.rowCount()
                self.quiz_play_table.insertRow(r)
                
                for i in range(8):
                    self.quiz_play_table.setItem(r, i, QTableWidgetItem(""))
                    
                for i in range(min(len(parts), 7)):
                    self.quiz_play_table.setItem(r, i, QTableWidgetItem(str(parts[i]).strip()))
                
                # Format answer letter if present
                ans_text = self.quiz_play_table.item(r, 5).text().strip()
                m = re.search(r'([A-Da-d])', ans_text)
                if m and len(ans_text) > 1:
                    self.quiz_play_table.setItem(r, 5, QTableWidgetItem(m.group(1).upper()))
                    
                del_btn = QPushButton("🗑️")
                del_btn.clicked.connect(lambda _, btn=del_btn: self.quiz_play_table.removeRow(self.quiz_play_table.indexAt(btn.pos()).row() if self.quiz_play_table.indexAt(btn.pos()).row() >= 0 else self.quiz_play_table.currentRow()))
                self.quiz_play_table.setCellWidget(r, 8, del_btn)
                imported_count += 1
                
            self.quiz_play_table.viewport().update()
            QMessageBox.information(self, "Success", f"✅ Successfully imported {imported_count} quiz questions!")

    def play_local_quiz(self):
        if self.quiz_play_table.rowCount() == 0:
            QMessageBox.warning(self, "No Data", "There are no questions in the quiz table to play.")
            return

        questions = []
        for r in range(self.quiz_play_table.rowCount()):
            img_item = self.quiz_play_table.item(r, 7)
            img_data = img_item.data(Qt.ItemDataRole.UserRole) if img_item else ""
            if not img_data and img_item: img_data = img_item.text()

            # Skip completely empty rows
            q_text = self.quiz_play_table.item(r, 0).text().strip() if self.quiz_play_table.item(r, 0) else ""
            if not q_text and not img_data: continue

            q_data = {
                "question": q_text,
                "optionA": self.quiz_play_table.item(r, 1).text() if self.quiz_play_table.item(r, 1) else "",
                "optionB": self.quiz_play_table.item(r, 2).text() if self.quiz_play_table.item(r, 2) else "",
                "optionC": self.quiz_play_table.item(r, 3).text() if self.quiz_play_table.item(r, 3) else "",
                "optionD": self.quiz_play_table.item(r, 4).text() if self.quiz_play_table.item(r, 4) else "",
                "answer": self.quiz_play_table.item(r, 5).text() if self.quiz_play_table.item(r, 5) else "",
                "solution": self.quiz_play_table.item(r, 6).text() if self.quiz_play_table.item(r, 6) else "",
                "questionImage": img_data
            }
            questions.append(q_data)

        if not questions:
            QMessageBox.warning(self, "No Valid Data", "Please fill in at least one question before playing.")
            return

        dialog = QuizPlayerDialog(questions, self)
        dialog.exec()

    def update_ocr_language(self):
        if self.ocr_engine:
            self.ocr_engine.use_bengali = self.cb_bengali_ocr.isChecked()

    def init_routine_tab(self):
        self.routine_widget = QWidget()
        self.routine_layout = QVBoxLayout(self.routine_widget)
        self.tabs.addTab(self.routine_widget, "📅 My Routine")

        # Header: User ID input, Connect button, and Status
        header = QHBoxLayout()
        title = QLabel("📅 Live Timetable & Tasks")
        title.setStyleSheet("font-size: 18px; font-weight: bold; color: #00BCD4;")
        header.addWidget(title)
        header.addStretch()

        header.addWidget(QLabel("User ID:"))
        self.routine_uid_input = QTextEdit()
        self.routine_uid_input.setPlaceholderText("Enter Firestore User ID")
        self.routine_uid_input.setMaximumHeight(35)
        self.routine_uid_input.setMaximumWidth(280)
        
        # Load user ID from QSettings
        settings = QSettings("CopyPro", "RoutineConfig")
        saved_uid = settings.value("userId", "")
        self.routine_uid_input.setPlainText(saved_uid)
        header.addWidget(self.routine_uid_input)

        self.btn_connect_routine = QPushButton("🔄 Connect Routine")
        self.btn_connect_routine.setObjectName("firebase")
        self.btn_connect_routine.clicked.connect(self.start_routine_listener)
        header.addWidget(self.btn_connect_routine)

        self.btn_pin_widget = QPushButton("🖥️ Pin Desktop Widget")
        self.btn_pin_widget.setObjectName("accent")
        self.btn_pin_widget.clicked.connect(self.launch_routine_widget)
        header.addWidget(self.btn_pin_widget)

        self.lbl_routine_status = QLabel("Disconnected ⚪")
        self.lbl_routine_status.setStyleSheet("font-weight: bold; color: #888; font-size: 13px;")
        header.addWidget(self.lbl_routine_status)

        self.routine_layout.addLayout(header)

        # Table showing 7 days of the week with 7 columns
        self.routine_table = QTableWidget(7, 7)
        self.routine_table.setHorizontalHeaderLabels([
            "Day", "Day Sub 1", "Day Sub 2", "Day Sub 3", 
            "Night Sub 1", "Night Sub 2", "Night Sub 3"
        ])
        
        self.routine_table.setColumnWidth(0, 80)
        self.routine_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.routine_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        self.routine_table.verticalHeader().setDefaultSectionSize(70) # taller to fit wrapped text nicely
        
        # Populate Day Names initially
        days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        for idx, day in enumerate(days):
            self.routine_table.setItem(idx, 0, QTableWidgetItem(day))
            self.routine_table.item(idx, 0).setFlags(Qt.ItemFlag.ItemIsEnabled)
            
            # Place placeholders in subjects cells
            for col in range(1, 7):
                item = QTableWidgetItem("-")
                item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.routine_table.setItem(idx, col, item)
            
        self.routine_layout.addWidget(self.routine_table)
        
        # Auto-connect if there was a saved UID and Firebase is active
        if saved_uid and self.db:
            QTimer.singleShot(1000, self.start_routine_listener)

    def start_routine_listener(self):
        uid = self.routine_uid_input.toPlainText().strip()
        if not uid:
            QMessageBox.warning(self, "Required", "Please enter a valid User ID.")
            return
            
        if not self.db:
            QMessageBox.critical(self, "Error", "Firebase is not initialized. Check serviceAccountKey.json.")
            return

        # Save to local settings
        settings = QSettings("CopyPro", "RoutineConfig")
        settings.setValue("userId", uid)

        self.lbl_routine_status.setText("Connecting... 🟡")
        self.lbl_routine_status.setStyleSheet("font-weight: bold; color: #FF9800;")

        # Unsubscribe existing listener if any
        if hasattr(self, 'routine_listener') and self.routine_listener:
            try:
                self.routine_listener.unsubscribe()
            except: pass

        try:
            doc_ref = self.db.collection("user_routines").document(uid)
            
            def on_snapshot(doc_snapshot, changes, read_time):
                for doc in doc_snapshot:
                    if doc.exists:
                        data = doc.to_dict()
                        # Thread safety: emit the signal to update UI in main thread
                        self.routine_data_received.emit(data)
                        return
                # If document does not exist
                self.routine_data_received.emit({})

            self.routine_listener = doc_ref.on_snapshot(on_snapshot)
        except Exception as e:
            self.lbl_routine_status.setText("Failed 🔴")
            self.lbl_routine_status.setStyleSheet("font-weight: bold; color: #F44336;")
            QMessageBox.critical(self, "Error", f"Failed to start routine listener: {e}")

    def launch_routine_widget(self):
        uid = self.routine_uid_input.toPlainText().strip()
        if not uid:
            QMessageBox.warning(self, "Required", "Please enter a valid User ID first.")
            return
        if not self.db:
            QMessageBox.critical(self, "Error", "Firebase is not initialized.")
            return
            
        global global_floating_widgets
        for w in list(global_floating_widgets):
            try:
                w.close()
            except: pass
        global_floating_widgets.clear()
            
        widget = RoutineWidget(self.db, uid)
        global_floating_widgets.append(widget)
        widget.show()

    def get_week_key(self):
        import datetime
        now = datetime.datetime.now()
        jan1 = datetime.datetime(now.year, 1, 1)
        day_of_year = now.timetuple().tm_yday
        offset = (jan1.weekday() + 1) % 7
        week = (day_of_year + offset - 1) // 7 + 1
        return f"{week}-{now.year}"

    def update_routine_ui(self, data):
        if not data:
            self.lbl_routine_status.setText("Empty Routine ⚪")
            self.lbl_routine_status.setStyleSheet("font-weight: bold; color: #888;")
            # Reset table cells
            for r in range(7):
                self.routine_table.setItem(r, 0, QTableWidgetItem(["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"][r]))
                for col in range(1, 7):
                    self.routine_table.setCellWidget(r, col, None)
                    item = QTableWidgetItem("-")
                    item.setFlags(Qt.ItemFlag.ItemIsEnabled)
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                    self.routine_table.setItem(r, col, item)
            return

        self.lbl_routine_status.setText("Live Synced 🟢")
        self.lbl_routine_status.setStyleSheet("font-weight: bold; color: #4CAF50;")
        
        # Save raw data list to edit later if toggled
        self.current_routine_days = data.get("days", [])
        
        days_list = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        
        # Map raw data by dayName for fast lookup (Abbreviated keys in DB like 'MON')
        days_map = {}
        for d in self.current_routine_days:
            name = d.get("dayName", "").strip().upper()
            if name: days_map[name] = d

        abbrev_map = {
            "monday": "MON", "tuesday": "TUE", "wednesday": "WED", 
            "thursday": "THU", "friday": "FRI", "saturday": "SAT", "sunday": "SUN"
        }

        for r, day_name in enumerate(days_list):
            abbrev = abbrev_map.get(day_name.lower(), "MON")
            d_data = days_map.get(abbrev, {})
            
            # 1. Update Holiday Status (Styled directly in Day Name)
            is_holiday = d_data.get("isHoliday", False)
            day_text = f"{day_name[:3].upper()}" + (" 🌴" if is_holiday else "")
            
            item_day = QTableWidgetItem(day_text)
            item_day.setFlags(Qt.ItemFlag.ItemIsEnabled)
            item_day.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if is_holiday:
                item_day.setForeground(QColor("#FF5722"))
                item_day.setToolTip("Holiday")
            self.routine_table.setItem(r, 0, item_day)
            
            completed_list = d_data.get("completedSubjects", [])
            
            # 2. Render Day Slot Subjects (Cols 1, 2, 3)
            day_subjects = d_data.get("daySubjects", [])
            for i in range(3):
                sub_name = day_subjects[i] if i < len(day_subjects) else ""
                self.render_single_subject(r, 1 + i, "day", sub_name, completed_list, day_name)
            
            # 3. Render Night Slot Subjects (Cols 4, 5, 6)
            night_subjects = d_data.get("nightSubjects", [])
            for i in range(3):
                sub_name = night_subjects[i] if i < len(night_subjects) else ""
                self.render_single_subject(r, 4 + i, "night", sub_name, completed_list, day_name)

    def render_single_subject(self, row, col, slot_type, sub_name, completed_list, day_name):
        if not sub_name or sub_name == "-":
            self.routine_table.setCellWidget(row, col, None)
            item = QTableWidgetItem("-")
            item.setFlags(Qt.ItemFlag.ItemIsEnabled)
            item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            item.setForeground(QColor("#555555"))
            self.routine_table.setItem(row, col, item)
            return
            
        btn = QPushButton(sub_name)
        btn.setToolTip(sub_name) # Hover to read full text
        btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        
        # Calculate original index in slot list
        index = (col - 1) if slot_type == "day" else (col - 4)
        
        # Map day name to abbreviated MON, TUE, etc.
        abbrev_map = {
            "monday": "MON", "tuesday": "TUE", "wednesday": "WED", 
            "thursday": "THU", "friday": "FRI", "saturday": "SAT", "sunday": "SUN"
        }
        day_abbrev = abbrev_map.get(day_name.lower(), "MON")
        
        week_key = self.get_week_key()
        slot_key = f"{day_abbrev}|{'D' if slot_type == 'day' else 'N'}|{index}|{week_key}"
        
        is_completed = slot_key in completed_list
        
        # Style button depending on completion status
        if is_completed:
            btn.setStyleSheet("""
                QPushButton {
                    background-color: rgba(76, 175, 80, 200);
                    border: 1px solid #4CAF50;
                    border-radius: 6px;
                    color: white;
                    font-weight: bold;
                    font-size: 11px;
                    padding: 2px;
                }
                QPushButton:hover {
                    background-color: rgba(76, 175, 80, 240);
                }
            """)
            btn.setText(f"{sub_name} ✓")
        else:
            btn.setStyleSheet("""
                QPushButton {
                    background-color: rgba(255, 255, 255, 10);
                    border: 1px solid rgba(255, 255, 255, 20);
                    border-radius: 6px;
                    color: #DCDCDC;
                    font-size: 11px;
                    padding: 2px;
                }
                QPushButton:hover {
                    background-color: rgba(0, 120, 212, 120);
                    border: 1px solid #0078D4;
                }
            """)

        btn.clicked.connect(lambda checked, d=day_name, s=slot_type, idx=index: self.toggle_subject_status_by_key(d, s, idx))
        self.routine_table.setCellWidget(row, col, btn)

    def toggle_subject_status_by_key(self, day_name, slot_type, index):
        uid = self.routine_uid_input.toPlainText().strip()
        if not uid or not self.db or not hasattr(self, 'current_routine_days'): return

        abbrev_map = {
            "monday": "MON", "tuesday": "TUE", "wednesday": "WED", 
            "thursday": "THU", "friday": "FRI", "saturday": "SAT", "sunday": "SUN"
        }
        day_abbrev = abbrev_map.get(day_name.lower(), "MON")
        week_key = self.get_week_key()
        slot_key = f"{day_abbrev}|{'D' if slot_type == 'day' else 'N'}|{index}|{week_key}"

        try:
            # Modify target day's completedSubjects in self.current_routine_days
            updated_days = []
            for d in self.current_routine_days:
                d_copy = dict(d)
                if d_copy.get("dayName", "").strip().upper() == day_abbrev:
                    comp = list(d_copy.get("completedSubjects", []))
                    if slot_key in comp:
                        comp.remove(slot_key) # Uncheck
                    else:
                        comp.append(slot_key) # Check (Green)
                    d_copy["completedSubjects"] = comp
                updated_days.append(d_copy)

            # Write standard set back to Firestore
            doc_ref = self.db.collection("user_routines").document(uid)
            doc_ref.set({"days": updated_days}, merge=True)
            print(f"Subject index {index} status toggled for {day_name}!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to update task status: {e}")

    def sync_daily_to_cloud(self):
        if not self.db: return
        if self.daily_table.rowCount() == 0: return
        
        try:
            batch = self.db.batch()
            for r in range(self.daily_table.rowCount()):
                word = self.daily_table.item(r, 0).text().strip() if self.daily_table.item(r, 0) else ""
                meaning = self.daily_table.item(r, 1).text().strip() if self.daily_table.item(r, 1) else ""
                phrases = self.daily_table.item(r, 2).text().strip() if self.daily_table.item(r, 2) else ""
                story = self.daily_table.item(r, 3).text().strip() if self.daily_table.item(r, 3) else ""
                
                # Mandatory Field Check
                if not all([word, meaning, phrases, story]):
                    QMessageBox.warning(self, "Data Missing", f"Row {r+1} is missing mandatory fields (Word, Meaning, Phrases, or Story). Sync aborted.")
                    return
                
                quiz_raw = self.daily_table.item(r, 4).text().strip() if self.daily_table.item(r, 4) else "[]"
                try:
                    quiz_data = json.loads(quiz_raw) if (quiz_raw.startswith('[') or quiz_raw.startswith('{')) else []
                    if isinstance(quiz_data, dict): quiz_data = [quiz_data] # Convert single object to list
                except: quiz_data = []

                data = {
                    "word": word,
                    "meaning": meaning,
                    "phrases": phrases,
                    "story": story,
                    "quiz": quiz_data,
                    "timestamp": firestore.SERVER_TIMESTAMP
                }
                # Use word as document ID (MD5) to avoid duplicates
                doc_id = hashlib.md5(word.lower().encode()).hexdigest()
                batch.set(self.db.collection("daily_practice").document(doc_id), data, merge=True)
            
            batch.commit()
            QMessageBox.information(self, "Success", "Daily Practice synced! Data is now A-Z Smart Sorted in the app.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Sync failed: {e}")

    def open_bulk_daily_import(self):
        d = BulkImportDialog(self, title="Daily Practice Bulk Import", expected_format="Word | Meaning | Phrases / Examples | Story / Usage | Quiz_JSON", min_cols=1)
        if d.exec() == QDialog.DialogCode.Accepted:
            rows = d.get_rows()
            if not rows:
                QMessageBox.warning(self, "No Data", "No valid rows found to import.")
                return
                
            imported_count = 0
            for parts in rows:
                if not any(parts): continue
                r = self.daily_table.rowCount()
                self.daily_table.insertRow(r)
                
                for i in range(7):
                    self.daily_table.setItem(r, i, QTableWidgetItem(""))
                    
                for i in range(min(len(parts), 5)):
                    self.daily_table.setItem(r, i, QTableWidgetItem(str(parts[i]).strip()))
                
                self.daily_table.setItem(r, 5, QTableWidgetItem(str(r + 1))) # Order
                del_btn = QPushButton("🗑️")
                del_btn.clicked.connect(lambda _, btn=del_btn: self.daily_table.removeRow(self.daily_table.indexAt(btn.pos()).row() if self.daily_table.indexAt(btn.pos()).row() >= 0 else self.daily_table.currentRow()))
                self.daily_table.setCellWidget(r, 6, del_btn)
                imported_count += 1
                
            self.daily_table.viewport().update()
            QMessageBox.information(self, "Success", f"✅ Successfully imported {imported_count} vocabulary items!")

    def add_mock_row(self):
        r = self.mock_table.rowCount(); self.mock_table.insertRow(r)
        for i in range(8): self.mock_table.setItem(r, i, QTableWidgetItem(""))
        
        # Category Dropdown (at index 8)
        cat_combo = QComboBox()
        cat_combo.addItems(["Reasoning", "GK", "Mathematics", "English"])
        cat_colors = {"Reasoning": "#9C27B0", "GK": "#4CAF50", "Mathematics": "#2196F3", "English": "#E91E63"}
        cat_combo.setStyleSheet(f"background-color: {cat_colors['Reasoning']}; font-weight: bold;")
        cat_combo.currentTextChanged.connect(lambda t, c=cat_combo: c.setStyleSheet(f"background-color: {cat_colors.get(t, '#444')}; font-weight: bold;"))
        self.mock_table.setCellWidget(r, 8, cat_combo)
        
        del_btn = QPushButton("🗑️")
        del_btn.clicked.connect(lambda _, row=r: self.mock_table.removeRow(self.mock_table.currentRow() if self.mock_table.currentRow() >=0 else row))
        self.mock_table.setCellWidget(r, 9, del_btn)



    def sync_mock_to_cloud(self):
        topic = self.mock_test_name.toPlainText().strip()
        if not topic:
            QMessageBox.warning(self, "Required", "Please enter a Mock Test Name.")
            return
        if self.mock_table.rowCount() == 0: return
        
        try:
            batch = self.db.batch()
            ts = firestore.SERVER_TIMESTAMP
            for r in range(self.mock_table.rowCount()):
                # Get image data from index 7 (previously 8)
                img_item = self.mock_table.item(r, 7)
                img_data = img_item.data(Qt.ItemDataRole.UserRole) if img_item else ""
                if not img_data and img_item: img_data = img_item.text()

                cat = self.mock_table.cellWidget(r, 8).currentText()
                data = {
                    "question": self.mock_table.item(r, 0).text() if self.mock_table.item(r, 0) else "",
                    "optionA": self.mock_table.item(r, 1).text() if self.mock_table.item(r, 1) else "",
                    "optionB": self.mock_table.item(r, 2).text() if self.mock_table.item(r, 2) else "",
                    "optionC": self.mock_table.item(r, 3).text() if self.mock_table.item(r, 3) else "",
                    "optionD": self.mock_table.item(r, 4).text() if self.mock_table.item(r, 4) else "",
                    "answer": self.mock_table.item(r, 5).text() if self.mock_table.item(r, 5) else "",
                    "solution": self.mock_table.item(r, 6).text() if self.mock_table.item(r, 6) else "",
                    "questionImage": img_data,
                    "type": cat, "subject": "Mock Test 🏆", "topic": topic, "timestamp": ts
                }
                batch.set(self.db.collection("quizzes").document(), data)
            batch.commit()
            QMessageBox.information(self, "Success", f"Mock Test '{topic}' synced to cloud!")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def open_bulk_mock_import(self):
        d = BulkImportDialog(self, title="Mock Creator Bulk Import", expected_format="Question | Option A | Option B | Option C | Option D | Answer | Solution | Category", min_cols=1)
        if d.exec() == QDialog.DialogCode.Accepted:
            rows = d.get_rows()
            if not rows:
                QMessageBox.warning(self, "No Data", "No valid rows found to import.")
                return
                
            imported_count = 0
            cat_colors = {"Reasoning": "#9C27B0", "GK": "#4CAF50", "Mathematics": "#2196F3", "English": "#E91E63"}
            
            for parts in rows:
                if not any(parts): continue
                r = self.mock_table.rowCount()
                self.mock_table.insertRow(r)
                
                for i in range(8):
                    self.mock_table.setItem(r, i, QTableWidgetItem(""))
                
                # Map parts[0-6] to Question, A, B, C, D, Ans, Solution
                for i in range(min(len(parts), 7)):
                    self.mock_table.setItem(r, i, QTableWidgetItem(str(parts[i]).strip()))
                
                # Format answer letter if present
                ans_text = self.mock_table.item(r, 5).text().strip()
                m = re.search(r'([A-Da-d])', ans_text)
                if m and len(ans_text) > 1:
                    self.mock_table.setItem(r, 5, QTableWidgetItem(m.group(1).upper()))
                
                # Category Dropdown (at index 8)
                cat_combo = QComboBox()
                cat_combo.addItems(["Reasoning", "GK", "Mathematics", "English"])
                target_cat = parts[7].strip() if len(parts) > 7 and parts[7].strip() in cat_colors else "Reasoning"
                cat_combo.setCurrentText(target_cat)
                cat_combo.setStyleSheet(f"background-color: {cat_colors.get(target_cat, '#444')}; font-weight: bold;")
                cat_combo.currentTextChanged.connect(lambda t, c=cat_combo: c.setStyleSheet(f"background-color: {cat_colors.get(t, '#444')}; font-weight: bold;"))
                self.mock_table.setCellWidget(r, 8, cat_combo)
                
                # Action button (index 9)
                del_btn = QPushButton("🗑️")
                del_btn.clicked.connect(lambda _, btn=del_btn: self.mock_table.removeRow(self.mock_table.indexAt(btn.pos()).row() if self.mock_table.indexAt(btn.pos()).row() >= 0 else self.mock_table.currentRow()))
                self.mock_table.setCellWidget(r, 9, del_btn)
                imported_count += 1
                
            self.mock_table.viewport().update()
            QMessageBox.information(self, "Success", f"✅ Successfully imported {imported_count} questions into Mock Creator!")

    def transfer_to_mock(self):
        r = self.table.currentRow()
        if r < 0: r = self.table.rowCount() - 1
        if r < 0: return
        
        # Get data from Extractor table
        q = self.table.item(r, 0).text() if self.table.item(r, 0) else ""
        a = self.table.item(r, 1).text() if self.table.item(r, 1) else ""
        b = self.table.item(r, 2).text() if self.table.item(r, 2) else ""
        c = self.table.item(r, 3).text() if self.table.item(r, 3) else ""
        d = self.table.item(r, 4).text() if self.table.item(r, 4) else ""
        ans = self.table.item(r, 5).text() if self.table.item(r, 5) else ""
        sol = self.table.item(r, 6).text() if self.table.item(r, 6) else ""
        
        # Switch to Mock Tab and Add
        self.tabs.setCurrentIndex(1)
        self.add_mock_row()
        mr = self.mock_table.rowCount() - 1
        
        self.mock_table.setItem(mr, 0, QTableWidgetItem(q))
        self.mock_table.setItem(mr, 1, QTableWidgetItem(a))
        self.mock_table.setItem(mr, 2, QTableWidgetItem(b))
        self.mock_table.setItem(mr, 3, QTableWidgetItem(c))
        self.mock_table.setItem(mr, 4, QTableWidgetItem(d))
        self.mock_table.setItem(mr, 5, QTableWidgetItem(ans))
        self.mock_table.setItem(mr, 6, QTableWidgetItem(sol))
        
        QMessageBox.information(self, "Success", "Question transferred to Mock Creator!")

    def update_subject_theme(self):
        subject = self.combo_subject.currentText()
        is_english = "English" in subject
        accent_color = "rgba(233, 30, 99, 180)" if is_english else "rgba(0, 120, 212, 150)"
        hover_color = "rgba(233, 30, 99, 220)" if is_english else "rgba(0, 120, 212, 180)"
        
        self.setStyleSheet(self.styleSheet() + f"""
            QPushButton#accent {{ background-color: {accent_color}; }}
            QPushButton#accent:hover {{ background-color: {hover_color}; }}
        """)

    def auto_adjust_table_height(self):
        try:
            if not self.btn_mini.isChecked():
                self.table.setMinimumHeight(250)
                self.table.setMaximumHeight(16777215)
                self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
                self.table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
            else:
                hh = self.table.horizontalHeader().height()
                rh = sum(self.table.rowHeight(i) for i in range(self.table.rowCount()))
                th = hh + rh + 2
                mh = 300
                fh = min(th, mh)
                self.table.setMinimumHeight(fh)
                self.table.setMaximumHeight(fh)
                self.table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded if th > mh else Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
                self.resize(self.width(), 32 + fh)
        except: pass

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            event.accept()
    def mouseMoveEvent(self, event):
        if event.buttons() == Qt.MouseButton.LeftButton:
            self.move(event.globalPosition().toPoint() - self.drag_pos)
            event.accept()

    def toggle_mini_mode(self, checked):
        try:
            if checked:
                self.setWindowFlags(Qt.WindowType.WindowStaysOnTopHint | Qt.WindowType.FramelessWindowHint)
                self.editor_label.hide(); self.detail_editor.hide()
                self.main_layout.setContentsMargins(0,0,0,0); self.main_layout.setSpacing(0)
                self.main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
                self.table_container.layout().setContentsMargins(0,0,0,0)
                self.toolbar_layout.setContentsMargins(4,2,4,2); self.toolbar_widget.setFixedHeight(32)
                self.table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
                self.table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
                self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
                self.table.horizontalHeader().setFixedHeight(20); self.table.verticalHeader().hide()
                self.auto_adjust_table_height()
                self.table.setStyleSheet("QTableWidget { border: none; background-color: rgba(30,30,30,180); }")
            else:
                self.setWindowFlags(Qt.WindowType.Window)
                self.main_layout.setAlignment(Qt.AlignmentFlag(0))
                self.editor_label.show(); self.detail_editor.show()
                self.main_layout.setContentsMargins(15,15,15,15); self.main_layout.setSpacing(10)
                self.table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
                self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
                self.table.setMinimumHeight(0); self.table.setMaximumHeight(16777215); self.table.setFixedHeight(16777215)
                self.table.verticalHeader().show(); self.table.horizontalHeader().show()
                self.toolbar_widget.setFixedHeight(16777215); self.resize(1150, 650)
                self.auto_adjust_table_height()
            self.show(); self.raise_()
        except: pass

    def sync_to_firebase(self):
        if not self.db: return
        if self.table.rowCount() == 0: return
        
        # Strip emoji for Firestore compatibility
        clean_subject = self.combo_subject.currentText().split(' ')[0]
        d = SyncMetadataDialog(clean_subject, self)
        if d.exec() == QDialog.DialogCode.Accepted:
            meta = d.get_data()
            # Ensure subject is clean even if user typed in dialog
            meta_subject = meta["subject"].split(' ')[0]
            try:
                batch = self.db.batch()
                ts = firestore.SERVER_TIMESTAMP
                for r in range(self.table.rowCount()):
                    q_text = self.table.item(r, 0).text().strip() if self.table.item(r, 0) else ""
                    if not q_text: continue
                    
                    row_id = self.table.item(r, 10).text() if self.table.item(r, 10) else ""

                    row_type = self.table.item(r, 7).text().strip() if self.table.item(r, 7) else ""
                    final_type = row_type if row_type else meta["type"]

                    data = {
                        "question": q_text,
                        "optionA": self.table.item(r, 1).text() if self.table.item(r, 1) else "",
                        "optionB": self.table.item(r, 2).text() if self.table.item(r, 2) else "",
                        "optionC": self.table.item(r, 3).text() if self.table.item(r, 3) else "",
                        "optionD": self.table.item(r, 4).text() if self.table.item(r, 4) else "",
                        "answer": self.table.item(r, 5).text() if self.table.item(r, 5) else "",
                        "solution": self.table.item(r, 6).text() if self.table.item(r, 6) else "",
                        "questionImage": self.table.item(r, 8).data(Qt.ItemDataRole.UserRole) or self.table.item(r, 8).text() if self.table.item(r, 8) else "",
                        "type": final_type, "subject": meta_subject, "topic": meta["topic"], 
                        "test_code": meta["test_code"], "is_complete": meta["is_complete"], "timestamp": ts
                    }
                    
                    if row_id:
                        # 🔄 Existing Document -> Update
                        batch.set(self.db.collection("quizzes").document(row_id), data, merge=True)
                    else:
                        # ✨ New Document -> Create with deterministic ID to prevent instant double-sync duplicates
                        content_id = hashlib.md5(f"{meta_subject}_{meta['topic']}_{q_text}".encode()).hexdigest()
                        batch.set(self.db.collection("quizzes").document(content_id), data, merge=True)
                batch.commit()
                QMessageBox.information(self, "Success", "Synced!")
            except: pass

    def open_history(self):
        if not self.db: return
        CloudHistoryDialog(self.db, self).exec()

    def load_from_cloud(self, topic, subject):
        try:
            docs = list(self.db.collection("quizzes").where("topic", "==", topic).where("subject", "==", subject).stream())
            if not docs:
                QMessageBox.warning(self, "Not Found", f"No questions found for '{topic}' under '{subject}'.")
                return

            current_tab = self.tabs.currentIndex()
            if current_tab == 3: # 🎮 Laptop Quiz Tab
                self.quiz_play_table.setRowCount(0)
                for doc in docs:
                    d = doc.to_dict()
                    r = self.quiz_play_table.rowCount(); self.quiz_play_table.insertRow(r)
                    cols = ["question", "optionA", "optionB", "optionC", "optionD", "answer", "solution"]
                    for i, key in enumerate(cols):
                        self.quiz_play_table.setItem(r, i, QTableWidgetItem(str(d.get(key, ""))))
                    
                    img_data = d.get("questionImage", "")
                    if img_data:
                        item = QTableWidgetItem("🖼️ Ready")
                        item.setData(Qt.ItemDataRole.UserRole, str(img_data))
                        self.quiz_play_table.setItem(r, 7, item)
                    else:
                        self.quiz_play_table.setItem(r, 7, QTableWidgetItem(""))
                    
                    del_btn = QPushButton("🗑️")
                    del_btn.clicked.connect(lambda _, btn=del_btn: self.quiz_play_table.removeRow(self.quiz_play_table.indexAt(btn.pos()).row() if self.quiz_play_table.indexAt(btn.pos()).row() >= 0 else self.quiz_play_table.currentRow()))
                    self.quiz_play_table.setCellWidget(r, 8, del_btn)
                self.quiz_play_table.viewport().update()
                QMessageBox.information(self, "Cloud Loaded", f"✅ Loaded {len(docs)} questions of '{topic}' into Laptop Quiz!")
            elif subject == "Mock Test 🏆" or current_tab == 1:
                self.tabs.setCurrentIndex(1)
                self.mock_table.setRowCount(0)
                self.mock_test_name.setPlainText(topic)
                for doc in docs:
                    d = doc.to_dict()
                    r = self.mock_table.rowCount(); self.mock_table.insertRow(r)
                    
                    cat_combo = QComboBox()
                    cat_combo.addItems(["Reasoning", "GK", "Mathematics", "English"])
                    target_cat = d.get("type", "Reasoning")
                    cat_colors = {"Reasoning": "#9C27B0", "GK": "#4CAF50", "Mathematics": "#2196F3", "English": "#E91E63"}
                    cat_combo.setCurrentText(target_cat)
                    cat_combo.setStyleSheet(f"background-color: {cat_colors.get(target_cat, '#444')}; font-weight: bold;")
                    
                    self.mock_table.setCellWidget(r, 8, cat_combo)
                    cols = ["question", "optionA", "optionB", "optionC", "optionD", "answer", "solution"]
                    for i, key in enumerate(cols):
                        self.mock_table.setItem(r, i, QTableWidgetItem(str(d.get(key, ""))))
                    
                    del_btn = QPushButton("🗑️")
                    del_btn.clicked.connect(lambda _, btn=del_btn: self.mock_table.removeRow(self.mock_table.indexAt(btn.pos()).row() if self.mock_table.indexAt(btn.pos()).row() >= 0 else self.mock_table.currentRow()))
                    self.mock_table.setCellWidget(r, 9, del_btn)
                self.mock_table.viewport().update()
                QMessageBox.information(self, "Loaded", f"✅ Loaded {len(docs)} questions into Mock Creator.")
            else:
                self.tabs.setCurrentIndex(0)
                self.table.setRowCount(0)
                for doc in docs:
                    d = doc.to_dict()
                    row = self.table.rowCount(); self.table.insertRow(row)
                    cols = ["question", "optionA", "optionB", "optionC", "optionD", "answer", "solution", "type"]
                    for i, key in enumerate(cols):
                        self.table.setItem(row, i, QTableWidgetItem(str(d.get(key, ""))))
                    self.table.setItem(row, 10, QTableWidgetItem(doc.id))
                    self.add_delete_button(row)
                self.auto_adjust_table_height()
                self.table.viewport().update()
                QMessageBox.information(self, "Loaded", f"✅ Loaded {len(docs)} questions of '{topic}' into Extractor.")
        except Exception as e:
            QMessageBox.critical(self, "Load Error", f"Failed to load from cloud: {e}")

    def open_bulk_import(self):
        d = BulkImportDialog(self, title="MCQ Extractor Bulk Import", expected_format="Question | Option A | Option B | Option C | Option D | Answer | Solution | Type", min_cols=1)
        if d.exec() == QDialog.DialogCode.Accepted:
            rows = d.get_rows()
            if not rows:
                QMessageBox.warning(self, "No Data", "No valid rows found to import.")
                return
            
            imported_count = 0
            for parts in rows:
                if not any(parts): continue
                r = self.table.rowCount()
                self.table.insertRow(r)
                
                # Initialize all 11 columns
                for i in range(11):
                    self.table.setItem(r, i, QTableWidgetItem(""))
                
                # Map parts: Q (0), A (1), B (2), C (3), D (4), Ans (5), Sol (6), Type (7)
                for i in range(min(len(parts), 8)):
                    self.table.setItem(r, i, QTableWidgetItem(str(parts[i]).strip()))
                
                # Default type if missing
                if len(parts) <= 7 or not self.table.item(r, 7).text():
                    self.table.setItem(r, 7, QTableWidgetItem("MCQ"))
                
                # Format answer letter if present
                ans_text = self.table.item(r, 5).text().strip()
                m = re.search(r'([A-Da-d])', ans_text)
                if m and len(ans_text) > 1:
                    self.table.setItem(r, 5, QTableWidgetItem(m.group(1).upper()))
                
                self.table.setItem(r, 10, QTableWidgetItem("")) # doc_id
                self.add_delete_button(r)
                imported_count += 1
                
            self.auto_adjust_table_height()
            self.table.viewport().update()
            QMessageBox.information(self, "Success", f"✅ Successfully imported {imported_count} questions into Extractor!")

    def export_to_simple_q(self):
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "No Data", "Table is empty. Please add or import questions first.")
            return
        default_path = os.path.join(os.path.expanduser("~"), "Downloads", "quizzes.xlsx")
        f, _ = QFileDialog.getSaveFileName(self, "Export to Simple Q", default_path, "Excel (*.xlsx)")
        if f:
            if not f.endswith('.xlsx'): f += '.xlsx'
            data = []
            for r in range(self.table.rowCount()):
                data.append([self.table.item(r, c).text().strip() if self.table.item(r, c) else "" for c in range(8)])
            try:
                pd.DataFrame(data, columns=["Question", "A", "B", "C", "D", "Ans", "Sol", "Type"]).to_excel(f, index=False)
                QMessageBox.information(self, "Export Successful", f"✅ Successfully exported {len(data)} questions to:\n{f}")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Failed to save file: {e}")

    def sync_table_to_editor(self, r, c, pr, pc):
        it = self.table.item(r, c)
        self.detail_editor.blockSignals(True); self.detail_editor.setPlainText(it.text() if it else ""); self.detail_editor.blockSignals(False)
    def sync_editor_to_table(self):
        r = self.table.currentRow(); c = self.table.currentColumn()
        if r >= 0 and c >= 0: self.table.setItem(r, c, QTableWidgetItem(self.detail_editor.toPlainText()))

    def show_context_menu(self, pos):
        menu = QMenu(); delete_action = menu.addAction("❌ Delete Row")
        delete_action.triggered.connect(self.delete_selected_row)
        menu.exec(self.table.viewport().mapToGlobal(pos))
    def delete_selected_row(self):
        r = self.table.currentRow()
        if r >= 0: self.table.removeRow(r); self.auto_adjust_table_height()

    def quick_set_answer(self, letter):
        current_tab = self.tabs.currentIndex()
        if current_tab == 0: # Extractor
            r = self.table.currentRow()
            if r < 0: r = self.table.rowCount() - 1
            if r >= 0: self.table.setItem(r, 5, QTableWidgetItem(letter))
        else: # Mock Creator
            r = self.mock_table.currentRow()
            if r < 0: r = self.mock_table.rowCount() - 1
            if r >= 0: self.mock_table.setItem(r, 6, QTableWidgetItem(letter))

    def add_blank_row(self):
        r = self.table.rowCount(); self.table.insertRow(r)
        for i in range(10): self.table.setItem(r, i, QTableWidgetItem(""))
        self.table.setItem(r, 7, QTableWidgetItem("MCQ"))
        self.add_delete_button(r); self.auto_adjust_table_height()

    def add_delete_button(self, r):
        btn = QPushButton("🗑️")
        btn.setStyleSheet("background-color: rgba(211,47,47,200); color: white;")
        btn.clicked.connect(lambda _, row=r: self.delete_row_by_index(row))
        self.table.setCellWidget(r, 9, btn)
    def delete_row_by_index(self, r):
        self.table.removeRow(self.table.currentRow() if self.table.currentRow() >=0 else r); self.auto_adjust_table_height()

    def upload_question_image(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select Question Image", "", "Images (*.png *.jpg *.jpeg)")
        if file_path:
            self.process_image_to_base64(file_path)

    def start_capture(self, target):
        current_tab = self.tabs.currentIndex()
        if current_tab == 0:
            if self.table.rowCount() == 0: self.add_blank_row()
        elif current_tab == 1:
            if self.mock_table.rowCount() == 0: self.add_mock_row()
        elif current_tab == 2:
            if self.daily_table.rowCount() == 0: self.add_daily_row()
        elif current_tab == 3:
            if self.quiz_play_table.rowCount() == 0: self.add_quiz_play_row()
            
        self.current_target = target; self.overlay = Overlay()
        self.overlay.image_captured.connect(self.process_image); self.overlay.show()

    def process_image(self, ip):
        if self.current_target in ["MockImage", "Image", "QuizImage"]:
            self.process_image_to_base64(ip)
        else:
            if not self.ocr_engine:
                QMessageBox.warning(self, "OCR Engine", "Windows OCR Engine is not initialized.")
                return
            t = OCRThread(ip, self.ocr_engine, self.parser, self.current_target)
            t.result_ready.connect(self.update_row)
            t.error_occurred.connect(lambda err: print(f"OCR Thread Error: {err}"))
            if not hasattr(self, '_active_threads'):
                self._active_threads = []
            self._active_threads.append(t)
            t.finished.connect(lambda: self._active_threads.remove(t) if t in self._active_threads else None)
            t.start()

    def process_image_to_base64(self, ip):
        try:
            img = PILImage.open(ip)
            if img.width > 800:
                h = int(img.height * (800 / img.width))
                img = img.resize((800, h), PILImage.Resampling.LANCZOS)
            output = io.BytesIO()
            img.save(output, format="JPEG", quality=70)
            base64_str = base64.b64encode(output.getvalue()).decode('utf-8')
            # If we uploaded, we might need a target
            if not hasattr(self, 'current_target') or self.tabs.currentIndex() in [1, 3]:
                if self.tabs.currentIndex() == 1:
                    self.current_target = "MockImage"
                elif self.tabs.currentIndex() == 3:
                    self.current_target = "QuizImage"
                else:
                    self.current_target = "Image"
            self.update_row(base64_str)
            return base64_str
        except Exception as e:
            print(f"Image processing error: {e}")
            return ""

    def update_row(self, data):
        current_tab = self.tabs.currentIndex()
        if current_tab == 0: # MCQ Extractor Tab
            r = self.table.currentRow()
            if r < 0: r = self.table.rowCount() - 1
            if self.current_target == "Question": self.table.setItem(r, 0, QTableWidgetItem(str(data)))
            elif self.current_target == "Options" and isinstance(data, dict):
                for i, k in enumerate(["Option A", "Option B", "Option C", "Option D"]):
                    self.table.setItem(r, i+1, QTableWidgetItem(str(data.get(k, ""))))
            elif self.current_target == "Answer":
                m = re.search(r'([A-Da-d])', str(data))
                self.table.setItem(r, 5, QTableWidgetItem(m.group(1).upper() if m else str(data)))
            elif self.current_target == "Solution": self.table.setItem(r, 6, QTableWidgetItem(str(data)))
            elif self.current_target == "Image": 
                item = QTableWidgetItem("🖼️ Ready")
                item.setData(Qt.ItemDataRole.UserRole, str(data)) # 🔒 Store real data here
                item.setToolTip("Image Captured Successfully")
                self.table.setItem(r, 8, item)
            elif self.current_target == "Type": self.table.setItem(r, 7, QTableWidgetItem(str(data)))
        elif current_tab == 1: # Mock Creator Tab
            r = self.mock_table.currentRow()
            if r < 0: r = self.mock_table.rowCount() - 1
            if self.current_target == "Question": self.mock_table.setItem(r, 0, QTableWidgetItem(str(data)))
            elif self.current_target == "Options" and isinstance(data, dict):
                for i, k in enumerate(["Option A", "Option B", "Option C", "Option D"]):
                    self.mock_table.setItem(r, i+1, QTableWidgetItem(str(data.get(k, ""))))
            elif self.current_target == "Answer":
                m = re.search(r'([A-Da-d])', str(data))
                self.mock_table.setItem(r, 5, QTableWidgetItem(m.group(1).upper() if m else str(data)))
            elif self.current_target == "Solution": self.mock_table.setItem(r, 6, QTableWidgetItem(str(data)))
            elif self.current_target == "MockImage": 
                item = QTableWidgetItem("🖼️ Ready")
                item.setData(Qt.ItemDataRole.UserRole, str(data))
                item.setToolTip("Image Captured Successfully")
                self.mock_table.setItem(r, 7, item)
        elif current_tab == 2: # Daily Practice Tab
            r = self.daily_table.currentRow()
            if r < 0: r = self.daily_table.rowCount() - 1
            if r < 0: return
            
            if self.current_target == "Word": self.daily_table.setItem(r, 0, QTableWidgetItem(str(data)))
            elif self.current_target == "Meaning": self.daily_table.setItem(r, 1, QTableWidgetItem(str(data)))
            elif self.current_target == "Phrases": self.daily_table.setItem(r, 2, QTableWidgetItem(str(data)))
            elif self.current_target == "Story": self.daily_table.setItem(r, 3, QTableWidgetItem(str(data)))
        elif current_tab == 3: # Laptop Quiz Tab
            r = self.quiz_play_table.currentRow()
            if r < 0: r = self.quiz_play_table.rowCount() - 1
            if r < 0: return
            
            if self.current_target == "QuizQuestion": self.quiz_play_table.setItem(r, 0, QTableWidgetItem(str(data)))
            elif self.current_target == "QuizOptions" and isinstance(data, dict):
                for i, k in enumerate(["Option A", "Option B", "Option C", "Option D"]):
                    self.quiz_play_table.setItem(r, i+1, QTableWidgetItem(str(data.get(k, ""))))
            elif self.current_target == "QuizOptionA": self.quiz_play_table.setItem(r, 1, QTableWidgetItem(str(data)))
            elif self.current_target == "QuizOptionB": self.quiz_play_table.setItem(r, 2, QTableWidgetItem(str(data)))
            elif self.current_target == "QuizOptionC": self.quiz_play_table.setItem(r, 3, QTableWidgetItem(str(data)))
            elif self.current_target == "QuizOptionD": self.quiz_play_table.setItem(r, 4, QTableWidgetItem(str(data)))
            elif self.current_target == "QuizAnswer":
                m = re.search(r'([A-Da-d])', str(data))
                self.quiz_play_table.setItem(r, 5, QTableWidgetItem(m.group(1).upper() if m else str(data)))
            elif self.current_target == "QuizSolution": self.quiz_play_table.setItem(r, 6, QTableWidgetItem(str(data)))
            elif self.current_target == "QuizImage": 
                item = QTableWidgetItem("🖼️ Ready")
                item.setData(Qt.ItemDataRole.UserRole, str(data))
                item.setToolTip("Image Captured Successfully")
                self.quiz_play_table.setItem(r, 7, item)

    def export_excel(self):
        if self.table.rowCount() == 0:
            QMessageBox.warning(self, "No Data", "Table is empty. Please add or import questions before exporting.")
            return
        default_path = os.path.join(os.path.expanduser("~"), "Downloads", "MCQ_Questions.xlsx")
        f, _ = QFileDialog.getSaveFileName(self, "Save Excel File", default_path, "Excel (*.xlsx)")
        if f:
            if not f.endswith('.xlsx'): f += '.xlsx'
            d = []
            for r in range(self.table.rowCount()):
                row_vals = [self.table.item(r, c).text().strip() if self.table.item(r, c) else "" for c in range(8)]
                d.append(row_vals)
            try:
                pd.DataFrame(d, columns=["Question", "Option A", "Option B", "Option C", "Option D", "Correct Answer", "Solution", "Type"]).to_excel(f, index=False)
                QMessageBox.information(self, "Export Successful", f"✅ Successfully exported {len(d)} rows to:\n{f}")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Failed to save Excel file: {e}")

    def export_mock_excel(self):
        if self.mock_table.rowCount() == 0:
            QMessageBox.warning(self, "No Data", "Mock test table is empty. Please add or import questions first.")
            return
        default_path = os.path.join(os.path.expanduser("~"), "Downloads", "Mock_Test.xlsx")
        f, _ = QFileDialog.getSaveFileName(self, "Save Mock Test to Excel", default_path, "Excel (*.xlsx)")
        if f:
            if not f.endswith('.xlsx'): f += '.xlsx'
            d = []
            for r in range(self.mock_table.rowCount()):
                row_vals = []
                for c in range(7): # Columns 0 to 6: Question, A, B, C, D, Ans, Sol
                    it = self.mock_table.item(r, c)
                    row_vals.append(it.text().strip() if it else "")
                cat_widget = self.mock_table.cellWidget(r, 8)
                cat_val = cat_widget.currentText() if cat_widget and hasattr(cat_widget, 'currentText') else "Reasoning"
                row_vals.append(cat_val)
                d.append(row_vals)
            try:
                df = pd.DataFrame(d, columns=["Question", "Option A", "Option B", "Option C", "Option D", "Correct Answer", "Solution", "Category"])
                df.to_excel(f, index=False)
                QMessageBox.information(self, "Export Successful", f"✅ Successfully exported {len(d)} mock questions to:\n{f}")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", f"Failed to save Mock Excel: {e}")

    def closeEvent(self, event):
        # Quit application if no other windows are visible
        visible_windows = [w for w in QApplication.topLevelWidgets() if w.isVisible() and w != self]
        if not visible_windows:
            QApplication.quit()
        event.accept()

def global_exception_handler(exctype, value, tb):
    import traceback
    err_msg = "".join(traceback.format_exception(exctype, value, tb))
    print(f"CRASH CAUGHT SAFELY:\n{err_msg}")
    try:
        with open("error.log", "a", encoding="utf-8") as f:
            f.write(f"\n--- ERROR AT {pd.Timestamp.now()} ---\n{err_msg}\n")
    except: pass
    try:
        QMessageBox.critical(None, "Application Error", f"An unexpected error occurred:\n{value}\n\nDetails saved to error.log.")
    except: pass

if __name__ == "__main__":
    sys.excepthook = global_exception_handler
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
