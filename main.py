import sys
import os
import subprocess

from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit, QLabel,
    QHBoxLayout, QSizePolicy, QMessageBox
)
from PyQt6.QtGui import QIcon, QFont, QColor, QTextCharFormat, QTextCursor
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtWidgets import QGraphicsDropShadowEffect

LOG_FILE = "installer.log"
PS1_FILE = "installer.ps1"


# ===================== صلاحيات الأدمن ===================== #

def is_admin():
    """يتأكد إذا البرنامج شغال كـ Administrator على ويندوز"""
    try:
        import ctypes
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


# لو مش أدمن، يعيد تشغيل نفسه كأدمن
if os.name == "nt" and not is_admin():
    import ctypes
    params = " ".join([f'"{arg}"' for arg in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, params, None, 1
    )
    sys.exit(0)


# ===================== ثريد تشغيل الـ PowerShell ===================== #

class InstallerThread(QThread):
    output_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(int)

    def __init__(self, ps1_path, parent=None):
        super().__init__(parent)
        self.ps1_path = ps1_path

    def run(self):
        if not os.path.exists(self.ps1_path):
            self.output_signal.emit(f"Error: {os.path.basename(self.ps1_path)} not found!")
            self.finished_signal.emit(-1)
            return

        cmd = [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", self.ps1_path
        ]

        try:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                shell=False
            )
        except Exception as e:
            self.output_signal.emit(f"Failed to start installer: {e}")
            self.finished_signal.emit(-1)
            return

        # نكتب اللوج في ملف ونطلعه على الواجهة في نفس الوقت
        with open(LOG_FILE, "w", encoding="utf-8") as log_file:
            if process.stdout is not None:
                for line in process.stdout:
                    line = line.rstrip("\n")
                    self.output_signal.emit(line)
                    log_file.write(line + "\n")

        process.wait()
        self.finished_signal.emit(process.returncode)


# ===================== دوال تحليل الـ PS1 ===================== #

def parse_installer_script(ps1_path: str):
    """
    يقرأ ملف installer.ps1 ويطلّع ملخص بالباكدجات اللي هتتنزل:
    - winget
    - choco
    - pip
    - أوامر أخرى
    بيتم التحليل بشكل بسيط عن طريق البحث في السطور.
    """
    result = {
        "winget": [],
        "choco": [],
        "pip": [],
        "other": []
    }

    if not os.path.exists(ps1_path):
        return result, False  # الملف مش موجود

    try:
        with open(ps1_path, "r", encoding="utf-8") as f:
            lines = f.readlines()
    except UnicodeDecodeError:
        with open(ps1_path, "r", encoding="cp1252", errors="ignore") as f:
            lines = f.readlines()

    for raw_line in lines:
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue  # تجاهل السطور الفاضية أو الكومنت

        lower = line.lower()

        # اكتشاف أوامر winget
        if "winget" in lower and "install" in lower:
            result["winget"].append(line)
        # اكتشاف أوامر choco
        elif ("choco " in lower or "choco.exe" in lower) and "install" in lower:
            result["choco"].append(line)
        # اكتشاف أوامر pip
        elif ("pip " in lower or "pip3 " in lower or "python -m pip" in lower) and "install" in lower:
            result["pip"].append(line)
        else:
            # أوامر تانية ممكن تكون مهمة (زي إعدادات، تمكين ميزات، ..إلخ)
            # لو حابب تضيف فلترة أكتر هنا
            result["other"].append(line)

    return result, True


def build_plan_text(parsed):
    """يبني نص لطيف يتعرض في مربع "الخطة" من نتيجة parse_installer_script."""
    sections = []

    def add_section(title, items):
        if not items:
            return
        text = f"{title} ({len(items)}):\n"
        for i, line in enumerate(items, 1):
            text += f"  {i}. {line}\n"
        sections.append(text)

    add_section("Winget Packages", parsed.get("winget", []))
    add_section("Chocolatey Packages", parsed.get("choco", []))
    add_section("Pip / Python Packages", parsed.get("pip", []))
    add_section("Other Commands", parsed.get("other", []))

    if not sections:
        return "لم يتم العثور على أوامر تثبيت في installer.ps1.\nتأكد أن السكربت يحتوي على أوامر winget / choco / pip."

    return "\n".join(sections)


# ===================== الواجهة الرئيسية ===================== #

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.ps1_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), PS1_FILE)

        self.setWindowTitle("Windows Dev Tools Installer - WDT")
        self.setMinimumSize(850, 550)
        if os.path.exists("icon.png"):
            self.setWindowIcon(QIcon("icon.png"))

        self.layout = QVBoxLayout(self)
        self.setLayout(self.layout)

        # -------- العنوان -------- #
        self.title = QLabel("Windows Dev Tools Installer (WDT)")
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title.setFont(QFont("Segoe UI", 20, QFont.Weight.Bold))
        self.title.setStyleSheet("color: #00aaff; margin-bottom: 4px;")
        self.layout.addWidget(self.title)

        # مسار السكربت
        self.ps1_label = QLabel(f"Script: {self.ps1_path}")
        self.ps1_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.ps1_label.setStyleSheet("color: #bbbbbb; font-size: 10pt;")
        self.layout.addWidget(self.ps1_label)

        # -------- مربع الخطة (إيه اللي هيتثبت) -------- #
        self.plan_title = QLabel("Planned Installations (قراءة من installer.ps1)")
        self.plan_title.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        self.plan_title.setStyleSheet("color: #ffffff; margin-top: 10px;")
        self.layout.addWidget(self.plan_title)

        self.plan_output = QTextEdit()
        self.plan_output.setReadOnly(True)
        self.plan_output.setStyleSheet("""
            background-color: #181818;
            color: #f0f0f0;
            border-radius: 8px;
            font-family: Consolas, monospace;
            font-size: 10pt;
            padding: 8px;
        """)
        self.plan_output.setFixedHeight(170)
        self.layout.addWidget(self.plan_output)

        # -------- مربع اللوج -------- #
        self.log_title = QLabel("Installation Log")
        self.log_title.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        self.log_title.setStyleSheet("color: #ffffff; margin-top: 6px;")
        self.layout.addWidget(self.log_title)

        self.output = QTextEdit()
        self.output.setReadOnly(True)
        self.output.setStyleSheet("""
            background-color: #121212;
            color: #e0e0e0;
            border-radius: 10px;
            font-family: Consolas, monospace;
            font-size: 11pt;
            padding: 10px;
        """)
        self.output.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.layout.addWidget(self.output, 1)

        # -------- صف الأزرار -------- #
        buttons_layout = QHBoxLayout()

        # زر تشغيل التثبيت
        self.start_btn = QPushButton("Start Installation")
        self.start_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.start_btn.setFont(QFont("Segoe UI", 13))
        self.start_btn.setFixedHeight(45)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #007ACC, stop:1 #005F9E);
                color: white;
                border-radius: 14px;
                padding: 10px 25px;
                border: none;
            }
            QPushButton:hover {
                background-color: #005F9E;
            }
            QPushButton:pressed {
                background-color: #003F6B;
            }
            QPushButton:disabled {
                background-color: #444444;
                color: #888888;
            }
        """)
        shadow = QGraphicsDropShadowEffect()
        shadow.setBlurRadius(25)
        shadow.setXOffset(0)
        shadow.setYOffset(6)
        shadow.setColor(QColor(0, 122, 204, 180))
        self.start_btn.setGraphicsEffect(shadow)
        buttons_layout.addWidget(self.start_btn)

        # زر تحديث الخطة (لو عدلت على installer.ps1)
        self.reload_plan_btn = QPushButton("Reload Plan")
        self.reload_plan_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.reload_plan_btn.setFont(QFont("Segoe UI", 13))
        self.reload_plan_btn.setFixedHeight(45)
        self.reload_plan_btn.setStyleSheet("""
            QPushButton {
                background-color: #555555;
                color: #f0f0f0;
                border-radius: 14px;
                padding: 10px 25px;
                border: none;
            }
            QPushButton:hover {
                background-color: #666666;
            }
            QPushButton:pressed {
                background-color: #3f3f3f;
            }
        """)
        buttons_layout.addWidget(self.reload_plan_btn)

        # زر مسح اللوج
        self.clear_btn = QPushButton("Clear Log")
        self.clear_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.clear_btn.setFont(QFont("Segoe UI", 13))
        self.clear_btn.setFixedHeight(45)
        self.clear_btn.setStyleSheet("""
            QPushButton {
                background-color: #444444;
                color: #cccccc;
                border-radius: 14px;
                padding: 10px 25px;
                border: none;
            }
            QPushButton:hover {
                background-color: #555555;
            }
            QPushButton:pressed {
                background-color: #333333;
            }
        """)
        buttons_layout.addWidget(self.clear_btn)

        self.layout.addLayout(buttons_layout)

        # -------- الإشارات -------- #
        self.start_btn.clicked.connect(self.start_installation)
        self.clear_btn.clicked.connect(self.output.clear)
        self.reload_plan_btn.clicked.connect(self.load_install_plan)

        # ثريد التثبيت
        self.installer_thread = InstallerThread(self.ps1_path)
        self.installer_thread.output_signal.connect(self.append_output)
        self.installer_thread.finished_signal.connect(self.installation_finished)

        # تحميل الخطة عند بداية التشغيل
        self.load_install_plan()

    # ===================== واجهة: ملخص الخطة ===================== #

    def load_install_plan(self):
        parsed, ok = parse_installer_script(self.ps1_path)
        if not ok:
            self.plan_output.setPlainText(
                f"لم يتم العثور على الملف {PS1_FILE} في نفس مجلد البرنامج.\n"
                f"المسار الحالي:\n{self.ps1_path}"
            )
            return

        plan_text = build_plan_text(parsed)
        self.plan_output.setPlainText(plan_text)

    # ===================== واجهة: اللوج ===================== #

    def append_output(self, text: str):
        # تلوين بعض الكلمات المهمة
        cursor = self.output.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)
        fmt = QTextCharFormat()

        lower = text.lower()
        if "error" in lower or "failed" in lower or "not found" in lower:
            fmt.setForeground(QColor("#ff5555"))  # أحمر للأخطاء
        elif "success" in lower or "completed" in lower or "done" in lower:
            fmt.setForeground(QColor("#55ff55"))  # أخضر للنجاح
        elif "warning" in lower:
            fmt.setForeground(QColor("#ffaa00"))  # برتقالي للتحذيرات
        else:
            fmt.setForeground(QColor("#e0e0e0"))  # رمادي فاتح للنص العادي

        cursor.setCharFormat(fmt)
        cursor.insertText(text + "\n")
        self.output.setTextCursor(cursor)
        self.output.ensureCursorVisible()

    # ===================== بدء التثبيت ===================== #

    def start_installation(self):
        if not os.path.exists(self.ps1_path):
            QMessageBox.critical(
                self,
                "Script Not Found",
                f"لم يتم العثور على {PS1_FILE}.\nتأكد أنه موجود في نفس مجلد البرنامج."
            )
            return

        self.output.clear()
        self.append_output("Starting installation...\n")
        self.start_btn.setDisabled(True)
        self.reload_plan_btn.setDisabled(True)

        self.installer_thread.start()

    def installation_finished(self, code: int):
        if code == 0:
            self.append_output("\nInstallation completed successfully.")
        else:
            self.append_output(f"\nInstallation finished with errors. Exit code: {code}")
        self.start_btn.setDisabled(False)
        self.reload_plan_btn.setDisabled(False)


# ===================== نقطة البداية ===================== #

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # لازم تتحط قبل الـ QSS

    qss_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "more.qss")
    if os.path.exists(qss_path):
        with open(qss_path, "r", encoding="utf-8") as f:
            app.setStyleSheet(f.read())

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
