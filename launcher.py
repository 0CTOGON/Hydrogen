import sys
import subprocess
import webbrowser
from pathlib import Path
from typing import List, Dict

import keyboard
from ddgs import DDGS
from PySide6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
)
from PySide6.QtCore import Qt, QThread, Signal, QSize, Slot
from PySide6.QtGui import QIcon, QPixmap
import win32com.client


class SearchWorker(QThread):
    """Fetch search results in background thread."""

    finished = Signal(list)

    def __init__(self, query: str):
        super().__init__()
        self.query = query

    def run(self):
        try:
            print(f"Searching for: {self.query}")

            results = []
            ddgs = DDGS()

            for r in ddgs.text(self.query, max_results=5):
                results.append({
                    "title": r.get("title", ""),
                    "url": r.get("href", ""),
                    "snippet": r.get("body", "")
                })

            print(f"Found {len(results)} results")

            self.finished.emit(results)

        except Exception as e:
            print(f"Search error: {e}")
            self.finished.emit([])


class IconExtractor:
    """Extract icons from Windows shortcuts."""

    @staticmethod
    def get_icon(path: str) -> QIcon:
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(path)

            icon_path = shortcut.IconLocation

            if icon_path:
                exe_path = (
                    icon_path.split(',')[0]
                    if ',' in icon_path
                    else icon_path
                )

                if Path(exe_path).exists():
                    pixmap = QPixmap(exe_path)

                    if not pixmap.isNull():
                        return QIcon(pixmap)

        except Exception:
            pass

        pixmap = QPixmap(32, 32)
        pixmap.fill(Qt.transparent)

        return QIcon(pixmap)


class AppDetector:
    """Detect installed applications on Windows."""

    @staticmethod
    def get_apps() -> List[Dict]:
        apps = []

        apps.append({
            "name": "Calculator",
            "path": "calc.exe",
            "type": "builtin"
        })

        apps.append({
            "name": "Web Search",
            "path": "web_search",
            "type": "builtin"
        })

        start_menu = (
            Path.home()
            / "AppData/Roaming/Microsoft/Windows/Start Menu/Programs"
        )

        if start_menu.exists():
            for item in start_menu.rglob("*.lnk"):
                apps.append({
                    "name": item.stem,
                    "path": str(item),
                    "type": "shortcut"
                })

        # Remove duplicates
        seen = set()
        unique_apps = []

        for app in sorted(apps, key=lambda x: x["name"]):
            if app["name"] not in seen:
                unique_apps.append(app)
                seen.add(app["name"])

        return unique_apps


class LauncherWindow(QMainWindow):

    toggle_signal = Signal()

    def __init__(self):
        super().__init__()

        self.apps = AppDetector.get_apps()
        self.filtered_apps = self.apps

        self.current_view = "apps"
        self.search_worker = None

        self.setWindowTitle("Zero Launcher")

        self.setWindowFlags(
            Qt.FramelessWindowHint |
            Qt.WindowStaysOnTopHint |
            Qt.Tool
        )

        self.setAttribute(Qt.WA_TranslucentBackground)

        self.setFixedWidth(520)

        self.apply_style()
        self.setup_ui()

        self.toggle_signal.connect(self.toggle_window)

        # Hide results on startup
        self.results.setVisible(False)

        try:
            keyboard.add_hotkey(
                'ctrl+shift+s',
                lambda: self.toggle_signal.emit()
            )

        except Exception as e:
            print(f"Hotkey error: {e}")

        self.hide()

    def apply_style(self):
        style = """
            QMainWindow {
                background-color: #0a0e27;
                border-radius: 20px;
            }

            QLineEdit {
                background-color: #1a1f3a;
                color: #ffffff;
                border: 2px solid #00d9ff;
                border-radius: 14px;
                padding: 14px 20px;
                font-size: 15px;
                font-weight: 600;
                font-family: 'Segoe UI', Arial;
                selection-background-color: #00d9ff;
                selection-color: #0a0e27;
            }

            QLineEdit:focus {
                border: 2px solid #00ffff;
            }

            QLineEdit::placeholder {
                color: #7ddfff;
            }

            QPushButton {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 #1e3a8a,
                    stop:1 #2d5a8c
                );

                color: #f0f0ff;

                border: none;
                border-radius: 14px;

                padding: 14px 18px;

                font-size: 15px;
                font-weight: 700;
            }

            QPushButton:hover {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 #2555b0,
                    stop:1 #3a6fb5
                );
            }

            QPushButton:pressed {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:1,
                    stop:0 #1a2d6a,
                    stop:1 #233a6f
                );
            }

            QListWidget {
                background-color: #0a0e27;
                border: none;
                outline: none;
                margin-top: 6px;
            }

            QListWidget::item {
                padding: 16px 18px;
                color: #f0f0ff;
                margin-bottom: 6px;
                border-radius: 10px;
                min-height: 48px;
            }

            QListWidget::item:hover {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #1a1f3a,
                    stop:1 #252d4a
                );
            }

            QListWidget::item:selected {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #00d9ff,
                    stop:1 #00a8cc
                );

                color: #0a0e27;
                font-weight: 700;
            }

            QScrollBar:vertical {
                background-color: #0a0e27;
                width: 14px;
                border-radius: 7px;
            }

            QScrollBar::handle:vertical {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #00d9ff,
                    stop:1 #00a8cc
                );

                border-radius: 7px;
                min-height: 24px;
            }

            QScrollBar::handle:vertical:hover {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #00ffff,
                    stop:1 #00d9ff
                );
            }
        """

        self.setStyleSheet(style)

    def setup_ui(self):

        central = QWidget()

        central.setStyleSheet("""
            QWidget {
                background-color: #0a0e27;
                border-radius: 16px;
            }
        """)

        self.main_layout = QVBoxLayout()

        self.main_layout.setContentsMargins(16, 16, 16, 16)
        self.main_layout.setSpacing(10)

        # Search row
        search_layout = QHBoxLayout()
        search_layout.setSpacing(10)

        self.search_input = QLineEdit()

        self.search_input.setPlaceholderText("Search...")
        self.search_input.setFixedHeight(48)

        self.search_input.textChanged.connect(self.filter_apps)

        search_layout.addWidget(self.search_input)

        self.search_button = QPushButton("🔍")
        self.search_button.setFixedSize(54, 48)

        self.search_button.clicked.connect(self.do_web_search)

        search_layout.addWidget(self.search_button)

        search_container = QWidget()
        search_container.setLayout(search_layout)

        self.main_layout.addWidget(search_container)

        # Results list
        self.results = QListWidget()

        self.results.setWordWrap(True)
        self.results.setUniformItemSizes(False)

        self.results.itemDoubleClicked.connect(self.launch_app)
        self.results.itemPressed.connect(self.on_item_click)

        self.results.hide()

        self.main_layout.addWidget(self.results)

        # Back button
        self.back_button = QPushButton("← Back to Apps")

        self.back_button.setFixedHeight(48)

        self.back_button.hide()

        self.back_button.clicked.connect(self.back_to_apps)

        self.main_layout.addWidget(self.back_button)

        central.setLayout(self.main_layout)

        self.setCentralWidget(central)

        self.update_results()

    def filter_apps(self):
        try:
            query = self.search_input.text().lower()

            self.filtered_apps = [
                app for app in self.apps
                if query in app["name"].lower()
            ]

            if self.current_view == "search":
                self.back_to_apps()

            self.update_results()

        except Exception as e:
            print(f"Filter error: {e}")

    def update_results(self):
        try:
            self.results.clear()
            self.results.show()

            for app in self.filtered_apps:
                item = QListWidgetItem(app["name"])

                item.setData(Qt.UserRole, app["path"])
                item.setData(Qt.UserRole + 1, app["type"])

                item.setSizeHint(QSize(0, 42))

                self.results.addItem(item)

            visible_items = min(len(self.filtered_apps), 6)

            if visible_items == 0:
                results_height = 80
            else:
                results_height = visible_items * 62

            total_height = 100 + results_height

            self.setFixedHeight(total_height)

        except Exception as e:
            print(f"Update results error: {e}")

    def do_web_search(self):

        query = self.search_input.text()

        if query:

            self.search_button.setText("⏳")
            self.search_button.setEnabled(False)

            if self.search_worker:
                self.search_worker.quit()
                self.search_worker.wait()

            self.search_worker = SearchWorker(query)

            self.search_worker.finished.connect(
                self.display_search_results
            )

            self.search_worker.start()

    def display_search_results(self, results):

        self.search_button.setText("🔍")
        self.search_button.setEnabled(True)

        if not results:
            return

        self.results.clear()

        for r in results:

            item = QListWidgetItem()

            text = (
                f"{r['title']}\n"
                f"{r['snippet'][:90]}...\n"
                f"{r['url'][:60]}"
            )

            item.setText(text)

            item.setData(Qt.UserRole, r["url"])
            item.setData(Qt.UserRole + 1, "search_result")

            item.setSizeHint(QSize(640, 110))

            self.results.addItem(item)

        self.current_view = "search"

        self.back_button.show()

        self.setFixedHeight(520)

    def back_to_apps(self):

        self.current_view = "apps"

        self.back_button.hide()

        self.filter_apps()

    def launch_app(self, item):
        self.on_item_click(item)

    def on_item_click(self, item):

        path = item.data(Qt.UserRole)
        app_type = item.data(Qt.UserRole + 1)

        try:

            if app_type == "shortcut":

                subprocess.Popen(
                    f'start "" "{path}"',
                    shell=True
                )

            elif app_type == "builtin":

                if path == "calc.exe":
                    subprocess.Popen("calc.exe")

                elif path == "web_search":
                    self.do_web_search()

            elif app_type == "search_result":

                webbrowser.open(path)

            self.hide()

        except Exception as e:
            print(f"Error: {e}")

    @Slot()
    def toggle_window(self):

        try:

            if self.isVisible():

                self.hide()

            else:

                self.current_view = "apps"

                self.back_button.hide()

                self.search_input.clear()

                self.filter_apps()

                self.show()

                # Center on screen
                screen = QApplication.primaryScreen().availableGeometry()

                x = (
                    screen.width() - self.width()
                ) // 2

                y = (
                    screen.height() - self.height()
                ) // 2

                self.move(x, y)

                self.raise_()
                self.activateWindow()

                self.search_input.setFocus()

        except Exception as e:
            print(f"Toggle error: {e}")


if __name__ == "__main__":

    app = QApplication(sys.argv)

    launcher = LauncherWindow()

    launcher.hide()

    sys.exit(app.exec())