import os
import sys
from pathlib import Path

# --- תיקון אייקון עבור שורת המשימות ב-Windows ---
if sys.platform == "win32":
    import ctypes

    myappid = "mycompany.pptxmerger.v1"
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
# --------------------------------------------------

from PySide6.QtCore import Qt
from PySide6.QtGui import QDragEnterEvent, QDropEvent, QPixmap, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QTreeWidget,
    QTreeWidgetItem,
    QFileDialog,
    QMessageBox,
    QLabel,
    QProgressBar
)

from main import merge_presentations


class FlatTreeWidget(QTreeWidget):
    def dropEvent(self, event):
        # מציאת הפריט שעליו מנסים לשחרר את הגרירה
        target = self.itemAt(event.position().toPoint())

        # אם מנסים לשחרר את הקובץ ישירות "בתוך" פריט קיים, נמנע זאת ונאלץ אותו להישאר כאיבר עצמאי ברשימה
        if target:
            # מנטרל הדבקה היררכית פנימית
            event.setDropAction(Qt.MoveAction)

        super().dropEvent(event)


class PPTXMergerUI(QWidget):
    def __init__(self):
        super().__init__()
        self.files = []
        self.completed_count = 0
        self.setWindowTitle("ממזג המצגות")
        self.resize(550, 440)

        if hasattr(sys, '_MEIPASS'):
            self.icon_path = os.path.join(sys._MEIPASS, "fire_icon.ico")
        else:
            self.icon_path = "fire_icon.ico"

        self.setWindowIcon(QIcon(self.icon_path))
        self.setAcceptDrops(True)
        self.build_ui()

    def build_ui(self):
        layout = QVBoxLayout()

        layout.addSpacing(4)

        buttons_layout = QHBoxLayout()

        self.add_btn = QPushButton("הוסף קבצים")
        self.add_btn.clicked.connect(self.add_files)
        buttons_layout.addWidget(self.add_btn)

        self.add_folder_btn = QPushButton("הוסף תיקייה")
        self.add_folder_btn.clicked.connect(self.add_folder)
        buttons_layout.addWidget(self.add_folder_btn)

        self.remove_btn = QPushButton("הסר נבחר")
        self.remove_btn.clicked.connect(self.remove_selected)
        buttons_layout.addWidget(self.remove_btn)

        self.move_up_btn = QPushButton("▲")
        self.move_up_btn.clicked.connect(self.move_up)
        buttons_layout.addWidget(self.move_up_btn)

        self.move_down_btn = QPushButton("▼")
        self.move_down_btn.clicked.connect(self.move_down)
        buttons_layout.addWidget(self.move_down_btn)

        self.clear_btn = QPushButton("נקה הכל")
        self.clear_btn.clicked.connect(self.clear_files)
        buttons_layout.addWidget(self.clear_btn)

        layout.addLayout(buttons_layout)

        self.tree_widget = FlatTreeWidget()
        self.tree_widget.setColumnCount(3)
        self.tree_widget.setHeaderLabels(["שם הקובץ", "גודל", "סטטוס"])

        self.tree_widget.setColumnWidth(0, 240)
        self.tree_widget.setColumnWidth(1, 140)
        self.tree_widget.setColumnWidth(2, 120)

        self.tree_widget.setDragDropMode(QTreeWidget.InternalMove)
        self.tree_widget.setDefaultDropAction(Qt.MoveAction)
        self.tree_widget.setSelectionMode(QTreeWidget.SingleSelection)
        self.tree_widget.setAcceptDrops(True)
        self.tree_widget.setDropIndicatorShown(True)

        self.tree_widget.dragEnterEvent = self.dragEnterEvent
        self.tree_widget.dragMoveEvent = self.dragMoveEvent
        self.tree_widget.dropEvent = self.dropEvent

        self.tree_widget.model().rowsMoved.connect(lambda: self.sync_files_order())

        layout.addWidget(self.tree_widget)

        self.progress = QProgressBar()
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        layout.addWidget(self.progress)

        self.merge_btn = QPushButton("מזג הכל")
        self.merge_btn.setStyleSheet("font-size: 20px; font-weight: bold; padding: 15px;")
        self.merge_btn.clicked.connect(self.merge)
        layout.addWidget(self.merge_btn)

        self.setLayout(layout)

    # --------------------------------
    # מנגנון DRAG & DROP החכם
    # --------------------------------

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls() or event.source() == self.tree_widget:
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls() or event.source() == self.tree_widget:
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
            for url in event.mimeData().urls():
                path = Path(url.toLocalFile())
                if path.is_dir():
                    self.scan_and_add_folder(path)
                elif path.suffix.lower() == ".pptx":
                    self.add_single_file(path)

        elif event.source() == self.tree_widget:
            FlatTreeWidget.dropEvent(self.tree_widget, event)
            self.sync_files_order()
        else:
            event.ignore()

    # --------------------------------
    # FILE & FOLDER MANAGEMENT
    # --------------------------------

    def get_file_size_str(self, path):
        try:
            size_bytes = path.stat().st_size
            if size_bytes < 1024 * 1024:
                return f"{size_bytes / 1024:.1f} KB"
            return f"{size_bytes / (1024 * 1024):.1f} MB"
        except:
            return "לא ידוע"

    def add_single_file(self, path):
        if path not in self.files:
            self.files.append(path)
            size_str = self.get_file_size_str(path)
            item = QTreeWidgetItem([path.name, size_str, "ממתין"])
            item.setData(0, Qt.UserRole, str(path))

            # התיקון הקריטי: מכבים את האפשרות שפריטים אחרים ייזרקו *בתוך* הפריט הזה
            item.setFlags(item.flags() & ~Qt.ItemIsDropEnabled)

            self.tree_widget.addTopLevelItem(item)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "בחר קבצי PowerPoint", "", "PowerPoint Files (*.pptx)"
        )
        for file in files:
            self.add_single_file(Path(file))

    def add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "בחר תיקייה המכילה קבצי PPTX")
        if folder:
            self.scan_and_add_folder(Path(folder))

    def scan_and_add_folder(self, folder_path):
        pptx_files = sorted(folder_path.glob("*.pptx"))
        if not pptx_files:
            QMessageBox.warning(self, "התראה", f"לא נמצאו מצגות (PPTX) בתיקייה:\n{folder_path.name}")
            return
        for file_path in pptx_files:
            self.add_single_file(file_path)

    def remove_selected(self):
        self.sync_files_order()
        root = self.tree_widget.invisibleRootItem()
        current_item = self.tree_widget.currentItem()
        if not current_item:
            return
        index = root.indexOfChild(current_item)
        if index != -1:
            root.removeChild(current_item)
            del self.files[index]

    def clear_files(self):
        self.files.clear()
        self.tree_widget.clear()

    # --------------------------------
    # MOVE BUTTONS
    # --------------------------------

    def move_up(self):
        self.sync_files_order()
        root = self.tree_widget.invisibleRootItem()
        current_item = self.tree_widget.currentItem()
        if not current_item:
            return
        current_row = root.indexOfChild(current_item)
        if current_row <= 0:
            return
        root.removeChild(current_item)
        root.insertChild(current_row - 1, current_item)
        self.tree_widget.setCurrentItem(current_item)
        self.sync_files_order()

    def move_down(self):
        self.sync_files_order()
        root = self.tree_widget.invisibleRootItem()
        current_item = self.tree_widget.currentItem()
        if not current_item:
            return
        current_row = root.indexOfChild(current_item)
        if current_row == -1 or current_row >= root.childCount() - 1:
            return
        root.removeChild(current_item)
        root.insertChild(current_row + 1, current_item)
        self.tree_widget.setCurrentItem(current_item)
        self.sync_files_order()

    # --------------------------------
    # SYNC ORDER & UTILS
    # --------------------------------

    def sync_files_order(self):
        new_order = []
        # התיקון כאן מבטיח שגם אם קובץ בטעות נכנס פנימה, המערכת תסרוק רק את ה-Top Level ותשטח אותו
        for i in range(self.tree_widget.topLevelItemCount()):
            item = self.tree_widget.topLevelItem(i)
            if item:
                full_path_str = item.data(0, Qt.UserRole)
                if full_path_str:
                    new_order.append(Path(full_path_str))
        self.files = new_order

    def set_ui_enabled(self, enabled: bool):
        self.add_btn.setEnabled(enabled)
        self.add_folder_btn.setEnabled(enabled)
        self.remove_btn.setEnabled(enabled)
        self.move_up_btn.setEnabled(enabled)
        self.move_down_btn.setEnabled(enabled)
        self.clear_btn.setEnabled(enabled)
        self.tree_widget.setEnabled(enabled)
        self.merge_btn.setEnabled(enabled)

    def open_file_location(self, file_path_str):
        import subprocess
        path = os.path.normpath(file_path_str)
        CREATE_NO_WINDOW = 0x08000000
        subprocess.Popen(['explorer', '/select,', path], creationflags=CREATE_NO_WINDOW)

    # --------------------------------
    # MERGE, PROGRESS & STATUS UPDATER
    # --------------------------------

    def find_item_by_path(self, file_path):
        for i in range(self.tree_widget.topLevelItemCount()):
            item = self.tree_widget.topLevelItem(i)
            if item and item.data(0, Qt.UserRole) == str(file_path):
                return item, i
        return None, -1

    def update_progress(self, completed_file_path):
        total_files = len(self.files)
        if total_files == 0:
            return

        completed_item, current_idx = self.find_item_by_path(completed_file_path)
        if completed_item:
            completed_item.setText(2, "סיים")

        if current_idx != -1 and current_idx + 1 < total_files:
            next_file_path = self.files[current_idx + 1]
            next_item, _ = self.find_item_by_path(next_file_path)
            if next_item:
                next_item.setText(2, "ממזג...")

        self.completed_count += 1
        actual_percent = int((self.completed_count / total_files) * 100)
        self.progress.setValue(min(actual_percent, 100))

        QApplication.processEvents()

    def merge(self):
        if len(self.files) < 2:
            QMessageBox.critical(self, "שגיאה", "צריך לפחות 2 מצגות (.PPTX)")
            return

        self.sync_files_order()

        output, _ = QFileDialog.getSaveFileName(
            self, "שמור קובץ ממוזג", "merged.pptx", "PowerPoint Files (*.pptx)"
        )
        if not output:
            return

        try:
            self.set_ui_enabled(False)
            self.completed_count = 0

            for i in range(self.tree_widget.topLevelItemCount()):
                item = self.tree_widget.topLevelItem(i)
                if item:
                    item.setText(2, "ממתין")

            if len(self.files) > 0:
                first_item, _ = self.find_item_by_path(self.files[0])
                if first_item:
                    first_item.setText(2, "ממזג...")

            self.progress.setRange(0, 100)
            self.progress.setValue(0)
            QApplication.processEvents()

            merge_presentations(self.files, output, progress_callback=self.update_progress)

            self.progress.setValue(100)

            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Icon.NoIcon)
            msg.setWindowTitle("הקובץ נוצר בהצלחה!")
            msg.setText(" **המיזוג הושלם בהצלחה!**")
            msg.setInformativeText(f"הקובץ הממוזג נשמר בנתיב:\n{output}")

            open_folder_btn = msg.addButton("פתח מיקום קובץ", QMessageBox.ButtonRole.ActionRole)
            close_btn = msg.addButton("סגור", QMessageBox.ButtonRole.AcceptRole)
            msg.setDefaultButton(close_btn)

            msg.setStyleSheet("""
                QLabel { font-size: 14px; color: #2e7d32; font-weight: bold; }
                QLabel#qt_msgbox_informativetext { font-size: 12px; color: #333333; font-weight: normal; }
                QPushButton { font-weight: bold; padding: 6px 15px; border-radius: 4px; }
                QPushButton[text="סגור"] { background-color: #e0e0e0; color: #333333; }
                QPushButton[text="סגור"]:hover { background-color: #d5d5d5; }
                QPushButton[text="פתח מיקום קובץ"] { background-color: #4caf50; color: white; }
                QPushButton[text="פתח מיקום קובץ"]:hover { background-color: #43a047; }
            """)

            msg.exec()

            if msg.clickedButton() == open_folder_btn:
                self.open_file_location(output)

            self.progress.setValue(0)

        except Exception as e:
            self.progress.setValue(0)
            QMessageBox.critical(self, "שגיאה", f"המיזוג נכשל:\n{str(e)}")

        finally:
            for i in range(self.tree_widget.topLevelItemCount()):
                item = self.tree_widget.topLevelItem(i)
                if item:
                    item.setText(2, "ממתין")
            self.set_ui_enabled(True)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PPTXMergerUI()
    window.show()
    sys.exit(app.exec())