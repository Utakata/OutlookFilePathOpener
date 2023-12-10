import os
import subprocess
import sys

from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QMessageBox, QDialog

from DatabaseHandler import DatabaseHandler


class ShortcutCreationThread(QThread):
    shortcutCreated = pyqtSignal()
    shortcutCreationFailed = pyqtSignal(str)

    def run(self):
        app_path = os.path.abspath(sys.argv[0])
        shortcut_name = "OutlookFilePathOpener"
        startup_dir = DatabaseHandler.get_startup_folder_path()
        shortcut_path = os.path.join(startup_dir, f"{shortcut_name}.lnk")

        powershell_script = f"$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('{shortcut_path}'); $s.TargetPath = '{app_path}'; $s.Save()"
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE

        try:
            subprocess.run(["powershell", "-Command", powershell_script], startupinfo=startupinfo, check=True)
            # noinspection PyUnresolvedReferences
            self.shortcutCreated.emit()  # 正しい方法でシグナルを発生させる
        except subprocess.CalledProcessError as e:
            # noinspection PyUnresolvedReferences
            self.shortcutCreationFailed.emit(str(e))  # 正しい方法でシグナルを発生させる


class SettingsWindow(QDialog):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.thread = ShortcutCreationThread()
        self.databaseHandler = DatabaseHandler(self.main_window, load_history=False)
        self.shortcutCheckbox = None
        self.init_ui()

    # その他のメソッド...

    def create_shortcut(self):
        if self.shortcutCheckbox.isChecked():
            self.thread = ShortcutCreationThread()
            # noinspection PyUnresolvedReferences
            self.thread.shortcutCreated.connect(self.on_shortcut_created)
            # noinspection PyUnresolvedReferences
            self.thread.shortcutCreationFailed.connect(self.on_shortcut_creation_failed)
            self.thread.start()

    def on_shortcut_created(self):
        QMessageBox.information(self, "PC起動時に自動起動をONにする", "自動起動用ショートカットが正常に作成されました.")

    def on_shortcut_creation_failed(self, error):
        QMessageBox.warning(self, "エラー", f"自動起動用ショートカットの作成中にエラーが発生しました: {error}")
