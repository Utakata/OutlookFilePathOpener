import os
from PyQt5 import QtWidgets, QtGui
import sqlite3
from PyQt5.QtWidgets import QMessageBox, QDialog
from DatabaseHandler import DatabaseHandler
from ShortcutCreationThread import ShortcutCreationThread

conn = sqlite3.connect('database.db')


class SettingsWindow(QDialog):
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.thread = ShortcutCreationThread()  # self.thread の初期化をコンストラクタで行う
        self.databaseHandler = DatabaseHandler(self.main_window, load_history=False)
        self.shortcutCheckbox = None
        self.init_ui()

    # noinspection PyUnresolvedReferences
    def init_ui(self):
        # ウィンドウのタイトルとアイコンを設定
        self.setWindowTitle('設定')
        self.setWindowIcon(QtGui.QIcon('app2.ico'))

        # ウィンドウの位置とサイズを設定
        self.setGeometry(300, 300, 300, 200)

        # ショートカット作成のチェックボックスを追加
        self.shortcutCheckbox = QtWidgets.QCheckBox('PC起動時に自動起動をONにする', self)
        self.shortcutCheckbox.setGeometry(10, 10, 280, 30)
        self.thread.shortcutCreated.connect(self.on_shortcut_created)
        self.thread.shortcutCreationFailed.connect(self.on_shortcut_creation_failed)

        # DatabaseHandlerを使用して、ショートカット作成フラグを設定
        self.databaseHandler = DatabaseHandler(self, load_history=False)
        # ショートカットが作成されているかどうかを確認
        if self.databaseHandler.is_shortcut_created("OutlookFilePathOpener"):
            self.shortcutCheckbox.setChecked(True)

        # OKボタンの設定
        ok_button = QtWidgets.QPushButton('OK', self)
        ok_button.setGeometry(10, 150, 280, 30)
        ok_button.clicked.connect(self.on_ok_clicked)

        # 履歴をクリアするボタンの設定
        clear_button = QtWidgets.QPushButton('履歴をクリア', self)
        clear_button.setGeometry(10, 50, 280, 30)
        clear_button.clicked.connect(self.on_clear_history_clicked)

    def create_shortcut(self):
        if self.shortcutCheckbox.isChecked():
            self.thread = ShortcutCreationThread()
            self.thread.shortcutCreated.connect(self.on_shortcut_created)
            self.thread.shortcutCreationFailed.connect(self.on_shortcut_creation_failed)
            self.thread.start()

    def on_shortcut_created(self):
        QMessageBox.information(self, "PC起動時に自動起動をONにする", "自動起動用ショートカットが正常に作成されました.")

    def on_shortcut_creation_failed(self, error):
        QMessageBox.warning(self, "エラー", f"自動起動用ショートカットの作成中にエラーが発生しました: {error}")

    def delete_shortcut(self):
        shortcut_name = "OutlookFilePathOpener"
        startup_dir = DatabaseHandler.get_startup_folder_path()  # DatabaseHandler クラスから直接呼び出す
        shortcut_path = os.path.join(startup_dir, f"{shortcut_name}.lnk")

        # ショートカットが存在するか確認
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)  # ショートカットを削除
            QMessageBox.information(self, "PC起動時に自動起動をOFFにする", "自動起動用ショートカットが正常に削除されました。")
        else:
            QMessageBox.information(self, "PC起動時に自動起動をOFFにする", "自動起動用ショートカットは存在しません。")

    def on_clear_history_clicked(self):
        # 履歴をクリアする処理を実装
        self.databaseHandler.clear_history()
        QMessageBox.information(self, "履歴クリア", "履歴がクリアされました。")

        # MainWindow インスタンスのコンボボックスを更新
        if hasattr(self, 'main_window'):
            # 履歴データを再読み込みし、コンボボックスを更新
            history_data = self.databaseHandler.load_history()
            self.main_window.update_combo_box(history_data)
        else:
            QMessageBox.warning(self, "エラー", "親ウィンドウの参照が不正です。")

    def on_ok_clicked(self):
        # チェックボックスがチェックされているか確認
        if self.shortcutCheckbox.isChecked():
            # ショートカットが既に存在するか確認
            if not self.databaseHandler.is_shortcut_created("OutlookFilePathOpener"):
                # ショートカットが存在しない場合、ショートカットを作成
                self.create_shortcut()
        else:
            # チェックボックスが未チェックの場合、ショートカットを削除
            self.delete_shortcut()

        # DBハンドラーを閉じる
        self.databaseHandler.close()
        # 設定ウィンドウを閉じる
        self.close()

    def closeEvent(self, event):
        self.databaseHandler.close()  # データベース接続を閉じる
        super().closeEvent(event)
