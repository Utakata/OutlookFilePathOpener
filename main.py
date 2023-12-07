import os
import re
import sqlite3
import subprocess
import sys

from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtCore import QThread
from PyQt5.QtWidgets import QMessageBox
from qt_material import apply_stylesheet


# UIの構造を定義するクラス
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.fileOpeningThread = None
        self.label = None
        self.descriptionLabel = None
        self.iconLabel = None
        self.comboBox = None
        self.init_ui()
        self.databaseHandler = DatabaseHandler(self)
        self.load_history_from_db()
    def init_ui(self):
        # ウィンドウの設定
        self.setWindowTitle('ファイル名を指定して実行')  # ウィンドウのタイトルを設定
        self.setWindowIcon(QtGui.QIcon('app2.ico'))  # ウィンドウのアイコンを設定
        self.setFixedSize(400, 150)  # ウィンドウの固定サイズを設定
        self.setStyleSheet("background-color: white;")  # ウィンドウの背景色を設定

        # コンボボックスの初期化と定義
        self.comboBox = QtWidgets.QComboBox(self)  # comboBox オブジェクトを作成し、親ウィンドウを指定
        self.comboBox.setEditable(True)  # コンボボックスを編集可能に設定
        self.comboBox.setGeometry(50, 40, 338, 30)  # コンボボックスの位置とサイズを設定

        # アイコンの追加
        self.iconLabel = QtWidgets.QLabel(self)
        pixmap = QtGui.QPixmap('app2.ico')
        # アイコンの位置を調整
        self.iconLabel.setGeometry(10, 10, 30, 30)  # X座標とY座標をそれぞれ増やす
        scaled_pixmap = pixmap.scaled(30, 30, QtCore.Qt.KeepAspectRatio, QtCore.Qt.SmoothTransformation)
        self.iconLabel.setPixmap(scaled_pixmap)

        # 説明文のラベルの配置と位置を調整
        self.descriptionLabel = QtWidgets.QLabel('フォルダまたは、ファイルのパスを入力してください。', self)
        self.descriptionLabel.setGeometry(50, 10, 340, 20)  # 位置を調整

        # コンボボックスの位置を調整（説明文の下に）
        self.label = QtWidgets.QLabel('パス:', self)  # ラベルを作成し、親ウィンドウを指定
        self.label.setGeometry(15, 40, 40, 30)  # ラベルの位置とサイズを設定
        self.comboBox.setEditable(True)  # コンボボックスを編集可能に設定
        self.comboBox.setGeometry(50, 40, 338, 30)  # コンボボックスの位置とサイズを設定

        # フレームの作成と設定（縦のサイズを大きくする）
        frame = QtWidgets.QFrame(self)  # フレームを作成し、親ウィンドウを指定
        frame.setGeometry(-10, 85, 410, 70)  # フレームの位置とサイズを設定（縦のサイズを70ピクセルに変更）
        frame.setStyleSheet("background-color: #F0F0F0;")  # フレームの背景色を設定

        # ボタンの作成とサイズ設定
        button1 = QtWidgets.QPushButton('開く', frame)  # ボタン1を作成し、親フレームを指定
        button1.setFixedSize(80, 30)  # ボタンの幅と高さを設定
        button2 = QtWidgets.QPushButton('クリア', frame)  # ボタン2を作成し、親フレームを指定
        button2.setFixedSize(80, 30)  # 同上
        button3 = QtWidgets.QPushButton('設定', frame)  # ボタン3を作成し、親フレームを指定
        button3.setFixedSize(80, 30)  # 同上

        # ボタンのシグナルを設定
        button1.clicked.connect(self.open_file)  # ボタン1がクリックされたときの処理を設定
        button2.clicked.connect(self.clear_combo_box)  # ボタン2がクリックされたときの処理を設定
        button3.clicked.connect(self.open_settings)  # ボタン3がクリックされたときの処理を設定

        # レイアウトの設定（水平方向にボタンを並べて右揃えにする）
        layout = QtWidgets.QHBoxLayout()  # 水平ボックスレイアウトを作成
        layout.addStretch(1)  # スペースを追加して右揃えにする
        layout.addWidget(button1)  # ボタン1をレイアウトに追加
        layout.addWidget(button2)  # ボタン2をレイアウトに追加
        layout.addWidget(button3)  # ボタン3をレイアウトに追加
        frame.setLayout(layout)  # フレームにレイアウトを設定

        # Enterキーが押されたときのシグナルを設定
        self.comboBox.lineEdit().returnPressed.connect(self.open_file)

    def clear_history(self):
        # 履歴をクリアする処理を実装
        self.databaseHandler.clear_history()
        QMessageBox.information(self, "履歴クリア", "履歴がクリアされました。")

    def load_history_from_db(self):
        self.historyLoadingThread = HistoryLoadingThread()
        self.historyLoadingThread.historyLoaded.connect(self.update_combo_box)
        self.historyLoadingThread.start()

    def open_file(self):
        """
        ユーザーが入力したパスに基づいてファイルまたはフォルダを開くメソッド。
        新しいスレッドを使用して非同期にファイルを開き、GUIの応答性を保持します。
        """
        path = self.comboBox.currentText()
        path = re.sub(r'[\r\n\"\'<>\[\]]', '', path)  # 改行と引用符を削除

        if path.startswith("file://"):
            path = path[7:]  # file:// を削除

        # OpenFileThread スレッドを起動してファイルを開く
        self.databaseHandler.save_history(path)  # 入力されたパスを履歴に保存
        self.fileOpeningThread = OpenFileThread(path)
        self.fileOpeningThread.errorOccurred.connect(self.show_error_message)
        self.fileOpeningThread.start()

    def show_error_message(self, message):
        QMessageBox.warning(self, "エラー", message)

    # コンボボックスのクリアボタンがクリックされたときの処理
    def clear_combo_box(self):
        self.comboBox.clearEditText()  # コンボボックスのテキストをクリアする

    # 履歴をクリアする機能
    def clear_history(self):
        # 履歴をクリアする処理を実装
        pass

    # 設定ウィンドウを開く機能
    def open_settings(self):
        settings_window = SettingsWindow(self)  # self を渡して MainWindow を親として設定
        settings_window.exec_()

    def update_combo_box(self, items):
        """コンボボックスを更新するメソッド"""
        self.comboBox.clear()  # この行は関数の内部にあるため、インデントが必要です。
        if items is not None:
            for item in items:
                self.comboBox.addItem(item)

    def closeEvent(self, event):
        if self.historyLoadingThread.isRunning():
            self.historyLoadingThread.terminate()
        event.accept()

# DB関連

# データベースの初期化および操作用クラス
class DatabaseHandler:
    def __init__(self, parent, load_history=True):
        self.parent = parent
        # データベースに接続し、カーソルを作成
        self.conn = sqlite3.connect('settings.db')
        self.cursor = self.conn.cursor()

        # history テーブルが存在しない場合は作成
        self.cursor.execute('CREATE TABLE IF NOT EXISTS history (command TEXT)')

        # shortcut テーブルが存在しない場合は作成
        self.cursor.execute('CREATE TABLE IF NOT EXISTS shortcut (created INTEGER)')

        # 履歴を読み込む
        if load_history:
            self.load_history()

        # ショートカット作成フラグを読み込む
        self.load_shortcut_flag()

        # DB使用後は閉じる関数

    def close(self):
        """データベース接続を閉じる"""
        self.conn.close()

        # 履歴をデータベースから読み込む関数

    def load_history(self):
        # 履歴データを取得して、親ウィンドウのメソッドを通じてコンボボックスに設定
        self.cursor.execute('SELECT command FROM history ORDER BY ROWID DESC LIMIT 6')
        history_data = [row[0] for row in self.cursor.fetchall()]
        return history_data  # 追加: 履歴データを返す

    def load_shortcut_flag(self):
        self.cursor.execute('SELECT created FROM shortcut')
        result = self.cursor.fetchone()
        if result:
            self.parent.shortcutCheckbox.setChecked(bool(result[0]))  # ショートカット作成フラグを親ウィンドウのチェックボックスに設定

    # ショートカット作成フラグを保存する関数
    def save_shortcut_flag(self, created):
        self.cursor.execute('DELETE FROM shortcut')  # 既存のデータを削除
        self.cursor.execute('INSERT INTO shortcut (created) VALUES (?)', (int(created),))  # 新しいデータを挿入
        self.conn.commit()  # 変更をコミット
        # コマンドを履歴に保存する関数

    def save_history(self, command):
        # 履歴を過去5つ以外削除
        self.cursor.execute(
            'DELETE FROM history WHERE ROWID NOT IN (SELECT ROWID FROM history ORDER BY ROWID DESC LIMIT 5)')

        # 新しいコマンドをデータベースに挿入
        self.cursor.execute('INSERT INTO history (command) VALUES (?)', (command,))

        # コミットしてデータベースを更新
        self.conn.commit()

        # 履歴を再読み込み
        self.load_history()

        # 履歴をクリアする関数

    def clear_history(self):
        # 履歴テーブルの全てのデータを削除
        self.cursor.execute('DELETE FROM history')

        # コミットしてデータベースを更新
        self.conn.commit()

        # 履歴を再読み込み
        self.load_history()

    @staticmethod
    def get_startup_folder_path():
        # Windowsでコンソールウィンドウを表示せずにプロセスを実行するための設定
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE

        # PowerShellを使ってスタートアップフォルダのパスを取得
        try:
            startup_path = subprocess.check_output(
                ["powershell", "-Command", "echo $((New-Object -ComObject WScript.Shell).SpecialFolders('Startup'))"],
                text=True,
                startupinfo=startupinfo
            ).strip()
            return startup_path
        except subprocess.CalledProcessError as e:
            print(f"エラー: {e}")
            return None

    def is_shortcut_created(self, shortcut_name):
        startup_dir = self.get_startup_folder_path()  # self を使用して静的メソッドを呼び出す
        shortcut_path = os.path.join(startup_dir, f"{shortcut_name}.lnk")
        return os.path.exists(shortcut_path)
class SettingsWindow(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.thread = None
        # DatabaseHandlerのインスタンスを作成（self.parentを引数として渡す）
        self.databaseHandler = DatabaseHandler(self.parent, load_history=False)
        self.shortcutCheckbox = None
        self.init_ui()

    def init_ui(self):
        # このメソッドでは、設定ウィンドウのUIコンポーネントが定義されます。
        # ウィンドウのタイトルとアイコンを設定
        self.setWindowTitle('設定')
        self.setWindowIcon(QtGui.QIcon('app2.ico'))

        # ウィンドウの位置とサイズを設定
        self.setGeometry(300, 300, 300, 200)

        # ショートカット作成のチェックボックスを追加
        self.shortcutCheckbox = QtWidgets.QCheckBox('PC起動時に自動起動をONにする', self)
        self.shortcutCheckbox.setGeometry(10, 10, 280, 30)

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
        QMessageBox.information(self, "PC起動時に自動起動をONにする", "自動起動用ショートカットが正常に作成されました。")

    def on_shortcut_creation_failed(self, error):
        QMessageBox.warning(self, "エラー", f"自動起動用ショートカットの作成中にエラーが発生しました: {error}")

    def delete_shortcut(self):
        shortcut_name = "OutlookFilePathOpener"
        startup_dir = DatabaseHandler.get_startup_folder_path()  # DatabaseHandler クラスから直接呼び出す
        shortcut_path = os.path.join(startup_dir, f"{shortcut_name}.lnk")

        # ショートカットが存在するか確認
        if os.path.exists(shortcut_path):
            os.remove(shortcut_path)  # ショートカットを削除
            QMessageBox.information(self, "PC起動時に自動起動をOFFにする",
                "自動起動用ショートカットが正常に削除されました。")
        else:
            QMessageBox.information(self, "PC起動時に自動起動をOFFにする", "自動起動用ショートカットは存在しません。")

    def on_clear_history_clicked(self):
        # 履歴をクリアする処理を実装
        self.databaseHandler.clear_history()
        QMessageBox.information(self, "履歴クリア", "履歴がクリアされました。")

        # MainWindow インスタンスのコンボボックスを更新
        main_window = self.parent()
        if isinstance(main_window, MainWindow):
            # 履歴データを再読み込みし、コンボボックスを更新
            history_data = self.databaseHandler.load_history()
            main_window.update_combo_box(history_data)
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


class ShortcutCreationThread(QThread):
    shortcutCreated = QtCore.pyqtSignal()
    shortcutCreationFailed = QtCore.pyqtSignal(str)

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
            self.shortcutCreated.emit()
        except subprocess.CalledProcessError as e:
            self.shortcutCreationFailed.emit(str(e))


# SettingsWindowクラスの create_shortcut メソッドの変更


def shortcut_created(self):
    QMessageBox.information(self, "自動起動をONにする。", "自動起動用ショートカットが正常に作成されました。")


def shortcut_creation_failed(self, error):
    QMessageBox.warning(self, "エラー", f"自動起動用ショートカットの作成中にエラーが発生しました: {error}")


class OpenFileThread(QThread):
    errorOccurred = QtCore.pyqtSignal(str)  # エラーが発生したときにエラーメッセージを通知するシグナル

    def __init__(self, path, parent=None):
        super().__init__(parent)
        self.path = path  # 開くファイルまたはディレクトリのパス

    def run(self):
        if os.path.exists(self.path):  # パスが存在するか確認
            if os.path.isdir(self.path):  # パスがディレクトリであるか確認
                # ディレクトリをエクスプローラーで開く
                subprocess.Popen(f'explorer "{self.path}"')
            else:
                # ファイルを実行（シェルを通じて）
                subprocess.Popen([self.path], shell=True)
        elif self.path.startswith('cmd ') or self.path.startswith('shell:'):
            # コマンドを実行（シェルを通じて）
            subprocess.Popen(self.path, shell=True)
        else:
            # 無効なパスまたはコマンドの場合、エラーを通知
            self.errorOccurred.emit("無効なパスまたはコマンドです")


class HistoryLoadingThread(QThread):
    historyLoaded = QtCore.pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)

    def run(self):
        try:
            conn = sqlite3.connect('settings.db')
            cursor = conn.cursor()
            # DISTINCT を使用して重複を除外
            cursor.execute('SELECT DISTINCT command FROM history ORDER BY ROWID DESC LIMIT 6')
            history_data = [row[0] for row in cursor.fetchall()]
            self.historyLoaded.emit(history_data)
        except Exception as e:
            print(f"履歴の読み込み中にエラーが発生しました: {e}")
        finally:
            conn.close()


# MainWindowの open_file メソッドで使用される subprocess.Popen も、
# 必要に応じてスレッドやプロセスで実行することが可能です。
# アプリケーションを実行する
def main():
    app = QtWidgets.QApplication(sys.argv)  # QApplicationのインスタンスを作成
    apply_stylesheet(app, theme='light_blue.xml')  # スタイルシートを適用
    main_win = MainWindow()  # MainWindowのインスタンスを作成
    main_win.show()  # MainWindowを表示
    sys.exit(app.exec_())  # アプリケーションの実行とイベントループの開始


if __name__ == '__main__':
    main()  # メイン関数を呼び出し、プログラムを実行
