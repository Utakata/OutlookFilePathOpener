import os
import re

from PyQt5 import QtWidgets, QtGui, QtCore
from PyQt5.QtWidgets import QMessageBox

from DatabaseHandler import DatabaseHandler
from HistoryLoadingThread import HistoryLoadingThread
from OpenFileThread import OpenFileThread
from SettingsWindow import SettingsWindow


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.historyLoadingThread = None
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
        """データベースから履歴を読み込んでコンボボックスを更新する"""
        self.historyLoadingThread = HistoryLoadingThread()
        self.historyLoadingThread.historyLoaded.connect(self.update_combo_box)
        self.historyLoadingThread.start()
        print("履歴の読み込みスレッドを開始しました")  # デバッグ情報

    def open_file(self):
        path = self.comboBox.currentText()
        path = re.sub(r'[\r\n\"\'<>\[\]]', '', path)  # 改行と引用符を削除

        if path.startswith("file:///"):
            # file:/// を削除し、バックスラッシュに変更
            path = path[8:].replace('/', '\\')

        # osモジュールを使用してファイルの存在を確認
        if not os.path.exists(path):
            self.show_error_message("指定されたパスは存在しません。")
            return

        # OpenFileThread スレッドを起動してファイルを開く
        self.databaseHandler.save_history(path)  # 入力されたパスを履歴に保存

        # コンボボックスのテキストから file:/// を削除
        self.comboBox.setEditText(path)

        self.fileOpeningThread = OpenFileThread(path)
        self.fileOpeningThread.errorOccurred.connect(self.show_error_message)
        self.fileOpeningThread.start()

    def show_error_message(self, message):
        QMessageBox.warning(self, "エラー", message)

    # コンボボックスのクリアボタンがクリックされたときの処理
    def clear_combo_box(self):
        self.comboBox.clearEditText()  # コンボボックスのテキストをクリアする

    # 履歴をクリアする機能
    def clear_histories(self):
        # 履歴をクリアする処理を実装
        pass

    # 設定ウィンドウを開く機能
    def open_settings(self):

        # メインウインドウを暗くする
        self.setWindowOpacity(0.5)  # 透明度を下げる
        settings_window = SettingsWindow(self)
        settings_window.setWindowModality(QtCore.Qt.ApplicationModal)  # モーダルダイアログとして設定

        # メインウインドウの位置とサイズを取得
        main_window_geometry = self.geometry()
        settings_window_width = settings_window.frameGeometry().width()
        settings_window_height = settings_window.frameGeometry().height()

        # 設定ウインドウをメインウインドウの近くに配置（整数に変換）
        settings_window_x = int(main_window_geometry.x() + (main_window_geometry.width() - settings_window_width) / 2)
        settings_window_y = int(main_window_geometry.y() + (main_window_geometry.height() - settings_window_height) / 2)

        settings_window.setGeometry(settings_window_x, settings_window_y, settings_window_width, settings_window_height)

        settings_window.exec_()  # モーダルダイアログを実行

        # ダイアログが閉じられたら、ウィンドウの透明度を元に戻す
        self.setWindowOpacity(1.0)

    def update_combo_box(self, items):
        """コンボボックスを履歴データで更新する"""
        self.comboBox.clear()  # 既存のアイテムをクリア
        if items is not None:
            for item in items:
                self.comboBox.addItem(item)
            print(f"コンボボックスを更新しました: {items}")  # デバッグ情報
        else:
            print("履歴データが空です")  # デバッグ情報

    def on_history_loading_finished(self):
        # スレッドが終了したら履歴データを取得してコンボボックスを更新
        history_data = self.historyLoadingThread.get_history_data()  # スレッドから履歴データを取得するメソッド
        self.update_combo_box(history_data)

    def closeEvent(self, event):
        if self.historyLoadingThread.isRunning():
            self.historyLoadingThread.terminate()
        event.accept()
