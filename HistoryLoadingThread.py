import sqlite3
from sqlite3 import Connection

from PyQt5.QtCore import QThread, pyqtSignal


class HistoryLoadingThread(QThread):
    historyLoaded = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.history_data = None

    def run(self):
        try:
            # ここで conn をローカル変数として定義
            conn: Connection = sqlite3.connect('settings.db')
            cursor = conn.cursor()
            # DISTINCT を使用して重複を除外
            cursor.execute('SELECT DISTINCT command FROM history ORDER BY ROWID DESC LIMIT 6')
            history_data = [row[0] for row in cursor.fetchall()]
            self.historyLoaded.emit(history_data)
        except Exception as e:
            print(f"履歴の読み込み中にエラーが発生しました: {e}")
        finally:
            # ここで conn を閉じる
            if conn:
                conn.close()

    def get_history_data(self):
        # 履歴データを取得して返す
        return self.history_data
