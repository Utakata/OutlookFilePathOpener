import os
import sqlite3
import subprocess


class DatabaseHandler:
    def __init__(self, parent, load_history=True):
        self.parent = parent
        self.database_path = 'settings.db'
        self.conn = None
        self.cursor = None
        try:
            self.conn = sqlite3.connect(self.database_path)
            self.cursor = self.conn.cursor()
            self.initialize_database()
            if load_history:
                self.load_history()
        except sqlite3.Error as e:
            print(f"データベースの接続/初期化中にエラーが発生しました: {e}")

    def initialize_database(self):
        """データベースとテーブルを初期化する"""
        self.cursor.execute('CREATE TABLE IF NOT EXISTS history (command TEXT)')
        self.cursor.execute('CREATE TABLE IF NOT EXISTS shortcut (created INTEGER)')

    def close(self):
        """データベース接続を閉じる"""
        if self.conn:
            self.conn.close()

    def load_history(self):
        """履歴データをデータベースから読み込む"""
        try:
            self.cursor.execute('SELECT command FROM history ORDER BY ROWID DESC LIMIT 6')
            return [row[0] for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            print(f"履歴の読み込み中にエラーが発生しました: {e}")
            return []

    def load_shortcut_flag(self):
        """ショートカット作成フラグを読み込む"""
        try:
            self.cursor.execute('SELECT created FROM shortcut')
            result = self.cursor.fetchone()
            if result:
                self.parent.shortcutCheckbox.setChecked(bool(result[0]))
        except sqlite3.Error as e:
            print(f"ショートカットフラグの読み込み中にエラーが発生しました: {e}")

    def save_shortcut_flag(self, created):
        """ショートカット作成フラグを保存する"""
        try:
            self.cursor.execute('DELETE FROM shortcut')
            self.cursor.execute('INSERT INTO shortcut (created) VALUES (?)', (int(created),))
            self.conn.commit()
        except sqlite3.Error as e:
            print(f"ショートカットフラグの保存中にエラーが発生しました: {e}")

    def save_history(self, command):
        """コマンドを履歴に保存する"""
        try:
            self.cursor.execute('DELETE FROM history WHERE ROWID NOT IN (SELECT ROWID FROM history ORDER BY ROWID DESC LIMIT 5)')
            self.cursor.execute('INSERT INTO history (command) VALUES (?)', (command,))
            self.conn.commit()
        except sqlite3.Error as e:
            print(f"履歴の保存中にエラーが発生しました: {e}")

    def clear_history(self):
        """履歴をクリアする"""
        try:
            self.cursor.execute('DELETE FROM history')
            self.conn.commit()
        except sqlite3.Error as e:
            print(f"履歴のクリア中にエラーが発生しました: {e}")

    @staticmethod
    def get_startup_folder_path():
        """Windowsスタートアップフォルダのパスを取得する"""
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = subprocess.SW_HIDE
        try:
            return subprocess.check_output(
                ["powershell", "-Command", "echo $((New-Object -ComObject WScript.Shell).SpecialFolders('Startup'))"],
                text=True, startupinfo=startupinfo).strip()
        except subprocess.CalledProcessError as e:
            print(f"スタートアップフォルダのパス取得中にエラーが発生しました: {e}")
            return None

    def is_shortcut_created(self, shortcut_name):
        """指定されたショートカットが作成されているか確認する"""
        startup_dir = self.get_startup_folder_path()
        if startup_dir:
            return os.path.exists(os.path.join(startup_dir, f"{shortcut_name}.lnk"))
        return False
