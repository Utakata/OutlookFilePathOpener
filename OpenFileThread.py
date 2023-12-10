import os
import subprocess

from PyQt5.QtCore import QThread, pyqtSignal


class OpenFileThread(QThread):
    errorOccurred = pyqtSignal(str)

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
