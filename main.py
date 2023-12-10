import sys

from PyQt5 import QtWidgets
from qt_material import apply_stylesheet

from MainWindow import MainWindow


# アプリケーションを実行する
def main():
    app = QtWidgets.QApplication(sys.argv)  # QApplicationのインスタンスを作成
    apply_stylesheet(app, theme='light_blue.xml')  # スタイルシートを適用
    main_win = MainWindow()  # MainWindowのインスタンスを作成
    main_win.show()  # MainWindowを表示
    sys.exit(app.exec_())  # アプリケーションの実行とイベントループの開始

if __name__ == '__main__':
    main()
