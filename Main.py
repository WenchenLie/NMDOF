from sys import argv as sysargv, exit as sysexit
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication

from core.win import MyWin


if __name__ == '__main__':
    test = False
    QApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps)
    app = QApplication(sysargv)
    myshow = MyWin(test)
    myshow.show()
    sysexit(app.exec_())

