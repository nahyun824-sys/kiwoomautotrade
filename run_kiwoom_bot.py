# -*- coding: utf-8 -*-
import sys
from PyQt5.QtWidgets import QApplication

from kiwoom_bot.kiwoom_client import Kiwoom


def main():
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.login()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
