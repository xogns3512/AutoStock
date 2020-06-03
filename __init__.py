
from Kiwoom.kiwoom import *
import sys
from PyQt5.QtWidgets import *

class Main():
    def __init__(self):
        print("Main() start")
        self.app = QApplication(sys.argv) #PyQt5 로 실행할 파일명을 자동 설정
        self.kiwoom = Kiwoom()
        self.app.exec_()

if __name__ == "__main__":
    Main()

