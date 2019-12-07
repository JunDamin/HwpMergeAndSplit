from PyQt5 import QtGui, QtCore, QtWidgets
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QBoxLayout
from PyQt5.QtCore import *
from manager import process
import os


class Form(QWidget):
    def __init__(self):
        QWidget.__init__(self, flags=Qt.Widget)
        self.setWindowTitle("한글파일 쪼개기")
        box = QBoxLayout(QBoxLayout.TopToBottom)
        self.hlb = QLabel()
        self.hxb = QLabel()
        self.hlb.setText("구획별 분리할 hwp파일을 선택하세요")
        self.hxb.setText("실행 후 모두 허용을 누르시면 병합됩니다.")
        self.path_box = QPushButton("파일선택")
        self.quit_box = QPushButton("종료")
        self.path = ""
        self.exe = QPushButton("분리시작")
        self.reset = QPushButton("초기화")
        box.addWidget(self.hlb)
        box.addWidget(self.hxb)
        box.addWidget(self.path_box)
        box.addWidget(self.exe)
        box.addWidget(self.reset)
        box.addWidget(self.quit_box)
        self.setLayout(box)
        self.path_box.clicked.connect(self.get_hwp_file_addr)
        self.exe.clicked.connect(self.split)
        self.reset.clicked.connect(self.resetting)
        self.quit_box.clicked.connect(QCoreApplication.instance().quit)
        self.file_addr = ""
        # Enable dragging and dropping onto the GUI
        self.setAcceptDrops(True)


    def get_hwp_file_addr(self):
        pathname = QFileDialog.getOpenFileName(self, "Select Directory" )
        print(pathname)
        self.hlb.setText("hwp파일경로 : " + pathname[0])
        self.hwpFileAddress = pathname[0]
        self.path = pathname[0]
        self.file_addr = pathname[0]

    def split(self):
        process(self.file_addr)
        self.hxb.setText("분리가 완료 되었습니다.")
        
    def resetting(self):
        self.hlb.setText("구획별 분리할 hwp파일을 선택하세요")
        self.hxb.setText("실행 후 모두 허용을 누르시면 병합됩니다.")
        self.path_box = QPushButton("파일선택")
        self.quit_box = QPushButton("종료")
        self.path = ""
        self.exe = QPushButton("분리시작")
        self.reset = QPushButton("초기화")

    # The following three methods set up dragging and dropping for the app
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls:
            e.accept()
        else:
            e.ignore()

    def dragMoveEvent(self, e):
        if e.mimeData().hasUrls:
            e.accept()
        else:
            e.ignore()

    def dropEvent(self, e):
        """
        Drop files directly onto the widget
        File locations are stored in fname
        :param e:
        :return:
        """
        if e.mimeData().hasUrls:
            e.setDropAction(QtCore.Qt.CopyAction)
            e.accept()
            # Workaround for OSx dragging and dropping
            for url in e.mimeData().urls():
               fname = str(url.toLocalFile())
               if os.path.isfile(fname):
                    print(fname)
                    self.path = fname
                    self.file_addr = fname
                    self.hlb.setText("hwp파일경로 : " + fname)
               else:
                    print(fname)
                    self.file_list.append(fname)
        else:
            e.ignore()


if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    form = Form()
    form.show()
    app.exec_()
