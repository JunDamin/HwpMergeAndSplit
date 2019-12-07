
from PyQt5 import QtGui, QtCore, QtWidgets
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QBoxLayout
from PyQt5.QtCore import *
from hwpmerger import hwpMerger
import os

class Form(QWidget):
    def __init__(self):
        QWidget.__init__(self, flags=Qt.Widget)
        self.setWindowTitle("한글문서 병합하기")
        box = QBoxLayout(QBoxLayout.TopToBottom)
        self.hlb = QLabel()
        self.hxb = QLabel()
        self.hlb.setText("병합할 hwp 폴더 경로를 선택하세요")
        self.hxb.setText("실행 후 모두 허용을 누르시면 병합됩니다.")
        self.path_box = QPushButton("병합 폴더선택")
        self.quit_box = QPushButton("종료")
        self.path = ""
        self.exe = QPushButton("병합")
        self.reset = QPushButton("초기화")
        box.addWidget(self.hlb)
        box.addWidget(self.hxb)
        box.addWidget(self.path_box)
        box.addWidget(self.exe)
        box.addWidget(self.reset)
        box.addWidget(self.quit_box)
        self.setLayout(box)
        self.path_box.clicked.connect(self.get_hwp_file_folder)
        self.exe.clicked.connect(self.merging)
        self.reset.clicked.connect(self.resetting)
        self.quit_box.clicked.connect(QCoreApplication.instance().quit)
        self.file_list = []
        # Enable dragging and dropping onto the GUI
        self.setAcceptDrops(True)


    def get_hwp_file_folder(self):
        pathname = QFileDialog.getExistingDirectory(self, "Select Directory" )
        self.hlb.setText("hwp파일경로 : " + pathname)
        self.hwpFileAddress = pathname[0]
        self.path = pathname
        path_listdir = os.listdir(pathname)
        file_list = [os.path.join(pathname, file) for file in path_listdir]
        self.file_list.extend(file_list)

    def merging(self):
        hwpMerger(self.file_list)
        self.hxb.setText("병합이 완료 되었습니다.")
        
    def resetting(self):
        self.hlb.setText("병합할 hwp 폴더 경로를 선택하세요")
        self.hxb.setText("실행 후 모두 허용을 누르시면 병합됩니다.")
        self.path_box = QPushButton("양식 hwp 파일 선택")
        self.quit_box = QPushButton("종료")
        self.path = "" 
        self.file_list = []

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
               if os.path.isdir(fname):
                    print(fname)
                    os.listdir(fname)
                    self.path = fname
                    path_listdir = os.listdir(fname)
                    file_list = [os.path.join(fname, file) for file in path_listdir if file[-3:]=='hwp']
                    self.file_list.extend(file_list)
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
