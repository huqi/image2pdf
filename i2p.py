import json
import os
import sys

from natsort import natsorted
from fpdf import FPDF
from PIL import Image
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QGuiApplication, QIcon, QPixmap
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox

from i2p_ui import Ui_Form

import ctypes

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myappid")


A3_SIZE = [297, 420]
A4_SIZE = [210, 297]
A5_SIZE = [148, 210]

FORMAT_SIZE = {'A3': A3_SIZE, 'A4': A4_SIZE, 'A5': A5_SIZE}
MAIN_FRAME_SIZE = {'WIDTH': 300, 'HEIGHT': 300}


class MyMainForm(QWidget, Ui_Form):
    '''
    窗体
    '''
    path = None
    i2pAction = None

    def __init__(self, parent=None):
        '''
        构造函数
        '''
        super(MyMainForm, self).__init__(parent)
        self.setupUi(self)

        self.img_blue = QPixmap('./src/drag-drop-blue.png')
        self.img_grey = QPixmap('./src/drag-drop-grey.png')

        self.setWindowTitle('图片转PDF')
        self.setWindowIcon(QIcon('./src/icon.png'))

        x, y = self.getCenterPos()

        self.setGeometry(
            x, y, MAIN_FRAME_SIZE['WIDTH'], MAIN_FRAME_SIZE['HEIGHT'])
        self.setFixedSize(self.width(), self.height())
        self.label.setPixmap(self.img_grey)
        self.progressBar.setValue(0)
        self.setAcceptDrops(True)

    def getCenterPos(self):
        '''
        获得居中坐标
        '''
        screen = QGuiApplication.primaryScreen().size()

        x = int((screen.width() - MAIN_FRAME_SIZE['WIDTH']) / 2)
        y = int((screen.height() - MAIN_FRAME_SIZE['HEIGHT']) / 2)

        return x, y

    def dragEnterEvent(self, event):
        '''
        拖入
        '''
        event.accept()
        self.label.setPixmap(self.img_blue)

    def dragMoveEvent(self, event):
        '''
        移动
        '''
        pass

    def dragLeaveEvent(self, event):
        '''
        移出
        '''
        self.label.setPixmap(self.img_grey)

    def dropEvent(self, event):
        '''
        松开鼠标
        '''
        if self.path:
            QMessageBox.warning(self, "错误", "有任务正在运行")
            return

        path = event.mimeData().urls()[0].toLocalFile()

        if not os.path.isdir(path):
            QMessageBox.warning(self, "错误", "请拖入目录")
            self.eventRestore()
            return

        self.path = path
        self.i2pAction = I2pThread()
        self.i2pAction.setPath(self.path)
        self.i2pAction.setFormat(self.checkRadioButton())
        self.i2pAction.setTD(self.checkTDButton())
        self.i2pAction.triggerIndex.connect(self.progressProc)
        self.i2pAction.triggerInit.connect(self.progressInitProc)
        self.i2pAction.triggerDone.connect(self.progressDoneProc)
        self.i2pAction.triggerError.connect(self.progressErrorProc)
        self.i2pAction.labelTriggerIndex.connect(self.labelProc)

        self.i2pAction.start()

    def checkRadioButton(self):
        '''
        获取纸张大小
        '''

        if self.radioButton_A3.isChecked():
            return "A3"
        elif self.radioButton_A4.isChecked():
            return "A4"
        elif self.radioButton_A5.isChecked():
            return "A5"
        else:
            assert (0)

    def checkTDButton(self):
        '''
        获取横向标识
        '''
        if self.checkBox.isChecked():
            return True
        else:
            return False

    def progressProc(self, i):
        '''
        进度条回调
        '''
        self.progressBar.setValue(i)

    def progressInitProc(self, max):
        '''
        进度条初始化
        '''
        self.progressBar.setMaximum(max)

    def eventRestore(self):
        '''
        进度条重置
        '''
        self.label.setPixmap(self.img_grey)
        self.progressBar.setValue(0)
        self.label_2.setText("(0/0)")
        self.path = None

    def progressDoneProc(self):
        '''
        完成回调
        '''
        replay = QMessageBox.information(self, "完成", "已完成")

        if replay == QMessageBox.StandardButton.Ok:
            os.startfile(self.path)

        self.eventRestore()

    def progressErrorProc(self, count):
        '''
        错误回调
        '''
        QMessageBox.warning(self, "错误", "未知错误")
        self.eventRestore()

    def labelProc(self, i, max):
        '''
        提示回调
        '''
        self.label_2.setText("({}/{})".format(i, max))


class I2pThread(QThread):
    '''
    image to pdf 线程
    '''

    triggerInit = pyqtSignal(int)
    triggerIndex = pyqtSignal(int)
    triggerError = pyqtSignal(bool)
    triggerDone = pyqtSignal()
    labelTriggerIndex = pyqtSignal(int, int)
    format = "A4"
    TD = False

    path = ""

    def __init__(self):
        '''
        构造函数
        '''
        super(I2pThread, self).__init__()

    def setFormat(self, format):
        '''
        设置纸张
        '''
        self.format = format

    def setTD(self, td):
        '''
        设置横向标识
        '''
        self.TD = td

    def setPath(self, path):
        '''
        设置路径
        '''
        self.path = path

    def get_size(self, filename):
        return Image.open(filename).size

    def get_resize(self, o, _width, _height):
        '''
        缩放图片
        '''

        width = FORMAT_SIZE[self.format][0]
        height = FORMAT_SIZE[self.format][1]

        if o == "L":
            width, height = height, width

        wi = width / _width
        hi = height / _height
        i = min(wi, hi, 1)
        return _width*i, _height*i

    def imgToPDF(self, dirs):
        '''
        图片转PDF主函数
        '''
        for i, dir in enumerate(dirs):
            imagelist = [i for i in os.listdir(dir) if is_image_file(i)]
            self.labelTriggerIndex.emit(i+1, len(dirs))

            if not imagelist:
                continue

            pdf = FPDF()
            pdf.set_auto_page_break(0)         # 自动分页设为False
            self.triggerInit.emit(len(imagelist))
            # imageSorted = sorted(
            #    imagelist, key=lambda x: str(x).split('.')[0])
            imageSorted = natsorted(imagelist)
            for i, image in enumerate(imageSorted):
                self.triggerIndex.emit(i)
                filename = os.path.join(dir, image)
                w, h = self.get_size(filename)

                o = 'L' if self.TD and w > h else ''
                w, h = self.get_resize(o, w, h)
                pdf.add_page(o)

                pdf.image(filename, x=0, y=0, w=w, h=h)

            filename = os.path.basename(dir) + ".pdf"
            pdf.output(os.path.join(dir, filename), "F")
            self.triggerIndex.emit(len(imagelist))

        self.triggerDone.emit()

    def run(self):
        '''
        线程启动函数
        '''
        try:
            dirs = getDirList(self.path)
            self.imgToPDF(dirs)
        except:
            self.triggerError.emit(True)


def is_image_file(filename):
    '''
    判断是否是图片
    '''

    return any(filename.endswith(ext) for ext in ['.png', '.jpg', 'jpeg', '.PNG', 'JPG', '.JPEG'])


def getDirList(path):
    '''
    获取所有目录名称
    '''

    dirList = [path, ]
    for home, dirs, files in os.walk(path):
        for dirName in dirs:
            dirList.append(os.path.join(home, dirName))

    return dirList


def main():
    '''
    主函数
    '''

    app = QApplication(sys.argv)
    myWin = MyMainForm()
    myWin.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
