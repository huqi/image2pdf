import os
import sys

from fpdf import FPDF
from PIL import Image
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtGui import QGuiApplication, QIcon, QPixmap
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox

from ui import Ui_Form

mainFrameSize = 300         # 窗体大小
A4 = (210, 297)             # 输出大小


class MyMainForm(QMainWindow, Ui_Form):
    '''
    窗体
    '''
    path = None
    i2pAction = None

    def __init__(self, parent=None):

        super(MyMainForm, self).__init__(parent)
        self.setupUi(self)

        self.img_blue = QPixmap('./src/drag-drop-blue.png')
        self.img_grey = QPixmap('./src/drag-drop-grey.png')

        self.setWindowTitle('图片转PDF')
        self.setWindowIcon(QIcon('./src/icon.png'))

        screen = QGuiApplication.primaryScreen().size()
        size = (mainFrameSize, mainFrameSize)

        x = int((screen.width() - size[0]) / 2)
        y = int((screen.height() - size[1]) / 2)
        self.setGeometry(x, y, size[0], size[1])
        self.setFixedSize(self.width(), self.height())
        self.label.setPixmap(self.img_grey)
        self.progressBar.setValue(0)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        event.accept()
        self.label.setPixmap(self.img_blue)

    def dragMoveEvent(self, event):
        pass

    def dragLeaveEvent(self, event):
        self.label.setPixmap(self.img_grey)

    def dropEvent(self, event):
        if self.path:
            QMessageBox.warning(self, "错误", "有任务正在运行")
            return

        path = event.mimeData().urls()[0].toLocalFile()

        if not os.path.isdir(path):
            QMessageBox.warning(self, "错误", "请拖入目录")
            return

        self.path = path
        self.i2pAction = I2pThread()
        self.i2pAction.setPath(self.path)
        self.i2pAction.triggerIndex.connect(self.progressProc)
        self.i2pAction.triggerInit.connect(self.progressInitProc)
        self.i2pAction.triggerDone.connect(self.progressDoneProc)
        self.i2pAction.triggerError.connect(self.progressErrorProc)
        self.i2pAction.labelTriggerIndex.connect(self.labelProc)

        self.i2pAction.start()

    def progressProc(self, i):
        self.progressBar.setValue(i)

    def progressInitProc(self, max):
        self.progressBar.setMaximum(max)

    def eventRestore(self):
        self.label.setPixmap(self.img_grey)
        self.progressBar.setValue(0)
        self.label_2.setText("(0/0)")
        self.path = None

    def progressDoneProc(self):
        replay = QMessageBox.information(self, "完成", "已完成")

        if replay == QMessageBox.StandardButton.Ok:
            os.startfile(self.path)

        self.eventRestore()

    def progressErrorProc(self, count):
        QMessageBox.warning(self, "错误", "未知错误")
        self.eventRestore()

    def labelProc(self, i, max):
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

    path = ""

    def __init__(self):
        super(I2pThread, self).__init__()

    def setPath(self, path):
        '''
        设置路径
        '''
        self.path = path

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
            imageSorted = sorted(
                imagelist, key=lambda x: int(str(x).split('.')[0]))
            for i, image in enumerate(imageSorted):
                self.triggerIndex.emit(i)
                pdf.add_page()
                filename = os.path.join(dir, image)
                w, h = get_resize(filename)
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


def scale(filename, width=None, height=None):
    '''
    缩放图片
    '''

    _width, _height = Image.open(filename).size
    wi = width / _width
    hi = height / _height
    i = min(wi, hi, 1)
    return _width*i, _height*i


def get_resize(filename):
    '''
    获取图片处理后的大小
    '''

    return scale(filename, A4[0], A4[1])


def getDirList(path):
    '''
    获取所有目录名称
    '''

    dirList = [path, ]
    for home, dirs, files in os.walk(path):
        for dirName in dirs:
            dirList.append(os.path.join(home, dirName))
    print(dirList)
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
