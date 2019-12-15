from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot
from mainwindows import Ui_MainWindow
import sys

class MainWindowUIClass(Ui_MainWindow):
    def __init__(self):
        super().__init__()

    def setupUi(self, MW):
        """ setup the UI of the super class, and add here code
        that relates to the way we want our UI to operate. """
        super().setupUi(MW)

    def info_print(self, msg):
        self.textEdit.append(msg)

    def returnPressedSlot(self):
        pass

    def generateSlot(self):
        pass

    def browseSlot(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        fileName, _ = QtWidgets.QFileDialog.getOpenFileName(
            None, 'QFileDialog.getOpenFileName()', '',
            'All Files (*);;Excel Files (*.xlsx)',
            options = options
        )
        if fileName:
            self.info_print('Location: ' + fileName)

def main():
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = MainWindowUIClass()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

main()
