from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot
from mainwindows import Ui_MainWindow
import sys
from generator import Generator

class MainWindowUIClass(Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.generator = Generator()
        self.fileName = None

    def setupUi(self, MW):
        """ setup the UI of the super class, and add here code
        that relates to the way we want our UI to operate. """
        super().setupUi(MW)

    def info_print(self, msg):
        self.textEdit.append(msg)

    def returnPressedSlot(self):
        self.generator.extract_specs(self.textEdit.toPlainText())
        self.generator.save_to_excel(self.textEdit.toPlainText())

    def generateSlot(self):
        data = self.generator.read_file(str(self.fileName))
        product = self.generator.get_product_name(data)
        year, month = self.generator.get_date(data)
        unique_spec = self.generator.extract_specs(data)
        self.info_print(unique_spec)
        tables = self.generator.pivot_table(data, unique_spec)
        self.generator.save_to_excel(product, year, month, tables, unique_spec)

    def browseSlot(self):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        self.fileName, _ = QtWidgets.QFileDialog.getOpenFileName(
            None, 'QFileDialog.getOpenFileName()', '',
            'All Files (*);;Excel Files (*.xlsx)',
            options = options
        )
        if self.fileName:
            self.info_print('Location: ' + self.fileName)
        

def main():
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = MainWindowUIClass()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())

main()
