from PyQt5 import QtGui, QtWidgets, QtCore
from home import Ui_MainWindow
from docacronym_master import DocAcronymMaster
from abbreviation_detector import find_abbreviations
from utils import get_users_desktop_folder
import os
import ctypes

myappid = 'tahiralauddin.acronym-master.1.0.0' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

class MyMainWindow(QtWidgets.QMainWindow):
    documentProgressSignal = QtCore.pyqtSignal(int)
    
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowIcon(QtGui.QIcon("images/logo-icon-transparent.ico"))
        self.setWindowTitle("Acronym Master")
        self.Ui_Components()


    def Ui_Components(self):
        self.ui.uploadDocumentFrame.set_upload_function(self.processDocument)
        self.ui.uploadDocumentFrame.mousePressEvent = self.uploadDocument
        self.ui.progressBar.setValue(0)
        self.ui.stackedWidget.setCurrentIndex(0)
        self.ui.downloadButton.clicked.connect(self.downloadDocument)
        self.ui.helpButton.clicked.connect(self.help)
        self.documentProgressSignal.connect(self.updateProgress)
        self.setMinimumSize(800, 500) # set the minimum size
        self.setWhiteTheme()

    def setWhiteTheme(self):
        self.ui.contentFrame.setStyleSheet("")
        self.ui.headerFrame.setStyleSheet("#headerFrame {background: rgba(0,0,0, 255)}")
        self.ui.uploadDocumentFrame.setStyleSheet("#uploadDocumentFrame {background: rgba(0,0,0,200); color: black;}")
        self.ui.downloadFrame.setStyleSheet("#downloadFrame {border-radius: 6px;background: rgb(0, 0, 0);}")
        self.ui.progressBar.setStyleSheet(self.ui.progressBar.styleSheet()+"QProgressBar {border-radius: 2px; background-color: rgba(0, 0, 0, 175); text-align: center}" )

    def resizeEvent(self, a0) -> None:
        self.ui.uploadDocumentFrame.setMinimumSize(QtCore.QSize(a0.size().width() // 3, a0.size().height() // 3))
        return super().resizeEvent(a0)

    def uploadDocument(self, *args):
        file, _ = QtWidgets.QFileDialog.getOpenFileName(self, 'Upload File', '.', 'MS Documents (*.docx)')

        if file:
            self.processDocument(file)

    def processDocument(self, file, *args):
        self.docMaster = DocAcronymMaster(file)
        # Emit signal
        self.documentProgressSignal.emit(10)
    
        # get the text of the document
        text = self.docMaster.get_text()

        # Emit signal
        self.documentProgressSignal.emit(20)
    
        # get the abbreviations in the text
        abbreviations = find_abbreviations(text, self.documentProgressSignal)
        # update the document with the table of abbreviations
        fullpath, filename = os.path.split(file)
        self.filepath = os.path.join(fullpath, f'{os.path.splitext(filename)[0]}-updated.docx')
        # Emit signal
        self.documentProgressSignal.emit(90)
    
        try:
            self.docMaster.update_document(abbreviations, self.filepath)
        except PermissionError:
            self.filepath = os.path.join(get_users_desktop_folder(), f'{filename}-updated.docx')
            self.docMaster.update_document(abbreviations, self.filepath)

        # Emit signal
        self.documentProgressSignal.emit(100)

        self.ui.stackedWidget.setCurrentIndex(1)

        # Set Download Page information
        self.ui.downloadButtonInformation.setText(filename)
    

    def addProgressBar(self):

        self.progressBar = QtWidgets.QProgressBar(self.ui.frame_4)
        self.progressBar.setMinimumSize(QtCore.QSize(0, 20))
        self.progressBar.setStyleSheet("QProgressBar {\n"
    "    border-radius: 2px;\n"
    "    background-color: rgba(255, 255, 255, 75);\n"
    "    text-align: center;\n"
    "}\n"
    "\n"
    "QProgressBar::chunk {\n"
    "    background-color: #F6751E;\n"
    "    border-radius: 2px;\n"
    "}\n"
    "")
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout_10.addWidget(self.progressBar)
    
    def updateProgress(self, value):
        self.ui.progressBar.setValue(value)

    def downloadDocument(self):
        self.docMaster.saveDocument(self.filepath)
        QtWidgets.QMessageBox.information(self, "File Downloaded", f"The document is saved as {os.path.basename(self.filepath)} successfully!")
        self.ui.progressBar.setValue(0)
        self.ui.stackedWidget.setCurrentIndex(0)

    def help(self):
        pass

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = MyMainWindow()
    mainWindow.show()
    sys.exit(app.exec_())
