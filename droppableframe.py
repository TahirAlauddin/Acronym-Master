from PyQt5.QtWidgets import QFrame
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QDragEnterEvent, QDropEvent

class DroppableFrame(QFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.setAcceptDrops(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        for url in urls:
            file_path = url.toLocalFile()
            self.upload_function(file_path)

    def set_upload_function(self, upload_function):
        self.upload_function = upload_function