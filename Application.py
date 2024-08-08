import sys

from FileWidgets import *
from PyQt6.QtCore import *
from PyQt6.QtWidgets import *

class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        
        # Create the file control widget
        self.fileControlWidget = FileControlWidget()
        
        # Layout setup
        layout = QGridLayout()
        layout.addWidget(self.fileControlWidget)
        
        # Central widget setup
        centralWidget = QWidget()
        centralWidget.setLayout(layout)
        self.setCentralWidget(centralWidget)
        
    def closeEvent(self, event):
        if self.fileControlWidget.office2PDFThread and self.fileControlWidget.office2PDFThread.isRunning():
            self.fileControlWidget.office2PDFThread.quit()
        
def main():
    # Create the application
    application = QApplication(sys.argv)
    application.setApplicationName("Office2PDF")
    
    # Create and configure the main window
    mainWindow = MainWindow()
    mainWindow.setObjectName("mainWindow")
    mainWindow.setWindowTitle("Office2PDF")
    mainWindow.show()
    
    # Run the application
    sys.exit(application.exec())
    
if __name__ == "__main__":
    main()