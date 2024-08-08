from Converter import *
from datetime import datetime
from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from Widgets import *

class Office2PDFThread(QThread):
    updateStatus = pyqtSignal(bool, str)
    time = pyqtSignal(float)
    finished = pyqtSignal()
    
    # Initialize office2pdf thread
    def __init__(self, filePaths, outputFolder):
        super().__init__()
        self.filePaths = filePaths
        self.outputFolder = outputFolder
        
        self.word = win32com.client.DispatchEx('Word.Application')
        self.word.Visible = False
        self.word.DisplayAlerts = False
        
        self.powerpoint = win32com.client.DispatchEx('PowerPoint.Application')
        
        self.excel = win32com.client.DispatchEx('Excel.Application')
        self.excel.ScreenUpdating = False
        self.excel.DisplayAlerts = False
        self.excel.EnableEvents = False
        self.excel.Interactive = False
        self.excel.Visible = False

    # Override to run qthread to convert filepaths
    def run (self):
        for filePath in self.filePaths:
            now = datetime.now()
            success, message = office2PDF(filePath, self.outputFolder, self.word, self.powerpoint, self.excel)
            later = datetime.now()
            try:
                self.time.emit((later - now).total_seconds())
                self.updateStatus.emit(success, message)
            except:
                pass
        try:
            self.finished.emit()
            self.time.emit(0.0)
            self.updateStatus.emit(True, "Completed converting files!")
        except:
            pass

class FileControlWidget(QWidget):
    # Initialize file control widget
    def __init__(self):      
        super().__init__()
        
        # Folder widgets
        self.folderButton = QPushButton("Select Folder(s)")
        self.folderButton.clicked.connect(self.folderButtonClicked)
        folderLabel = QLabel("Folders Added:")
        self.folderScrollable = ScrollableWidget()
        self.folderScrollable.setMinimumHeight(200)
        self.folderScrollable.setMinimumWidth(400)
        self.folderItems = []
        
        # File widgets
        self.fileButton = QPushButton("Select File(s)")
        self.fileButton.clicked.connect(self.fileButtonClicked)
        fileLabel = QLabel("Files Added:")
        self.fileScrollable = ScrollableWidget()
        self.fileScrollable.setMinimumHeight(200)
        self.fileScrollable.setMinimumWidth(500)
        self.fileItems = []
        
        # Output widgets
        outputLabel = QLabel("Output Path:")
        self.outputLineEdit = QLineEdit()
        self.outputLineEdit.setMaximumHeight(24)
        self.outputButton = QPushButton("...")
        self.outputButton.clicked.connect(self.outputButtonClicked)
        
        # Convert button
        self.convertButton = QPushButton("Convert")
        self.convertButton.clicked.connect(self.convertButtonClicked)
        self.convertButton.setDisabled(True)
        
        # Time widgets
        self.timeLabel = QLabel("")
        self.messageLabel = QLabel("")
        self.messageLabel.setWordWrap(True)
        
        # Main layout setup using grid layout
        layout = QGridLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.folderButton, 0, 0, 1, 3)
        layout.addWidget(self.fileButton, 1, 0, 1, 3)
        layout.addWidget(outputLabel, 2, 0, 1, 1)
        layout.addWidget(self.convertButton, 3, 0, 1, 3)
        layout.addWidget(folderLabel, 4, 0, 1, 1)
        layout.addWidget(self.folderScrollable, 5, 0, 1, 3)
        layout.addWidget(fileLabel, 6, 0, 1, 1)
        layout.addWidget(self.fileScrollable, 7, 0, 1, 3)
        layout.addWidget(self.outputLineEdit, 2, 1, 1, 1)
        layout.addWidget(self.outputButton, 2, 2, 1, 1)
        layout.addWidget(self.timeLabel, 8, 0, 1, 1)
        layout.addWidget(self.messageLabel, 8, 1, 1, 2)  # Spanning 1 row and 2 columns
        self.setLayout(layout)
        
        self.office2PDFThread:Office2PDFThread = None
    
    # Click Function to add selected folder to the list
    def folderButtonClicked(self):
        fileDialog = QFileDialog(self)
        fileDialog.setWindowTitle('Select Folder')
        fileDialog.setFileMode(QFileDialog.FileMode.Directory)
        
        if fileDialog.exec() == QFileDialog.DialogCode.Accepted:
            folderPath = fileDialog.selectedFiles()[0]
            
            item = PathItem(folderPath)
            item.removeSignal.connect(self.removeItem)
            
            self.folderScrollable.addWidget(item)
            self.folderItems.append(item)
            
            self.toggleConvertButton()
    
    # Click Function to add selected files to the list
    def fileButtonClicked(self):
        fileDialog = QFileDialog(self)
        fileDialog.setWindowTitle('Select File(s)')
        fileDialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        fileDialog.setNameFilters([
            "Microsoft Office (*.doc *.docx *.rtf *.ppt *.pptx *.xls *.xlsx)",
            "Microsoft Word (*.doc *.docx *.rtf)",
            "Microsoft PowerPoint (*.ppt *.pptx)",
            "Microsoft Excel (*.xls *.xlsx)"
        ])
        
        if fileDialog.exec() == QFileDialog.DialogCode.Accepted:
            filePaths = fileDialog.selectedFiles()
            for filePath in filePaths:
                item = PathItem(filePath)
                item.removeSignal.connect(self.removeItem)
                
                self.fileScrollable.addWidget(item)
                self.fileItems.append(item)
                
                self.toggleConvertButton()
    
    # Click Function that sets the output path
    def outputButtonClicked(self):
        fileDialog = QFileDialog(self)
        fileDialog.setWindowTitle('Select Folder')
        fileDialog.setFileMode(QFileDialog.FileMode.Directory)
        
        if fileDialog.exec() == QFileDialog.DialogCode.Accepted:
            outputPath = fileDialog.selectedFiles()[0]
            
            self.outputLineEdit.setText(outputPath)
    
    # Click Function that gathers file paths and performs batch conversion
    def convertButtonClicked(self):
        try:
            self.disableUserInput()
            
            folderPaths = [item.path for item in self.folderItems]
            filePaths = [item.path for item in self.fileItems]
            
            for folderPath in folderPaths:
                folderFiles = folder2FileList(folderPath)
                filePaths.extend(folderFiles)
                
                for folderFile in folderFiles:
                    item = PathItem(folderFile)
                    self.fileItems.append(item)
                    self.fileScrollable.addWidget(item)
            
            self.folderScrollable.removeAllWidgets()
            self.folderItems.clear()
            
            self.office2PDFThread:Office2PDFThread = Office2PDFThread(filePaths, self.outputLineEdit.text())
            self.office2PDFThread.updateStatus.connect(self.updateStatus)
            self.office2PDFThread.time.connect(self.updateTime)
            self.office2PDFThread.finished.connect(self.convertFinished)
            self.office2PDFThread.start()
        except Exception as e:
            self.enableUserInput()
            pass
    
    # Function to enable/disable convert button based on added folders and files
    def toggleConvertButton(self):
        self.convertButton.setEnabled(bool(self.folderItems or self.fileItems))
    
    # Function to remove item from respective list and toggle convert button
    def removeItem(self, item:PathItem):
        if item in self.folderItems:
            self.folderItems.remove(item)
        if item in self.fileItems:
            self.fileItems.remove(item)
        
        self.toggleConvertButton()
    
    # Function to disable user inputs
    def disableUserInput(self):
        self.folderButton.setDisabled(True)
        self.fileButton.setDisabled(True)
        self.outputLineEdit.setDisabled(True)
        self.outputButton.setDisabled(True)
        self.convertButton.setDisabled(True)
    
    # Function to enable user inputs
    def enableUserInput(self):
        self.folderButton.setEnabled(True)
        self.fileButton.setEnabled(True)
        self.outputLineEdit.setEnabled(True)
        self.outputButton.setEnabled(True)
        self.convertButton.setEnabled(True)
    
    # Function for when done converting files to pdf
    def convertFinished(self):
        self.enableUserInput()
        
        self.fileScrollable.removeAllWidgets()
        self.folderScrollable.removeAllWidgets()
        self.fileItems.clear()
        self.folderItems.clear()
        
        self.toggleConvertButton()
    
    # Function to update time label
    def updateTime(self, time):
        formattedTime = "{:.3f}".format(time)
        self.timeLabel.setText(f"[{formattedTime}]s")
    
    # Function to update status label
    def updateStatus(self, success, message):
        self.messageLabel.setText(str(message))
        if not success:
            QMessageBox.critical(self, "Conversion Error", "An error occured during conversion:\n" + message)