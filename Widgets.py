from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *

class PathItem(QWidget):
    removeSignal = pyqtSignal(QWidget)
    
    # Initialize path item
    def __init__(self, path:str, elidedLength:int=400):
        super().__init__()
        
        self.path = path
        
        #Create a QLabel to display the path, elided if necessary
        fontMetrics = QFontMetrics(self.font())
        elidedText = fontMetrics.elidedText(path, Qt.TextElideMode.ElideMiddle, elidedLength)
        elidedPathLabel = QLabel(elidedText)
        
        # Create a QPushButton to remove the PathItem
        removeButton = QPushButton("X")
        removeButton.setFixedSize(20, 20)
        removeButton.clicked.connect(self.removeClicked)
        
        # Layout setup
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(elidedPathLabel)
        layout.addWidget(removeButton)
        self.setLayout(layout)
    
    # Function to remove widget from parent
    def removeClicked(self):
        # Emit signal to notify parent widget about removal
        self.removeSignal.emit(self)
        self.deleteLater()
        
class ScrollableWidget(QWidget):
    # Initialize scrollable widget
    def __init__(self):
        super().__init__()
        
        # Create layout for holding PathItems
        self.scrollLayout = QVBoxLayout()
        self.scrollLayout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.scrollLayout.setDirection(QBoxLayout.Direction.TopToBottom)
        self.scrollLayout.setContentsMargins(0, 0, 0, 0)
        self.scrollLayout.setSpacing(0)
        
        # Create a widget to hol dthe scrollable content
        self.scrollWidget = QWidget()
        self.scrollWidget.setLayout(self.scrollLayout)
        
        # Create the scroll area
        self.scrollArea = QScrollArea()
        self.scrollArea.setFrameShape(QFrame.Shape.Box)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scrollArea.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scrollArea.setWidget(self.scrollWidget)
        
        # Layout setup
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.scrollArea)
    
    # Function to add widget
    def addWidget(self, widget: QWidget):
        # Add a widget to the scroll layout
        widget.setParent(self.scrollWidget)
        self.scrollLayout.addWidget(widget)
        
        self.scrollArea.verticalScrollBar().setValue(self.scrollArea.verticalScrollBar().minimum())
        
    def removeAllWidgets(self):
        while self.scrollLayout.count() > 0:
            item = self. scrollLayout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()