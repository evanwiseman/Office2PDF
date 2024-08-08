import os
import logging
import win32com.client

# Setup logger
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Constants
FORMAT_WORD_PDF = 17
FORMAT_POWERPOINT_PDF = 32
FORMAT_EXCEL_PDF = 57
OFFICE_EXTENSIONS = ('.doc', '.docx', '.rtf', '.ppt', '.pptx', '.xls', '.xlsx')
WORD_EXTENSIONS = ('.doc', '.docx', '.rtf')
PPT_EXTENSIONS = ('.ppt', '.pptx')
EXCEL_EXTENSIONS = ('.xls', '.xlsx')

# Function to validate if a path exists and is a directory
def validatePath(path):
    return bool(path and os.path.isdir(path))

# Function to validate if a file exists and has a .doc or .docx extnesion
def validateFile(file):
    return bool(file and os.path.isfile(file) and file.lower().endswith(OFFICE_EXTENSIONS))

# Function to get list of .doc and .docx files in a folder
def folder2FileList(folderPath):
    filePaths = []
    
    # Iterate over files in the folder
    for filePath in os.listdir(folderPath):
        filePath = os.path.join(folderPath, filePath)
        # Check if the file is a .doc or .docx file
        if validateFile(filePath):
            filePaths.append(filePath)
    
    return filePaths

# Function to convert file from word to pdf
def word2PDF(inputFile, outputFile, word=None):
    # Create Word Application COM object
    if word is None:
        word = win32com.client.DispatchEx('Word.Application')
        word.Visible = False
        word.DisplayAlerts = False
    
    # Open the input Word document
    document = word.Documents.Open(inputFile)
    
    # Save the document as PDF
    document.SaveAs(outputFile, FileFormat=FORMAT_WORD_PDF)

    # Close the document and quit Word application
    document.Close()

# Function to convert file from ppt to pdf
def ppt2PDF(inputFile, outputFile, powerpoint=None):
    # Create PowerPoint Application COM object
    if powerpoint is None:
        powerpoint = win32com.client.DispatchEx('PowerPoint.Application')
    
    # Open the input PowerPoint presentation
    presentation = powerpoint.Presentations.Open(inputFile, WithWindow=False)
    
    # Save the presentation as PDF
    presentation.SaveAs(outputFile, FileFormat=FORMAT_POWERPOINT_PDF)
    
    # Close the presentation and quit PowerPoint application
    presentation.Close()

# Function to convert file from excel to pdf
def excel2PDF(inputFile, outputFile, excel=None):
    # Create Excel Applicatoin COM object
    if excel is None:
        excel = win32com.client.gencache.EnsureDispatch('Excel.Application')
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.Interactive = False
        excel.Visible = False
    
    # Open the input Excel workbook
    workbook = excel.Workbooks.Open(inputFile, ReadOnly=True, UpdateLinks=False)
    
    # Save the presentation as PDF
    workbook.ExportAsFixedFormat(0, outputFile, 1, 0)
    
    # Close the workbook and quit Excel application
    workbook.Close()

# Function to convert a Word document to PDF
def office2PDF(inputFile, outputFolder, word=None, powerpoint=None, excel=None):    
    inputFile = os.path.abspath(inputFile)
    logger.info(f"Converting {inputFile} to PDF...")
    logger.debug(f"Input file: {inputFile}")
    inputFileExtension = os.path.splitext(inputFile)[1]

    # Create the output file
    outputFile = os.path.abspath(os.path.join(outputFolder, os.path.splitext(os.path.basename(inputFile))[0] + '.pdf'))
    logger.debug(f"Output file: {outputFile}")
    
    # Validate input file
    if not validateFile(inputFile):
        logger.error(f"Invalid input file path: {inputFile}")
        return False, f"Invalid input file path: {inputFile}"
    
    # Validate output path
    if not validatePath(outputFolder):
        logger.error(f"Invalid output folder path: {outputFolder}")
        return False, f"Invalid output folder path: {outputFolder}"
    
    try:
        if inputFileExtension in WORD_EXTENSIONS:
            word2PDF(inputFile, outputFile)
        elif inputFileExtension in PPT_EXTENSIONS:
            ppt2PDF(inputFile, outputFile)
        elif inputFileExtension in EXCEL_EXTENSIONS:
            excel2PDF(inputFile, outputFile)
        else:
            logger.error(f"Unsupported file format: {inputFileExtension}")
            return False, f"Unsupported file format: {inputFileExtension}"
        logger.info(f"Converted {inputFile} to PDF successfully.") 
               
        return True, f"Converted {inputFile} to PDF successfully."
    except Exception as e:
        logger.exception(f"Error converting {inputFile} to PDF: {e}")
        return False, str(e)