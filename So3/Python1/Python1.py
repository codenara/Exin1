# Set Cell Color

# pip install pywin32
import win32com.client

# Initialize the Excel Application object
exApplication = win32com.client.Dispatch("Excel.Application")

# Set Visible property to true
exApplication.Visible = True

# Open Workbook
exWorkbook = exApplication.Workbooks.Open(r"c:\Work\test2.xlsx")

# Set cell color
exWorksheet = exWorkbook.ActiveSheet
exWorksheet.Range("A1:B1").Interior.ColorIndex = 10
exWorksheet.Range("A2:B3").Interior.ColorIndex = 27

# Release resources
exWorkbook.Close(True) # True : save changes
exApplication.Quit()

