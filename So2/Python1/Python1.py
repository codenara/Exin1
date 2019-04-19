# Open & Read Excel File

# pip install pywin32
import win32com.client

# Initialize the Excel Application object
exApplication = win32com.client.Dispatch("Excel.Application")

# Set Visible property to true
exApplication.Visible = True

# Open Workbook
exWorkbook = exApplication.Workbooks.Open(r"c:\Work\test2.xlsx")

# Read contents from Worksheet
exWorksheet = exWorkbook.ActiveSheet
print(exWorksheet.Cells(1, 1).Value)
print(exWorksheet.Cells(1, 2).Value)
print(exWorksheet.Cells(2, 1).Value)
print(exWorksheet.Cells(2, 2).Value)
print(exWorksheet.Cells(3, 1).Value)
print(exWorksheet.Cells(3, 2).Value)

# Release resources
exWorkbook.Close()
exApplication.Quit()
