# Create & Save Excel File

# pip install pywin32
import win32com.client

# Initialize the Excel Application object
exApplication = win32com.client.Dispatch("Excel.Application")

# Set Visible property to true
exApplication.Visible = True

# Create new Workbook
exWorkbook = exApplication.Workbooks.Add()

# Write contents to Worksheet
exWorksheet = exWorkbook.Worksheets("Sheet1")
exWorksheet.Cells(1, 1).Value = "ID"
exWorksheet.Cells(1, 2).Value = "Name"
exWorksheet.Cells(2, 1).Value = "3"
exWorksheet.Cells(2, 2).Value = "Three"
exWorksheet.Cells(3, 1).Value = "4"
exWorksheet.Cells(3, 2).Value = "Four"

# Save Workbook
exWorkbook.SaveAs(r"c:\Work\test2.xlsx")

# Release resources
exWorkbook.Close()
exApplication.Quit()
