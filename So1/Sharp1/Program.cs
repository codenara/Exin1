// Create & Save Excel File

using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Sharp1
{
    class Program
    {
        static void Main(string[] args)
        {
            // Initialize the Excel Application object
            Microsoft.Office.Interop.Excel.Application exApplication = new Microsoft.Office.Interop.Excel.Application();

            // Before creating new Excel Workbook, you should check whether Excel is installed in your system
            if (exApplication == null)
            {
                Console.WriteLine("Excel is not properly installed!");
                return;
            }

            // Set Visible property to true
            exApplication.Visible = true;

            // Create new Workbook
            Workbook exWorkbook = exApplication.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            // Write contents to Worksheet
            Worksheet exWorksheet = (Worksheet)exWorkbook.Worksheets.get_Item(1);
            exWorksheet.Cells[1, 1] = "ID";
            exWorksheet.Cells[1, 2] = "Name";
            exWorksheet.Cells[2, 1] = "1";
            exWorksheet.Cells[2, 2] = "One";
            exWorksheet.Cells[3, 1] = "2";
            exWorksheet.Cells[3, 2] = "Two";
            exWorksheet.Range["D1", "E2"].Value2 = "Hi";

            // Save Workbook
            // You can use @ verbatim identifier to specify full path
            // If you do not include full path, file will be saved in current folder <- this is not sure, please check. file is saved in ~\Documents\ folder
            exWorkbook.SaveAs(@"c:\Work\test1.xlsx");

            // Release resources
            exWorkbook.Close(true);
            exApplication.Quit();

            Marshal.ReleaseComObject(exWorksheet);
            Marshal.ReleaseComObject(exWorkbook);
            Marshal.ReleaseComObject(exApplication);
            exWorksheet = null;
            exWorkbook = null;
            exApplication = null;

            // Press Ctrl+F5 (or go to Debug > Start Without Debugging) to run your app.
            Console.WriteLine("Done!");
            Console.ReadKey();
        }
    }
}
