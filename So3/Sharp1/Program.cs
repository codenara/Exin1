// Set Cell Color

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

            // Open Workbook
            // You can use @ verbatim identifier to specify full path
            Workbook exWorkbook = exApplication.Workbooks.Open(@"c:\Work\test1.xlsx");

            // Set cell color
            Worksheet exWorksheet = (Worksheet)exWorkbook.ActiveSheet;
            exWorksheet.get_Range("A1", "B1").Interior.ColorIndex = 10;
            exWorksheet.get_Range("A2", "B3").Interior.ColorIndex = 27;

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
