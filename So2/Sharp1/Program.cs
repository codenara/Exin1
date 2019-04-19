// Open & Read Excel File

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

            // Read contents from Worksheet
            Worksheet exWorksheet = (Worksheet)exWorkbook.Worksheets.get_Item(1);
            Console.WriteLine(exWorksheet.Cells[1, 1].Value);
            Console.WriteLine(exWorksheet.Cells[1, 2].Value);
            Console.WriteLine(exWorksheet.Cells[2, 1].Value);
            Console.WriteLine(exWorksheet.Cells[2, 2].Value);
            Console.WriteLine(exWorksheet.get_Range("A3", "A3").Value);
            Console.WriteLine(exWorksheet.get_Range("B3", "B3").Value);

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
