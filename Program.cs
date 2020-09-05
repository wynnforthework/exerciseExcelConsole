using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace FabricioEx
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application xlApp = new Excel.Application();
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\samue\source\repos\FabricioEx\Data\arquivo.xlsx");
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"E:\exerciseExcelConsole\Data\Guide.xlsx");

            Excel.Sheets sheets = xlWorkbook.Sheets;
            int sheetCount = sheets.Count;

            for(int i = 1; i <= sheetCount; i++)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[i];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for(int j = 1; j <= rowCount; j++)
                {
                    for(int k = 1; k <= colCount; k++)
                    {
                        string v = xlRange.Cells[j, k].Value2.ToString();
                        Console.Write(v + "\t");
                    }
                    Console.Write("\r\n");
                }

                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
        }
    }
}
