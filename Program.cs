using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

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
            Excel._Worksheet xlWorksheet = sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            JArray array = new JArray();
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for(int i = 2; i <= rowCount; i++)
            {
                JObject jObject = new JObject();
                for(int j = 1; j <= colCount; j++)
                {
                    string name = xlRange.Cells[1, j].Value2.ToString();
                    string v = xlRange.Cells[i, j].Value2.ToString();
                    if (name.StartsWith("!"))
                    {
                        continue;
                    }
                    else if (name.Contains("@#anchor"))
                    {
                        string join = xlWorksheet.Name.Replace(".json","/") + name.Replace("@#anchor", "@#join");
                        for(int m = 2; m < sheets.Count; m++)
                        {
                            Excel.Range range = sheets[m].UsedRange;
                            if (join.Equals(range.Cells[1, 1].Value2.ToString()))
                            {
                                JArray array1 = new JArray();
                                for(int n = 2; n < range.Rows.Count; n++)
                                {
                                    if(v.Equals(range.Cells[n,1]))
                                    {
                                        for(int o = 2; o < range.Columns.Count; o++)
                                        {
                                            array1.Add(range.Cells[n, o]);
                                        }
                                    }                               
                                }
                                jObject.Add(name.Replace("@#anchor",""), array1);
                            }
                        }
                    }
                    else if (name.Contains("@"))
                    {
                        string[] items = v.Split(new char[] { ',' });
                        JArray array1 = new JArray();
                        for(var k = 0; k < items.Length; k++)
                        {
                            array1.Add(items[k]);
                        }
                        jObject.Add(name.Replace("@", ""), array1);
                    } 
                    else
                    {
                        jObject.Add(name, v);
                    }
                }
                array.Add(jObject);
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
