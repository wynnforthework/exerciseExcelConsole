using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.IO;

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
                    Excel.Range vCell = xlRange.Cells[i, j];
                    if (vCell.Value2==null)
                    {
                        continue;
                    }
                    string v = vCell.Value2.ToString();
                    if (name.StartsWith("!"))
                    {
                        continue;
                    }
                    else if (name.Contains("@#anchor"))
                    {
                        Action<JObject,int,string,int> AnchorIterator = null;
                        AnchorIterator = (jO,anchorIndex,colName, sheetIndex) =>
                        {
                            if (sheetIndex < sheets.Count)
                            {
                                Excel._Worksheet xl = sheets[anchorIndex];
                                string join = xl.Name.Replace(".json", "")+ "/" + colName.Replace("@#anchor", "@#join");
                                ++sheetIndex;
                                int m = sheetIndex;

                                Excel.Range range = sheets[m].UsedRange;
                                string firstCell = range.Cells[1, 1].Value2.ToString();
                                if (join.Equals(firstCell))
                                {
                                    JArray array1 = new JArray();
                                    for (int n = 2; n <= range.Rows.Count; n++)
                                    {
                                        if (v.Equals(range.Cells[n, 1].Value2.ToString()))
                                        {
                                            for (int o = 2; o <= range.Columns.Count; o++)
                                            {
                                                string colName2 = range.Cells[1,o].Value2.ToString();
                                                if (colName2.StartsWith("!"))
                                                {
                                                    continue;
                                                }
                                                else if (colName2.Contains("@#anchor"))
                                                {
                                                    JObject jO2 = new JObject();
                                                    AnchorIterator(jO2, n, colName2, sheetIndex);
                                                    string colName3 = colName2.Replace("@#anchor", "");
                                                    jO.Add(colName3, jO2[colName3]);
                                                }
                                                else if (colName2.Contains("@"))
                                                {
                                                    Excel.Range vCell2 = range.Cells[i, j];
                                                    if (vCell2.Value2 == null)
                                                    {
                                                        continue;
                                                    }
                                                    string v2 = vCell2.Value2.ToString();
                                                    string[] items = v2.Split(new char[] { ',' });
                                                    JArray array2 = new JArray();
                                                    for (var k = 0; k < items.Length; k++)
                                                    {
                                                        array2.Add(items[k]);
                                                    }
                                                    jO.Add(colName2.Replace("@", ""), array2);
                                                }
                                                else
                                                {
                                                    array1.Add(range.Cells[n, o].Value2.ToString());
                                                }
                                            }
                                        }
                                    }
                                    jO.Add(colName.Replace("@#anchor", ""), array1);
                                }
                                else
                                {
                                    AnchorIterator(jO, anchorIndex, colName, sheetIndex);
                                }
                            }
                        };
                        AnchorIterator(jObject,1,name,1);
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
            File.WriteAllText(@"E:\exerciseExcelConsole\Data\" + xlWorksheet.Name, array.ToString());
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
