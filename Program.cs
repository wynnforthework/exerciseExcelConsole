using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Linq;
using System.Collections.Generic;

namespace FabricioEx
{
    class Program
    {
        static void Main(string[] args)
        {


            string root = Environment.CurrentDirectory;
            string outPutPath = root;
            if (args.Length == 2)
            {
                if (args[0] == "-o")
                {
                    outPutPath = args[1];
                }
            }
            root = @"E:\exerciseExcelConsole\Data";
            outPutPath = "json";
            IEnumerable<string> files = Directory.GetFiles(root + @"\", "*.xls*").Where(s => !s.StartsWith(@"~$") &&(s.EndsWith("xlsx") || s.EndsWith("xls")));
            if (files.Count<string>() > 0)
            {
                Excel.Application xlApp = new Excel.Application();
                foreach (string file in files)
                {
                    Excel2json(file, xlApp,outPutPath,root);
                }
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
        }

        static void Excel2json(string filePath, Excel.Application xlApp,string outPutPath,string root)
        {
            
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);

            Excel.Sheets sheets = xlWorkbook.Sheets;
            Excel._Worksheet xlWorksheet = sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            JArray array = new JArray();
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                JObject jObject = new JObject();
                for (int j = 1; j <= colCount; j++)
                {
                    string name = xlRange.Cells[1, j].Value2.ToString();
                    Excel.Range vCell = xlRange.Cells[i, j];
                    if (vCell.Value2 == null)
                    {
                        continue;
                    }
                    string v = vCell.Value2.ToString();
                    if (name.StartsWith("!"))
                    {
                        continue;
                    }
                    else if (name.Contains("#anchor"))
                    {
                        Action<JObject, int, string, int> AnchorIterator = null;
                        AnchorIterator = (jO, anchorIndex, colName, sheetIndex) =>
                        {
                            if (sheetIndex < sheets.Count)
                            {
                                Excel._Worksheet xl = sheets[anchorIndex];
                                bool isAnchorObj = true;
                                string type = "";
                                string join = "";
                                if (name.Contains("@#anchor"))
                                {
                                    type = "@#anchor";
                                    isAnchorObj = false;
                                    join = xl.Name.Replace(".json", "") + "/" + colName.Replace(type, "@#join");
                                }
                                else
                                {
                                    type = "#anchor";
                                    isAnchorObj = true;
                                    join = xl.Name.Replace(".json", "") + "/" + colName.Replace(type, "#join");
                                }
                                ++sheetIndex;
                                int m = sheetIndex;

                                Excel.Range range = sheets[m].UsedRange;
                                string firstCell = range.Cells[1, 1].Value2.ToString();
                                if (join.Equals(firstCell))
                                {
                                    JArray array1 = new JArray();
                                    JObject jObject1 = new JObject();
                                    for (int n = 2; n <= range.Rows.Count; n++)
                                    {
                                        if (v.Equals(range.Cells[n, 1].Value2.ToString()))
                                        {
                                            JObject jObject2 = new JObject();
                                            for (int o = 2; o <= range.Columns.Count; o++)
                                            {
                                                string colName2 = range.Cells[1, o].Value2.ToString();
                                                if (colName2.StartsWith("!"))
                                                {
                                                    continue;
                                                }
                                                else if (colName2.Contains("#anchor"))
                                                {

                                                    JObject jO2 = new JObject();
                                                    AnchorIterator(jO2, n, colName2, sheetIndex);
                                                    string colName3 = "";
                                                    if (name.Contains("@#anchor"))
                                                    {
                                                        colName3 = colName2.Replace("@#anchor", "");
                                                    }
                                                    else
                                                    {
                                                        colName3 = colName2.Replace("#anchor", "");
                                                    }
                                                    jO.Add(colName3, jO2[colName3]);
                                                }
                                                else if (colName2.Contains("@"))
                                                {
                                                    Excel.Range vCell2 = range.Cells[n, o];
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
                                                    jObject2.Add(colName2.Replace("@", ""), array2);
                                                }
                                                else
                                                {
                                                    if (isAnchorObj)
                                                    {
                                                        jObject1.Add(colName, range.Cells[n, o].Value2.ToString());
                                                        break;
                                                    }
                                                    else
                                                    {
                                                        jObject2.Add(colName2, range.Cells[n, o].Value2.ToString());
                                                    }
                                                }
                                            }
                                            array1.Add(jObject2);
                                        }
                                    }
                                    if (isAnchorObj)
                                    {
                                        if (jObject1.Count == 0)
                                        {
                                            jO.Add(colName.Replace(type, ""), jObject1);
                                        }
                                        else
                                        {
                                            jO.Add(colName.Replace(type, ""), jObject1[colName]);
                                        }
                                    }
                                    else
                                    {
                                        jO.Add(colName.Replace(type, ""), array1);
                                    }
                                }
                                else
                                {
                                    AnchorIterator(jO, anchorIndex, colName, sheetIndex);
                                }

                                Marshal.ReleaseComObject(range);
                                Marshal.ReleaseComObject(xl);
                            }
                        };
                        AnchorIterator(jObject, 1, name, 1);
                    }
                    else if (name.Contains("@"))
                    {
                        string[] items = v.Split(new char[] { ',' });
                        JArray array1 = new JArray();
                        for (var k = 0; k < items.Length; k++)
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
            string outputDir = root + @"\" + outPutPath;
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            File.WriteAllText(outputDir + @"\" + xlWorksheet.Name, array.ToString());

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

        }
    }
}
