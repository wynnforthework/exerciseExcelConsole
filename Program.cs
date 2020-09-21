using System;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading;
using System.Diagnostics;
using CU;

namespace FabricioEx
{
    class Program
    {
        static string root = Environment.CurrentDirectory;
        static string outPutPath = "";
        static bool isAllowBlank = false;
        static bool needArg = false;
        static bool needWatch = false;
        static FileSystemWatcher Watch;
        static int maxFileNameLenght = 0;
        static bool isDebug = false;
        static void Main(string[] args)
        {
            ParseArgs(args);
        }

        static void ParseArgs(string[] args)
        {
            if (args.Length > 0)
            {
                if (args[0].Equals("d"))
                {
                    return;
                }
                for (int i = 0; i < args.Length; i++)
                {
                    if (needArg)
                    {
                        needArg = false;
                        continue;
                    }
                    switch (args[i])
                    {
                        case "-o":
                            // 输出目录
                            outPutPath = args[i + 1];
                            needArg = true;
                            break;
                        case "-h":
                            // 帮助
                            ShowHelpDoc();
                            break;
                        case "-b":
                            // 允许空白
                            if (args[i + 1].ToLower() == "y")
                            {
                                isAllowBlank = true;
                            }
                            needArg = true;
                            break;
                        case "-w":
                            // 监听当前目录是否有文件被修改、新增、删除等
                            if (args[i + 1].ToLower() == "y")
                            {
                                needWatch = true;
                            }
                            needArg = true;
                            break;
                        default:
                            if (args[i].StartsWith("-"))
                            {
                                Console.WriteLine("unknown option: "+args[i]);
                                Console.WriteLine("usage: [-h] [-o] [-b]");
                            }
                            else
                            {
                                Console.WriteLine(args[i] + " is not a git command. See 'git -h'.");
                            }
                            break;
                    }
                }
                if (needWatch)
                {

                    Thread thread = new Thread(WatchRootFilesChanged);
                    //thread.IsBackground = true;
                    thread.Start();
                }
                ParseExcel();
            }
            else if (args.Length > 0)
            {
                Console.WriteLine("Wrong number of parameters, please enter - h to view the help document.");
                Console.WriteLine("Please enter this option or d quit.");
                string str = Console.ReadLine();
                ParseArgs(str.Split(' '));
            }
            else
            {
                if (isDebug)
                {
                    root = @"E:\exerciseExcelConsole\Data";
                    outPutPath = "json";
                    isAllowBlank = true;
                    needWatch = false;

                    if (needWatch)
                    {
                        Thread thread = new Thread(WatchRootFilesChanged);
                        //thread.IsBackground = true;
                        thread.Start();
                        Console.WriteLine("开始当前目录下文件的改变。");
                    }
                    ParseExcel();
                }
                else
                {
                    ShowHelpDoc();
                }

            }

        }
        static void ShowHelpDoc()
        {
            Console.WriteLine("Welcome to the excel to json tools");
            Console.WriteLine("-h View help documentation.");
            Console.WriteLine("-o Set the output directory, if not set, the default is the current folder.");
            Console.WriteLine("-b Whether cells are allowed to be empty (non first column and first row),The parameter is Y or N.");
            Console.WriteLine("Please enter this option or d quit.");
            string str = Console.ReadLine();
            ParseArgs(str.Split(' '));
        }
        static void ParseExcel()
        {
            IEnumerable<string> allFiles = Directory.GetFiles(root + @"\", "*.xls*").Where(s => s.EndsWith("xlsx") || s.EndsWith("xls"));
            List<string> files = new List<string>();
            if (allFiles.Count<string>() > 0)
            {
                Excel.Application xlApp = new Excel.Application();
                foreach (string file in allFiles)
                {
                    if ((new FileInfo(file).Attributes & FileAttributes.Hidden) != FileAttributes.Hidden)
                    {
                        var fileName = Path.GetFileName(file);
                        if (fileName.Length > maxFileNameLenght)
                        {
                            maxFileNameLenght = fileName.Length;
                        }
                        files.Add(file);
                    }
                }
                Console.WriteLine("Found " + files.Count + " xlsx or xls files.");
                Console.WriteLine("convert succeed:");
                int progress = 1;
                bool right = true;
                foreach (string file in files)
                {
                    right = Excel2json(file, xlApp, outPutPath, root, progress + @"/" + files.Count + "\t", isAllowBlank);
                    if (right)
                    {
                        progress++;
                    } 
                    else
                    {
                        break;
                    }
                }
                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                if (right)
                {
                    Console.WriteLine("All the files have been converted.");
                }
                else
                {
                    Console.WriteLine("Something went wrong.");
                    Console.Write("按任意键退出...");
                    Console.ReadKey(true);
                }
            }
            else
            {
                Console.WriteLine("There are no Excel files in this directory.");
            }
        }

        static bool Excel2json(string filePath, Excel.Application xlApp, string outPutPath, string root, string progress, bool isAllowBlank)
        {

            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filePath);

            Excel.Sheets sheets = xlWorkbook.Sheets;
            Excel._Worksheet xlWorksheet = sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            bool isRight = true;
            try 
            {
                var fileName = Path.GetFileName(filePath);
                var tCount = Math.Ceiling((double)maxFileNameLenght / 8);
                Console.Write(progress + Path.GetFileName(filePath));
                for (int i = 1; i <= (tCount + 1- Math.Ceiling((double)fileName.Length / 8)); i++)
                {
                    Console.Write("\t");
                }
                
                JArray array = new JArray();
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int totalCount = rowCount * colCount;

                ConsoleUtility.WriteProgressBar(0);
                for (int i = 2; i <= rowCount; i++)
                {
                    JObject jObject = new JObject();
                    for (int j = 1; j <= colCount; j++)
                    {
                        ConsoleUtility.WriteProgressBar((int)(((i - 1) * colCount + j)*100) / totalCount,true);
                        var ce1 = xlRange.Cells[1, j];
                        if ((string)ce1.Text == "")
                        {
                            throw new Exception("第一行第一列不允许为空");
                        }
                        string name = ce1.Value2.ToString();
                        Excel.Range vCell = xlRange.Cells[i, j];

                        string v;
                        if (vCell.Text=="")
                        {
                            if (isAllowBlank)
                            {
                                v = "";
                            }
                            else
                            {
                                throw new Exception("当前设置不允许单元格为空");
                            }
                        }
                        else
                        {
                            v = vCell.Value2.ToString();
                        }
                        if (v.Equals("")) 
                        { 
                            jObject.Add(name, v);
                        } 
                        else if (name.StartsWith("!"))
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
                                    bool isAnchorObj;
                                    string type;
                                    string join;
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
                                    ce1 = range.Cells[1, 1];
                                    if ((string)ce1.Text == "")
                                    {
                                        throw new Exception("第一行第一列不允许为空");
                                    }
                                    string firstCell = ce1.Value2.ToString();
                                    if (join.Equals(firstCell))
                                    {
                                        JArray array1 = new JArray();
                                        JObject jObject1 = new JObject();
                                        for (int n = 2; n <= range.Rows.Count; n++)
                                        {
                                            ce1 = range.Cells[n, 1];
                                            if ((string)ce1.Text == "")
                                            {
                                                throw new Exception("第一行第一列不允许为空");
                                            }
                                            if (v.Equals(ce1.Value2.ToString()))
                                            {
                                                JObject jObject2 = new JObject();
                                                for (int o = 2; o <= range.Columns.Count; o++)
                                                {
                                                    ce1 = range.Cells[1, o];
                                                    if ((string)ce1.Text == "")
                                                    {
                                                        throw new Exception("第一行第一列不允许为空");
                                                    }
                                                    string colName2 = ce1.Value2.ToString();
                                                    if (colName2.StartsWith("!"))
                                                    {
                                                        continue;
                                                    }
                                                    else if (colName2.Contains("#anchor"))
                                                    {

                                                        JObject jO2 = new JObject();
                                                        AnchorIterator(jO2, n, colName2, sheetIndex);
                                                        string colName3;
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
                                                        JArray array2 = new JArray();
                                                        if (vCell2.Text == "")
                                                        {
                                                            if (isAllowBlank)
                                                            {
                                                                jObject2.Add(colName2.Replace("@", ""), array2);
                                                            }
                                                            else
                                                            {
                                                                throw new Exception("当前设置不允许单元格为空");
                                                            }
                                                        }
                                                        else
                                                        {
                                                            string v2 = vCell2.Value2.ToString();
                                                            string[] items = v2.Split(new char[] { ',' });
                                                            for (var k = 0; k < items.Length; k++)
                                                            {
                                                                array2.Add(items[k]);
                                                            }
                                                            jObject2.Add(colName2.Replace("@", ""), array2);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (isAnchorObj)
                                                        {
                                                            Excel.Range vCell2 = range.Cells[n, o];
                                                            if (vCell2.Text == "")
                                                            {
                                                                if (isAllowBlank)
                                                                {
                                                                    jObject1.Add(colName, "");
                                                                }
                                                                else
                                                                {
                                                                    throw new Exception("当前设置不允许单元格为空");
                                                                }
                                                            }
                                                            else
                                                            {
                                                                jObject1.Add(colName, vCell2.Value2.ToString());
                                                            }
                                                            break;
                                                        }
                                                        else
                                                        {
                                                            Excel.Range vCell2 = range.Cells[n, o];
                                                            if (vCell2.Text == "")
                                                            {
                                                                if (isAllowBlank)
                                                                {
                                                                    jObject2.Add(colName2, "");
                                                                }
                                                                else
                                                                {
                                                                    throw new Exception("当前设置不允许单元格为空");
                                                                }
                                                            }
                                                            else
                                                            {
                                                                jObject2.Add(colName2, vCell2.Value2.ToString());
                                                            }
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
                string outPutDir;
                if (Path.IsPathRooted(outPutPath))
                {
                    outPutDir = outPutPath;
                }
                else
                {
                    outPutDir = root + @"\" + outPutPath;
                }
                if (!Directory.Exists(outPutDir))
                {
                    Directory.CreateDirectory(outPutDir);
                }
                File.WriteAllText(outPutDir + @"\" + xlWorksheet.Name, array.ToString());
            }
            catch (Exception e) 
            {
                isRight = false;
                Console.WriteLine(e.ToString());
                Debug.Write(e.ToString());
            }
            finally 
            { 
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
            };
            return isRight;
        }

        static void WatchRootFilesChanged()
        {
            Watch = new FileSystemWatcher(root);
            Watch.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.DirectoryName;
            Watch.IncludeSubdirectories = false;
            Watch.Changed += new FileSystemEventHandler(Watch_changed);
            Watch.Created += new FileSystemEventHandler(Watch_created);
            Watch.Deleted += new FileSystemEventHandler(Watch_deleted);
            Watch.IncludeSubdirectories = false;
            Watch.EnableRaisingEvents = true;

            while (true)
            {
                Thread.Sleep(1000);
            }
        }
        static void Watch_changed(object source, FileSystemEventArgs e)
        {
            if (Watch != null)
            {
                try
                {
                    if (!e.Name.StartsWith("~$"))
                    {
                        Console.WriteLine("有文件改动");
                        ParseExcel();
                    }
                    Watch.EnableRaisingEvents = false;
                }
                finally
                {
                    Watch.EnableRaisingEvents = true;
                }
            }
        }
        static void Watch_created(object source, FileSystemEventArgs e)
        {
            if (Watch != null)
            {
                try
                {
                    if (!e.Name.StartsWith("~$"))
                    {
                        Console.WriteLine("有文件被创建");
                    }
                    Watch.EnableRaisingEvents = false;
                }
                finally
                {
                    Watch.EnableRaisingEvents = true;
                }
            }
        }
        static void Watch_deleted(object source, FileSystemEventArgs e)
        {
            if (Watch != null)
            {
                try
                {
                    if (!e.Name.StartsWith("~$"))
                    {
                        Console.WriteLine("有文件删除");
                        ParseExcel();
                    }
                    Watch.EnableRaisingEvents = false;
                }
                finally
                {
                    Watch.EnableRaisingEvents = true;
                }
            }
        }
    }
}
