using System;
using System.IO;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;

namespace TreeToExcel
{
    class Program
    {
        static string assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
        static void Main(string[] args)
        {
            Console.WriteLine("開始 GO!!");
            GetFilePath();
        }

        static void GetFilePath()
        {
            //string[] folders = System.IO.Directory.GetDirectories(assemblyPath, "*", SearchOption.AllDirectories);
            string[] files = Directory.GetFiles(assemblyPath, "*.*", SearchOption.AllDirectories);
            Console.WriteLine("取得檔案清單完畢");
            ToExcel(files);
        }


        static void ToExcel(string[] files)
        {
            Excel.Application xlApp = new Excel.Application();
            try
            {
                string excelPath = Path.Combine(assemblyPath, "伺服器應用程式及檔案異動清單(空白).xlsx");
                //創建
                //Excel.Application xlApp = new Excel.Application();
                xlApp.DisplayAlerts = false;
                xlApp.Visible = false;
                xlApp.ScreenUpdating = false;

                //打開Excel
                Excel.Workbook xlsWorkBook = xlApp.Workbooks.Open(excelPath);

                //處理數據過程
                Excel.Worksheet sheet = xlsWorkBook.Worksheets[1];  //工作簿從1開始，不是0

                int rowNow = 2;
                Console.WriteLine("開始Excel處理");
                for (int i = 0; i < files.Length; i++)
                {
                    if (rowNow != 2)
                    {
                        Excel.Range range = sheet.get_Range("A2", "G2").EntireRow;
                        Excel.Range toRange = sheet.get_Range("A" + rowNow, "G" + rowNow).EntireRow;
                        toRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, range.Copy(Type.Missing));
                    }

                    string pathFileName = files[i].Replace(assemblyPath, string.Empty);
                    int idx = pathFileName.LastIndexOf("\\");
                    if (idx == 0)
                        continue;
                    string path = pathFileName.Substring(1, idx);
                    string fileName = pathFileName.Replace(path, string.Empty).Replace("\\", string.Empty);

                    sheet.Cells[rowNow, 2] = path;
                    sheet.Cells[rowNow, 3] = fileName;

                    rowNow++;
                }

                sheet.get_Range("A2", "A" + (rowNow - 1)).Merge();
                sheet.get_Range("G2", "G" + (rowNow - 1)).Merge();

                sheet.Columns.AutoFit();
                sheet.Rows.AutoFit();

                xlsWorkBook.SaveAs(Path.Combine(assemblyPath, "伺服器應用程式及檔案異動清單.xlsx"));

                xlsWorkBook.Close();
                xlApp.Quit();

                Console.WriteLine("完成!!請按任意鍵繼續..");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                xlApp.Quit();
                Console.WriteLine(e.ToString());
                Console.ReadKey();
            }
        }

    }
}
