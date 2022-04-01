using ClosedXML.Excel;
using System;
using System.IO;
using System.Reflection;

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
            try
            {
                string excelPath = Path.Combine(assemblyPath, "伺服器應用程式及檔案異動清單(空白).xlsx");
                using (var wbook = new XLWorkbook(excelPath))
                {
                    var sheet = wbook.Worksheet(1);

                    int rowNow = 2;
                    Console.WriteLine("開始Excel處理");
                    var row = sheet.Row(1);
                    for (int i = 0; i < files.Length; i++)
                    {
                        row = sheet.Row(rowNow);

                        string pathFileName = files[i].Replace(assemblyPath, string.Empty);
                        int idx = pathFileName.LastIndexOf("\\");
                        if (idx == 0)
                            continue;
                        string path = pathFileName.Substring(1, idx);
                        string fileName = pathFileName.Replace(path, string.Empty).Replace("\\", string.Empty);

                        row.Cell(2).Value = path;
                        row.Cell(3).Value = fileName;
                        row.Cell(5).Value = "是";

                        row.Height = 20;
                        rowNow++;
                        row.InsertRowsBelow(1);
                    }
                    row = sheet.Row(rowNow);
                    row.Delete();

                    rowNow--;
                    sheet.Range("A2", "A" + rowNow).Merge();
                    sheet.Range("G2", "G" + rowNow).Merge();

                    sheet.Columns().AdjustToContents();  // Adjust column width
                    //sheet.Rows().AdjustToContents();     // Adjust row heights
                    wbook.SaveAs(Path.Combine(assemblyPath, "伺服器應用程式及檔案異動清單.xlsx"));
                }

                Console.WriteLine("完成!!請按任意鍵繼續..");
                Console.ReadKey();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
                Console.ReadKey();
            }
        }

    }
}
