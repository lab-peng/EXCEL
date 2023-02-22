using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
//using Microsoft.Office.Interop;

namespace OpenXmlA
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //ReadWriteExcel("D:\\ProLan\\csharp\\EXCEL\\ExelFiles\\a.xlsx");
            HeaderFooter("D:\\ProLan\\csharp\\EXCEL\\ExelFiles\\b.xlsx");

        }

        public static void ReadWriteExcel(string path)
        {
            Application excel = new Application();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(path);
            if (wb == null)
            {
                Console.WriteLine("File does't exist!");
            } else
            {
                ws = wb.Worksheets[1];
                Microsoft.Office.Interop.Excel.Range cellsToRead = ws.Range["A1: B10"];
                foreach (string cell in cellsToRead.Value)
                {
                    Console.WriteLine(cell);
                }
                wb.Close();
            }

            Application xl = new Application();
            Workbook book;
            Worksheet sheet;
            book = xl.Workbooks.Open(path);
            sheet = book.Worksheets[1];
            Microsoft.Office.Interop.Excel.Range cellsToWrite = sheet.Range["B5:B5"];
            cellsToWrite.Value = "rice";

            xl.DisplayAlerts = false;
            book.SaveAs(path);
            book.Close();
        }

        public static void HeaderFooter(string path)
        {
            string imgPath = "D:\\ProLan\\csharp\\EXCEL\\ExelFiles\\imgs\\logo1.png";


            Application xl = new Application();
            Workbook book;
            Worksheet sheet;
            book = xl.Workbooks.Open(path);
            sheet = book.Worksheets[1];
            sheet.PageSetup.LeftHeader = imgPath;
            sheet.PageSetup.LeftHeader = "&G";
            //sheet.PageSetup.LeftHeader = "&I";  // for German Excel


            xl.DisplayAlerts = false;
            book.SaveAs(path);
            book.Close();
        }
            
    }
}