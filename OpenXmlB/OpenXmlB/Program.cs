using Spire.Xls;
using Spire.Xls.Core;
using System.Security.Cryptography.Xml;

namespace OpenXmlB
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ModifyExcel();
        }

        public static void ModifyExcel()
        {
            // https://www.e-iceblue.com/Knowledgebase/Spire.XLS/Program-Guide/How-to-Edit-Excel-Data-with-Spire.XLS.html
            // https://www.e-iceblue.com/Knowledgebase/Spire.XLS/Program-Guide/How-to-Edit-Excel-Data-with-Spire.XLS.html
            Workbook book = new Workbook();
            //book.LoadFromFile(@"d:\prolan\csharp\excel\exelfiles\b.xlsx");
            //book.loadfromfile(@"d:\prolan\csharp\excel\exelfiles\b.xlsx");

            //Worksheet sheet = book.Worksheets[0];
            //sheet.Range["D2"].Text = "Kelly Cooper";
            //book.SaveToFile(@"D:\ProLan\csharp\EXCEL\ExelFiles\b.xlsx");
        }
    }
}