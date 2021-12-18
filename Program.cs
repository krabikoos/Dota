using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Dota
{
    class Program
    {
        static void Main(string[] args)
        {
            var Application = new Excel.Application();
            Application.SheetsInNewWorkbook = 3;
            Excel.Workbook WorkBook = Application.Workbooks.Add();
            Excel.Worksheet Sheet = (Excel.Worksheet)Application.Worksheets.Item[1];

            WorkBook.SaveAs("C:\\Users\\Zalman\\Desktop\\Код\\C#\\Dota.xls");
            WorkBook.Close();
            Application.Quit();
        }
    }
}
