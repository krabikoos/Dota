using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Dota
{
    class Program
    {
        static void Main(string[] args)
        {
            //Создаёт файл Excel просто с названиями листов
            if (!File.Exists("C:\\Users\\Zalman\\Desktop\\Код\\C#\\Dota\\Dota.xlsx"))
            {
                var Application = new Excel.Application();
                Application.SheetsInNewWorkbook = 4;
                Excel.Workbook WorkBook = Application.Workbooks.Add();

                Excel.Worksheet Workers;
                Excel.Worksheet People;
                Excel.Worksheet Items;
                Excel.Worksheet BuyList;

                Workers = (Excel.Worksheet)Application.Worksheets.Item[1];
                Workers.Name = "Список Сатрудников";
                People = (Excel.Worksheet)Application.Worksheets.Item[2];
                People.Name = "Список клиентов";
                Items = (Excel.Worksheet)Application.Worksheets.Item[3];
                Items.Name = "Список товаров";
                BuyList = (Excel.Worksheet)Application.Worksheets.Item[4
                    ];
                BuyList.Name = "Список Заказов";
                WorkBook.SaveAs("C:\\Users\\Zalman\\Desktop\\Код\\C#\\Dota\\Dota.xlsx");
                Application.Quit();
            }
            
        }
    }
}
