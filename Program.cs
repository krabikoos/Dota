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
            if (!File.Exists("C:\\Users\\Zalman\\Documents\\Код\\Dota.xlsx"))
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
                Workers.StandardWidth = 28;
                Workers.Cells[1, 1] = "Логин";
                Workers.Cells[1, 2] = "ФИО";
                Workers.Cells[1, 3] = "Пароль";
                Workers.Cells[1, 4] = "Номер Телефона";
                Workers.Cells[1, 5] = "Почта";
                Workers.Cells[1, 6] = "Дата рождения";
                Workers.Cells[1, 7] = "Возраст";
                Workers.Cells[1, 8] = "Зп";
                Workers.Cells[1, 9] = "Должность";
                People = (Excel.Worksheet)Application.Worksheets.Item[2];
                People.Name = "Список клиентов";
                People.StandardWidth = 20;
                People.Cells[1, 1] = "Логин";
                People.Cells[1, 2] = "ФИО";
                People.Cells[1, 3] = "Пароль";
                People.Cells[1, 4] = "Номер Телефона";
                People.Cells[1, 5] = "Почта";
                Items = (Excel.Worksheet)Application.Worksheets.Item[3];
                Items.Name = "Список товаров";
                Items.StandardWidth = 28; 
                Items.Cells[1, 1] = "Название";
                Items.Cells[1, 2] = "Количество";
                Items.Cells[1, 3] = "Герой";
                Items.Cells[1, 4] = "Слот";
                Items.Cells[1, 5] = "Цена";
                Items.Cells[2, 1] = "Vest of the Bloodroot Guard";
                Items.Cells[2, 2] = "141";
                Items.Cells[2, 3] = "Phantom assassin";
                Items.Cells[2, 4] = "Жилет";
                Items.Cells[2, 5] = "10.65";
                Items.Cells[3, 1] = "Pauldrons of the Battleranger";
                Items.Cells[3, 2] = "554";
                Items.Cells[3, 3] = "Wind reandger";
                Items.Cells[3, 4] = "Плечи";
                Items.Cells[3, 5] = "5.75";
                BuyList = (Excel.Worksheet)Application.Worksheets.Item[4];
                BuyList.StandardWidth = 28;
                BuyList.Name = "Список Заказов";
  

                WorkBook.SaveAs("C:\\Users\\Zalman\\Documents\\Код\\Dota.xlsx");
                Application.Quit();
            }
            
        }
    }
}
