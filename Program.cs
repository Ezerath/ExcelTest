using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelTest
{
    class Program
    {
        static void Main(string[] args)
        {
            var bankAccounts = new List<Account>
            {
                new Account {
                              ID = 345678,
                              Balance = 541.27
                            },
                new Account {
                              ID = 1230221,
                              Balance = -127.44
                            }
            };
            StartExcel(bankAccounts);

        }


        private static void StartExcel(IEnumerable<Account> accounts)
        {
            var excelApp = new Excel.Application();
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet worksheet = excelApp.ActiveSheet;
            worksheet.Cells[1, "A"] = "ID Number";
            worksheet.Cells[1, "B"] = "Current Balance";
            worksheet.Columns.AutoFit();
            var row = 1;
            foreach (var item in accounts)
            {
                row++;
                worksheet.Cells[row, "A"] = item.ID;
                worksheet.Cells[row, "B"] = item.Balance;
            }
        }
    }
}
