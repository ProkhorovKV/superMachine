using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;

namespace superMachine
{
    internal class Program
    {
        static void Main(string[] args)
        {

            // Путь к Excel файлу
            string filePath = @"C:\Users\Кирилл\Desktop\Results.xlsx";
            string init = " ";
            string name = "Прохоров";

            // Открытие файла Excel
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                char[] fio = { 'п', 'р', 'о', 'х', 'о', 'р', 'о', 'в', '#', 'Λ' };
                var worksheet = package.Workbook.Worksheets[0];  

                

                for (int i = 0; i < fio.Length; i++)
                {
                    var cellValue = worksheet.Cells[2, i + 2].Text;
                    if (cellValue == "R")
                    {
                        Console.WriteLine("Попалась R, идем дальше");
                    }
                    else if (cellValue == "K,R")
                    {
                        Console.WriteLine("Попалась K,R, добавляем инициал K");
                        name += "K";
                    }
                    else if (cellValue == "B,N,!")
                    {
                        Console.WriteLine("Попалась B,N,!, добавляем инициал B");
                        name += "B";
                    }
                }

                

                // Если вам нужно значение как число, используйте:
                // var cellValue = worksheet.Cells["A1"].Value; 

               
            }



            using (var package = new ExcelPackage())
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Results");

                char[] arr = { 'п', 'р', 'о', 'х', 'в', 'к', 'и', 'л', 'а', 'д', 'м', 'Λ' };
                char[] fio = { 'п', 'р', 'о', 'х', 'о', 'р', 'о', 'в', '#', 'Λ' };
                

                sheet.Cells[2,1].Value = "q0";
                for (int i = 0; i < fio.Length; i++)
                {
                    sheet.Cells[1, i + 2].Value = fio[i];
                    sheet.Cells[2, i + 2].Value = "R";
                    
                }
                sheet.Cells[2, fio.Length].Value = "K,R";
                sheet.Cells[2, fio.Length + 1].Value = "B,N,!";

                for (int i = 0; i < name.Length; i++)
                {
                    sheet.Cells[4, i + 2].Value = name[i];
                }


                var fileInfo = new FileInfo(@"C:\Users\Кирилл\Desktop\Results.xlsx");
                package.SaveAs(fileInfo);

                Console.WriteLine("Завершено!");
            }
        }
    }
}
