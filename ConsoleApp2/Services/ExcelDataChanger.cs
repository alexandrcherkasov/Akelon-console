using ClosedXML.Excel;
using ConsoleApp2.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp2.Services
{
    public class ExcelDataChanger
    {
        private string filePath;

        public void SetExcelFilePath(string filePath)
        {
            this.filePath = filePath;
        }

        public void UpdateContactPerson(string organizationName, string newContactPerson)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(2); // Предполагаем, что клиенты находятся на второй странице
                var range = worksheet.RangeUsed();
                bool isUpdated = false; // Флаг для отслеживания, было ли произведено обновление

                foreach (var row in range.RowsUsed().Skip(1)) // Пропускаем заголовок
                {
                    if (row.Cell(2).Value.ToString() == organizationName)
                    {
                        row.Cell(4).Value = newContactPerson;
                        isUpdated = true; // Устанавливаем флаг обновления
                    }
                }

                workbook.Save();

                if (isUpdated)
                {
                    Console.WriteLine($"Контактное лицо для организации '{organizationName}' успешно обновлено на '{newContactPerson}'.");
                }
                else
                {
                    Console.WriteLine($"Организация с названием '{organizationName}' не найдена.");
                }
            }
        }
    }
}
