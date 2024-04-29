using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using XChanger_TelegramBot.Models;

namespace XChanger_TelegramBot.Brokers
{
    internal class ExcelBroker : IDisposable
    {
        public void Dispose() { }

        public async ValueTask<List<Person>> ReadAllExternalPersonsAsync(string filePath)
        {
            int a;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var externalPersons = new List<Person>();
            FileInfo file = new FileInfo(filePath);
            int row = 2, column = 1;

            using var excelPackage =
                new ExcelPackage(file);

            ExcelWorksheet workSheet =
                excelPackage.Workbook.Worksheets[PositionID: 0];

            await excelPackage.LoadAsync(file);

            while (!IsTrailingFinalRow(row, column, workSheet))
            {
                Person externalPerson = new Person();

                externalPerson.Name = workSheet.Cells[row, column].Value.ToString();
                externalPerson.Age = workSheet.Cells[row, column + 1].Value.ToString();
                externalPerson.PetName = workSheet.Cells[row, column + 2].Value.ToString();
                externalPersons.Add(externalPerson);
                row++;
            }

            return externalPersons;

            static bool IsTrailingFinalRow(int row, int column, ExcelWorksheet workSheet) =>
                String.IsNullOrEmpty(workSheet.Cells[row, column].Value?.ToString());
        }
    }
}
