using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using TestdataGenerator.Data;
using TestdataGenerator.Library.Extentions;
using TestdataGenerator.Models;

namespace TestdataGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            var path = Path.Combine(Environment.CurrentDirectory, "Data", "Data.xlsx");
            var file = new FileInfo(path);
            using (var package = new ExcelPackage(file))
            {
                // ExcelExtentionのサンプル
                var characters = package.Workbook.Worksheets.First(_ => _.Name == "Character")
                    .ToList<Character>();

                var Items = package.Workbook.Worksheets.First(_ => _.Name == "Item")
                    .ToList<Item>();
            }

            path = Path.Combine(Environment.CurrentDirectory, "Data", "DbData.xlsx");
            var db = new TestDataContextFactory().CreateDbContext(null);

            // DbContextExtentionのサンプル
            db.ReadExcelWriteDb(path);
        }
    }
}
