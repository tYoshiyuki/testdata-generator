using ChainingAssertion;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using TestdataGenerator.Library.Extentions;
using Xunit;

namespace TestdataGenerator.Library.Test.Extentions
{
    public class ExcelExtentionTest
    {

        [Fact]
        public void ToList_正常系()
        {
            // Arrange
            var path = Path.Combine(Environment.CurrentDirectory, "Data", "TestData.xlsx");
            var file = new FileInfo(path);

            using (var package = new ExcelPackage(file))
            {
                // Act
                var objects = package.Workbook.Worksheets.First(_ => _.Name == "TestObject")
                    .ToList<TestObject>();

                // Assert
                objects.Count().Is(3);
                var obj = objects.First();
                obj.Col01.Is(1);
                obj.Col02.Is("AAA");
                obj.Col03.Is(DateTime.Parse("2000/01/01 00:00:00"));
            }
        }

        // TODO 異常系テストの実装

        public class TestObject
        {
            public int Col01 { get; set; }
            public string Col02 { get; set; }
            public DateTime Col03 { get; set; }
        }
    }
}
