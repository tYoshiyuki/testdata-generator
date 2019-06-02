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
                obj.ColInt.Is(10);
                obj.ColNullableInt.Is(10);
                obj.ColShort.Is((short)1);
                obj.ColNullableShort.Is((short)1);
                obj.ColLong.Is(10000000000L);
                obj.ColNullableLong.Is(10000000000L);
                obj.ColDecimal.Is((decimal) 1.1);
                obj.ColNullableDecimal.Is((decimal)1.1);
                obj.ColDouble.Is((double)11.1);
                obj.ColNullableDouble.Is((double)11.1);
                obj.ColDateTime.Is(DateTime.Parse("2000/01/01 00:00:00"));
                obj.ColNullableDateTime.Is(DateTime.Parse("2100/01/01 00:00:00"));
                obj.ColString.Is("AAA");

                obj = objects.Skip(1).First();
                obj.ColInt.Is(20);
                obj.ColNullableInt.IsNull();
                obj.ColShort.Is((short)2);
                obj.ColNullableShort.IsNull();
                obj.ColLong.Is(20000000000L);
                obj.ColNullableLong.IsNull();
                obj.ColDecimal.Is((decimal)2.2);
                obj.ColNullableDecimal.IsNull();
                obj.ColDouble.Is((double)22.2);
                obj.ColNullableDouble.IsNull();
                obj.ColDateTime.Is(DateTime.Parse("2000/01/02 00:00:00"));
                obj.ColNullableDateTime.IsNull();
                obj.ColString.Is(string.Empty);
            }
        }

        // TODO 異常系テストの実装

        public class TestObject
        {
            public int ColInt { get; set; }
            public int? ColNullableInt { get; set; }
            public short ColShort { get; set; }
            public short? ColNullableShort { get; set; }
            public long ColLong { get; set; }
            public long? ColNullableLong { get; set; }
            public decimal ColDecimal { get; set; }
            public decimal? ColNullableDecimal { get; set; }
            public double ColDouble { get; set; }
            public double? ColNullableDouble { get; set; }
            public DateTime ColDateTime { get; set; }
            public DateTime? ColNullableDateTime { get; set; }
            public string ColString { get; set; }
        }
    }
}
