using ChainingAssertion;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using TestdataGenerator.Library.Extentions;
using Xunit;

namespace TestdataGenerator.Library.Test.Extentions
{
    public class ExcelExtentionTest
    {
        private readonly string _rootPath = Path.Combine(Environment.CurrentDirectory, "Data", "ExcelExtentionTest");

        [Fact]
        public void ToList_正常系()
        {
            // Arrange
            var path = Path.Combine(_rootPath, "TestData.xlsx");
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
                obj.ColDecimal.Is((decimal)1.1);
                obj.ColNullableDecimal.Is((decimal)1.1);
                obj.ColDouble.Is(11.1);
                obj.ColNullableDouble.Is(11.1);
                obj.ColDateTime.Is(DateTime.Parse("2000/01/01 00:00:00"));
                obj.ColNullableDateTime.Is(DateTime.Parse("2100/01/01 00:00:00"));
                obj.ColString.Is("AAA");
                obj.ColEnum.Is(TestEnum.Eval1);

                obj = objects.Skip(1).First();
                obj.ColInt.Is(20);
                obj.ColNullableInt.IsNull();
                obj.ColShort.Is((short)2);
                obj.ColNullableShort.IsNull();
                obj.ColLong.Is(20000000000L);
                obj.ColNullableLong.IsNull();
                obj.ColDecimal.Is((decimal)2.2);
                obj.ColNullableDecimal.IsNull();
                obj.ColDouble.Is(22.2);
                obj.ColNullableDouble.IsNull();
                obj.ColDateTime.Is(DateTime.Parse("2000/01/02 00:00:00"));
                obj.ColNullableDateTime.IsNull();
                obj.ColString.Is(string.Empty);
                obj.ColEnum.Is(TestEnum.Eval2);
            }
        }

        [Fact]
        public void ToList_正常系_map条件指定有り()
        {
            // Arrange
            var path = Path.Combine(_rootPath, "TestData.xlsx");
            var file = new FileInfo(path);
            var map = new Dictionary<string, string>() { { "ColInt", "ColInt" } };

            using (var package = new ExcelPackage(file))
            {
                // Act
                var objects = package.Workbook.Worksheets.First(_ => _.Name == "TestObject")
                    .ToList<TestObject>(map);

                // Assert
                objects.Count().Is(3);
                var obj = objects.First();
                obj.ColInt.Is(10);
                obj.ColNullableInt.IsNull();

                obj = objects.Skip(1).First();
                obj.ColInt.Is(20);
                obj.ColNullableInt.IsNull();
            }

        }

        [Fact]
        public void ToList_正常系_データ無し()
        {
            // Arrange
            var path = Path.Combine(_rootPath, "TestData-データ無し.xlsx");
            var file = new FileInfo(path);

            using (var package = new ExcelPackage(file))
            {
                // Act
                var objects = package.Workbook.Worksheets.First(_ => _.Name == "TestObject")
                    .ToList<TestObject>();

                // Assert
                objects.Count().Is(0);
            }
        }

        [Fact]
        public void ToList_異常系_ヘッダ無し()
        {
            // Arrange
            var path = Path.Combine(_rootPath, "TestData-ヘッダ無し.xlsx");
            var file = new FileInfo(path);

            using (var package = new ExcelPackage(file))
            {
                // Act・Assert
                var ex = Assert.Throws<ArgumentException>(() => package.Workbook.Worksheets.First(_ => _.Name == "TestObject")
                    .ToList<TestObject>());
                ex.Message.Is("ワークシート1列目にマッピング対象のオブジェクトのプロパティ名を設定してください。[ワークシート名:TestObject]");
            }
        }

        [Fact]
        public void ToList_異常系_ヘッダ重複()
        {
            // Arrange
            var path = Path.Combine(_rootPath, "TestData-ヘッダ重複.xlsx");
            var file = new FileInfo(path);

            using (var package = new ExcelPackage(file))
            {
                // Act・Assert
                var ex = Assert.Throws<ArgumentException>(() => package.Workbook.Worksheets.First(_ => _.Name == "TestObject")
                    .ToList<TestObject>());
                ex.Message.Is("ワークシート1列目に重複しているプロパティ名が存在します。[プロパティ名:Col01,Col02]");
            }
        }

        [Fact]
        public void ToJson_正常系()
        {
            // Arrange
            var path = Path.Combine(_rootPath, "TestData.xlsx");
            var file = new FileInfo(path);

            using (var package = new ExcelPackage(file))
            {
                // Act
                var json = package.Workbook.Worksheets.First(_ => _.Name == "TestObject")
                    .ToJson<TestObject>();

                // Assert
                json.IsNotEmpty();
            }
        }

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
            public TestEnum ColEnum { get; set; }
        }

        public enum TestEnum
        {
            Eval1 = 1, Eval2, Eval3
        }
    }
}
