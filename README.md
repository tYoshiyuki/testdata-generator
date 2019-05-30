# testdata-generator
Unit Test向けにテストデータを生成を支援するライブラリ

## Feature
- .NET Core 2.2
- Entity Framework Core
- EPPlus

## Description
- ***ExcelExtention (ExcelWorksheetの拡張メソッドを実装するクラス)***
	- ToList<T>
		- 対象となるExcelワークシートからセルの値を取得し、対応するオブジェクトのリストを生成します。  
		ワークシート1列目の値を元に、マッピング対象のオブジェクトのプロパティをマッピングします。  
		ワークシート2列目以降の値を元に、オブジェクトの値を設定します。

``` C#
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
```

- ***DbContextExtention (DbContextの拡張メソッドを実装するクラス)***
	- ReadExcelWriteDb
		- ExcelブックのデータをDBに書き込みます。  
          ワークシート名を元に対象となるテーブル名を特定し、データの書き込みを行います。

	- ReadExcelReplaceDb
        - ExcelブックのデータでDBのデータを置き換えます。(delete & insert)  
          ワークシート名を元に対象となるテーブル名を特定し、データの書き込みを行います。
		

	- ReadExcelWriteTable
        - Excelブックのデータを対象のテーブルに書き込みます。  
          ワークシート名を元に対象となるテーブル名を特定し、データの書き込みを行います。

	- ReadExcelReplaceTable
        - Excelブックのデータで対象のテーブルを置き換えます。(delete & insert)  
          ワークシート名を元に対象となるテーブル名を特定し、データの書き込みを行います。

``` C#
            path = Path.Combine(Environment.CurrentDirectory, "Data", "DbData.xlsx");
            var db = new TestDataContextFactory().CreateDbContext(null);

            // DbContextExtentionのサンプル
            db.ReadExcelWriteDb(path);
```
