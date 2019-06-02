# testdata-generator
Unit Test向けにテストデータを生成を支援するライブラリ

## Feature
- .NET Core 2.2
- Entity Framework Core
- EPPlus

## Project
- TestdataGenerator.Library
    - ライブラリ本体
- TestdataGenerator.Library.Test
    - ライブラリのユニットテスト
- TestdataGenerator
    - サンプル実装のプロジェクト

## Description
### **ExcelExtention (ExcelWorksheetの拡張メソッドを実装するクラス)**
- ToList<T>

  - 対象となるExcelワークシートからセルの値を取得し、対応するオブジェクトのリストを生成します。  
		ワークシート1列目の値を元に、マッピング対象のオブジェクトのプロパティをマッピングします。  
		ワークシート2列目以降の値を元に、オブジェクトの値を設定します。  
    セル内に "NULL" (文字列) を設定した場合、nullとして値を設定します。  
    対応している型は int, int?, short, short?, long, long?, decimal, decimal?, double, double?, DataTime, DataTime?, string, Enum です。  
    上記以外は System.Convert.ChangeType による変換を試みます。

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

### **DbContextExtention (DbContextの拡張メソッドを実装するクラス)**
- 共通仕様
  - ワークシート名を元に対象となるテーブル名を特定し、データの書き込みを行います。  
    ワークシート1列目の値を元に、insert文のカラム名を構築します。  
    ワークシート2列目以降の値を元に、insert文のvalues部を構築します。  
    セル内に "NULL" (文字列) を設定した場合、nullとして値を設定します。

- ReadExcelWriteDb
  - ExcelブックのデータをDBに書き込みます。  

- ReadExcelReplaceDb
  - ExcelブックのデータでDBのデータを置き換えます。(delete & insert)  
		
- ReadExcelWriteTable
  - Excelブックのデータを対象のテーブルに書き込みます。  

- ReadExcelReplaceTable
  - Excelブックのデータで対象のテーブルを置き換えます。(delete & insert)  

``` C#
            path = Path.Combine(Environment.CurrentDirectory, "Data", "DbData.xlsx");
            var db = new TestDataContextFactory().CreateDbContext(null);

            // DbContextExtentionのサンプル
            db.ReadExcelWriteDb(path);
```
