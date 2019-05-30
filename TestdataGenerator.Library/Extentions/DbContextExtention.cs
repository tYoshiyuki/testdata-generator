using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

namespace TestdataGenerator.Library.Extentions
{
    public static class DbContextExtention
    {
        /// <summary>
        /// ExcelブックのデータをDBに書き込みます
        /// ワークシート名を元に対象となるテーブル名を特定し、データの書き込みを行います
        /// </summary>
        /// <param name="dbContext"></param>
        /// <param name="path">Excelファイルのパス</param>
        public static void ReadExcelWriteDb(this DbContext dbContext, string path)
        {
            var file = new FileInfo(path);
            using (var package = new ExcelPackage(file))
            {
                foreach (var sheet in package.Workbook.Worksheets)
                {
                    WriteTable(dbContext, sheet);
                }
            }
        }

        /// <summary>
        /// ExcelブックのデータでDBのデータを置き換えます
        /// ワークシート名を元に対象となるテーブル名を特定し、データの書き込みを行います
        /// </summary>
        /// <param name="dbContext"></param>
        /// <param name="path">Excelファイルのパス</param>
        public static void ReadExcelReplaceDb(this DbContext dbContext, string path)
        {
            var file = new FileInfo(path);
            using (var package = new ExcelPackage(file))
            {
                foreach (var sheet in package.Workbook.Worksheets)
                {
                    dbContext.Database.ExecuteSqlCommand($"delete {sheet.Name}");
                }

                foreach (var sheet in package.Workbook.Worksheets)
                {
                    WriteTable(dbContext, sheet);
                }
            }
        }

        /// <summary>
        /// Excelブックのデータを対象のテーブルに書き込みます
        /// ワークシート名を元に対象となるテーブル名を特定し、データの書き込みを行います
        /// </summary>
        /// <param name="dbContext"></param>
        /// <param name="path">Excelファイルのパス</param>
        /// <param name="tableName"></param>
        public static void ReadExcelWriteTable(this DbContext dbContext, string path, string tableName)
        {
            var file = new FileInfo(path);
            using (var package = new ExcelPackage(file))
            {
                var sheet = package.Workbook.Worksheets.FirstOrDefault(_ => _.Name == tableName)
                    ?? throw new ArgumentException("登録対象のテーブルがExcelブックに存在しませんでした。");                
                WriteTable(dbContext, sheet);
            }
        }

        /// <summary>
        /// Excelブックのデータで対象のテーブルを置き換えます
        /// ワークシート名を元に対象となるテーブル名を特定し、データの書き込みを行います
        /// </summary>
        /// <param name="dbContext"></param>
        /// <param name="path">Excelファイルのパス</param>
        /// <param name="tableName"></param>
        public static void ReadExcelReplaceTable(this DbContext dbContext, string path, string tableName)
        {
            var file = new FileInfo(path);
            using (var package = new ExcelPackage(file))
            {
                dbContext.Database.ExecuteSqlCommand($"delete {tableName}");
                var sheet = package.Workbook.Worksheets.FirstOrDefault(_ => _.Name == tableName)
                    ?? throw new ArgumentException("登録対象のテーブルがExcelブックに存在しませんでした。");
                WriteTable(dbContext, sheet);
            }
        }

        /// <summary>
        /// ワークシートの内容をテーブルに書き込みます
        /// ワークシート1列目の値を元に、カラムを構築します
        /// ワークシート2列目以降の値を元に、データを構築します
        /// </summary>
        /// <param name="dbContext"></param>
        /// <param name="sheet"></param>
        private static void WriteTable(DbContext dbContext, ExcelWorksheet sheet)
        {
            // Excelデータを取得します
            var values = sheet.GetCellValues();

            // SQLのカラム部分を構築します
            var columns = values.FirstOrDefault()
                ?? throw new ArgumentException($"ワークシート1列目にカラム名を設定してください。[ワークシート名:{sheet.Name}]");

            var keys = columns.GroupBy(_ => _).Where(_ => _.Count() > 2).Select(_ => _.Key);
            if (keys.Count() > 0) throw new ArgumentException($"ワークシート1列目に重複しているカラム名が存在します。[カラム名:{string.Join(',', keys)}]");

            var cols = string.Join(',', columns);

            // SQLのデータ部分を構築します
            var rows = values.Skip(1).Select(l => $"({string.Join(',', l.Select(_ => _.ToUpper() != "NULL" ? $"'{_}'" : "null"))})");
            if (rows.Count() == 0) return;

            var sql = $"insert into {sheet.Name} ({cols}) values {string.Join(',', rows)}";            

            try
            {
#pragma warning disable EF1000
                dbContext.Database.ExecuteSqlCommand(sql);
#pragma warning restore EF1000
            }
            catch
            {
                Console.Error.WriteLine($"SQL実行に失敗しました:[{sql}]");
                throw;
            }
        }
    }
}
