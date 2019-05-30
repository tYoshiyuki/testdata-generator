using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace TestdataGenerator.Library.Extentions
{
    public static class ExcelExtention
    {
        /// <summary>
        /// 対象となるExcelワークシートからセルの値を取得し、対応するオブジェクトのリストを生成します
        /// ワークシート1列目の値を元に、マッピング対象のオブジェクトのプロパティをマッピングします
        /// ワークシート2列目以降の値を元に、オブジェクトの値を設定します
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="worksheet"></param>
        /// <param name="map">Excelとオブジェクト プロパティ名のマッピング情報</param>
        /// <returns></returns>
        public static List<T> ToList<T>(this ExcelWorksheet worksheet, Dictionary<string, string> map = null) where T : new()
        {
            var props = typeof(T).GetProperties()
                .Select(prop =>
                {
                    var displayAttribute = (DisplayAttribute)prop.GetCustomAttributes(typeof(DisplayAttribute), false).FirstOrDefault();
                    return new
                    {
                        prop.Name,
                        PropertyInfo = prop,
                        prop.PropertyType,
                    };
                })
            .ToList();

            var values = worksheet.GetCellValues();
            var columnMaps = new List<ExcelMap>();

            // 1列目よりカラム情報を取得します
            var columns = values.FirstOrDefault()
                ?? throw new ArgumentException($"ワークシート1列目にマッピング対象のオブジェクトのプロパティ名を設定してください。[ワークシート名:{worksheet.Name}]");

            var keys = columns.GroupBy(_ => _).Where(_ => _.Count() > 1).Select(_ => _.Key);
            if (keys.Count() > 0) throw new ArgumentException($"ワークシート1列目に重複しているプロパティ名が存在します。[プロパティ名:{string.Join(',', keys)}]");

            var i = 1;
            foreach(var col in columns)
            {
                columnMaps.Add(new ExcelMap()
                {
                    Name = col,
                    // デフォルトはカラム名をマッピング対象とし、マッピング情報を指定した場合はそちらを採用します
                    MappedTo = map == null || map.Count == 0 ? col :
                        map.ContainsKey(col) ? map[col] : string.Empty,
                    Index = i
                });
                i++;
            }

            var retList = new List<T>();
            // 2列目以降よりデータを取得します
            i = 1;
            foreach (var row in values.Skip(1))
            {
                var item = new T();
                foreach(var column in columnMaps)
                {
                    // マッピング情報を元にマッピング先のプロパティ情報を取得します
                    // マッピング先が見つからない場合は、処理をスルーします
                    var prop = string.IsNullOrWhiteSpace(column.MappedTo) ? null : props.FirstOrDefault(p => p.Name.Contains(column.MappedTo));
                    if (prop != null)
                    {
                        var value = row[column.Index - 1].ToString();
                        if (value.ToUpper() == "NULL") value = null; // NULLを指定していた場合の処理

                        var propType = prop.PropertyType;
                        object parsed = null;

                        // プロパティの型に応じて変換処理を行います
                        switch (propType)
                        {
                            case Type intType when propType == typeof(int): parsed = int.Parse(value); break;
                            case Type nullableIntType when propType == typeof(int?): if (!string.IsNullOrEmpty(value)) parsed = (int?)int.Parse(value); break;
                            case Type shortType when propType == typeof(short): parsed = short.Parse(value); break;
                            case Type nullableShortType when propType == typeof(short?): if (!string.IsNullOrEmpty(value)) parsed = (short?)short.Parse(value); break;
                            case Type longType when propType == typeof(long): parsed = long.Parse(value); break;
                            case Type nullableLongType when propType == typeof(long?): if (!string.IsNullOrEmpty(value)) parsed = (long?)long.Parse(value); break;
                            case Type longType when propType == typeof(decimal): parsed = decimal.Parse(value); break;
                            case Type nullableDecimalType when propType == typeof(decimal?): if (!string.IsNullOrEmpty(value)) parsed = (decimal?)decimal.Parse(value); break;
                            case Type doubleType when propType == typeof(double): parsed = double.Parse(value); break;
                            case Type nullableDoubleType when propType == typeof(double?): if (!string.IsNullOrEmpty(value)) parsed = (double?)double.Parse(value); break;
                            case Type dateTimeType when propType == typeof(DateTime): parsed = DateTime.Parse(value); break;
                            case Type nullableDateTimeType when propType == typeof(DateTime?): if (!string.IsNullOrEmpty(value)) parsed = DateTime.Parse(value); break;
                            case Type stringType when propType == typeof(string): parsed = value; break;
                            default: parsed = Convert.ChangeType(value, propType); break;
                        };
                        prop.PropertyInfo.SetValue(item, parsed);
                    }
                }
                retList.Add(item);
            }
            return retList;
        }

        /// <summary>
        /// 対象のワークシートよりセルの値を取得します
        /// 値を設定したセルの列最大値・行最大値の範囲に対して値を取得します
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static IEnumerable<List<string>> GetCellValues(this ExcelWorksheet worksheet)
        {
            if (worksheet.Dimension == null) yield break;
            for (var i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                var list = new List<string>();
                for (var j = 1; j <= worksheet.Dimension.End.Column; j++)
                {
                    list.Add(worksheet.Cells[i, j].Value?.ToString() ?? string.Empty);
                }
                yield return list;
            }
        }
    }

    /// <summary>
    /// マッピング情報です
    /// </summary>
    public class ExcelMap
    {
        /// <summary>
        /// ワークシートで指定している項目名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// マッピング先のプロパティ名
        /// </summary>
        public string MappedTo { get; set; }

        /// <summary>
        /// インデックス
        /// </summary>
        public int Index { get; set; }
    }
}
