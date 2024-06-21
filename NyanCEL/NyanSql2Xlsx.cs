// Copyright (c) 2024 Toshiki Iga
//
// Released under the MIT license
// https://opensource.org/license/mit

using System;
using System.IO;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Data.Sqlite;

namespace NyanCEL
{
    public class NyanSql2Xlsx
    {
        public async static Task<byte[]> Sql2Xlsx(SqliteConnection connection, string sql)
        {
            var results = await NyanSqlQuery.Sql2ListDictionary(connection, sql);
            try
            {
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("NyanCEL");

                if (results.Count == 0)
                {
                    throw new Exception("NyanSql2Xlsx: No data.");
                }
                var firstRow = results[0];
                int colNo = 1;
                // タイトル行
                foreach (var row in firstRow) {
                    worksheet.Cell(1, colNo).Value = row.Key;
                    worksheet.Cell(1, colNo).Style.Font.Bold = true;
                    colNo++;
                }

                // データ行
                int rowNo = 2;
                foreach (var row in results)
                {
                    colNo = 1;
                    foreach (var col in row)
                    {
                        if (col.Value.GetType() == typeof(System.Double))
                        {
                            worksheet.Cell(rowNo, colNo).Value = (double) col.Value;
                        }
                        else if (col.Value.GetType() == typeof(System.Int64))
                        {
                            worksheet.Cell(rowNo, colNo).Value = (Int64) col.Value;
                            worksheet.Cell(rowNo, colNo).Style.NumberFormat.Format = "#,##0";
                        }
                        else
                        {
                            worksheet.Cell(rowNo, colNo).Value = col.Value.ToString();
                        }
                        colNo++;
                    }
                    rowNo++;
                }

                using (var memoryStream = new MemoryStream())
                {
                    workbook.SaveAs(memoryStream);
                    byte[] binaryData = memoryStream.ToArray();
                    return binaryData;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("NyanSql2Xlsx: " + ex.Message);
            }
        }
    }
}
