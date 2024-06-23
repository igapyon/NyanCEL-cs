// Copyright (c) 2024 Toshiki Iga
//
// Released under the MIT license
// https://opensource.org/license/mit

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Globalization;
using Microsoft.Data.Sqlite;
using ClosedXML.Excel;

namespace NyanCEL
{
    public class NyanXlsx2Sqlite
    {
        public async static Task<List<NyanTableInfo>> LoadExcelFile(SqliteConnection _connection, MemoryStream fileData)
        {
            List<NyanTableInfo> tableInfoList = new List<NyanTableInfo>();
            using (var workbook = new XLWorkbook(fileData))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    var tableInfo = new NyanTableInfo();
                    tableInfo.TableName = worksheet.Name;
                    BuildColumnInfo(worksheet, tableInfo);
                    tableInfoList.Add(tableInfo);

                    // 得られた情報をもとにテーブルを作成
                    var createTableSql = tableInfo.GetCreateTableSql();
                    try
                    {
                        var createTableCommand = _connection.CreateCommand();
                        createTableCommand.CommandText = createTableSql;
                        await createTableCommand.ExecuteNonQueryAsync();
                        Console.WriteLine(createTableSql);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("NyanXlsx2Sqlite: " + ex.Message);
                    }

                    using (var transaction = _connection.BeginTransaction())
                    {
                        await LoadData(_connection, worksheet, tableInfo);
                        transaction.Commit();
                    }
                }
            }
            return tableInfoList;
        }

        private static void BuildColumnInfo(IXLWorksheet worksheet, NyanTableInfo tableInfo)
        {
            // 1行目はヘッダー行としてテーブルの列名に利用
            var firstRow = worksheet.Row(1);
            foreach (var cell in firstRow.Cells())
            {
                if ("NyanRowId" == cell.Value.ToString())
                {
                    // その列名は NyanCEL では予約語です。スキップします。
                    continue;
                }

                var column = new NyanTableColumnInfo();
                column.ColumnName = cell.Value.ToString();
                column.ColumnIndex = cell.Address.ColumnNumber;

                // 列名ダブりチェック
                if (IsColumnNameExists(tableInfo, cell.Value.ToString()))
                {
                    // ダブり回避のため変更
                    for (int incNo = 2; ; incNo++)
                    {
                        string newName = cell.Value.ToString() + "_" + incNo;
                        if (IsColumnNameExists(tableInfo, newName))
                        {
                            continue;
                        }
                        column.ColumnName = newName;
                        break;
                    }
                }

                tableInfo.Columns.Add(column);
            }

            // 2行目以降は列のデータ型の推測のために利用
            int rowCount = worksheet.LastRowUsed().RowNumber();
            for (int rowNo = 2; rowNo <= rowCount; rowNo++)
            {
                var checkRow = worksheet.Row(rowNo);
                foreach (var cell in checkRow.Cells())
                {
                    var columnInfo = tableInfo.FindColumnByIndex(cell.Address.ColumnNumber);
                    if (columnInfo == null)
                    {
                        // 該当する列情報がないのでスキップします。
                        continue;
                    }
                    if (cell.IsEmpty() || cell.Value.IsBlank)
                    {
                        // 空のセルは判定には用いません。
                        continue;
                    }
                    if (columnInfo.NumberFormatId != -1) {
                        // 型が確定した列です。処理不要。次の項目に進みます。
                        continue;
                    }
                    var numberFormatId = cell.Style.NumberFormat.NumberFormatId;
                    switch (numberFormatId)
                    {
                        case 0:
                            // 0 : TEXT
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "TEXT";
                            break;
                        case 1:
                            // 1 : INTEGER 0
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "INTEGER";
                            break;
                        case 2:
                        case 3:
                        case 4:
                            // 2 : REAL    0.00
                            // 3 : REAL    #,##0
                            // 4 : REAL    #,##0.00
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "REAL";
                            break;
                        case 9:
                        case 10:
                            // 9 : REAL    0 %
                            // 10 : REAL   0.00 %
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "REAL";
                            break;
                        case 11:
                            // 11 : TEXT   0.00E+00
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "TEXT";
                            break;
                        case 12:
                        case 13:
                            // 12 : TEXT   # ?/?
                            // 13 : TEXT   # ??/??
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "TEXT";
                            break;
                        case 14:
                        case 15:
                            // 日付形式に変換
                            // 14 : TEXT  d / m / yyyy
                            // 15 : TEXT  d - mmm - yy
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "TEXT";
                            break;
                        case 16:
                        case 17:
                        case 18:
                        case 19:
                        case 20:
                        case 21:
                            // 各型式
                            // 16 : TEXT d - mmm
                            // 17 : TEXT mmm - yy
                            // 18 : TEXT h: mm tt
                            // 19 : TEXT h: mm: ss tt
                            // 20 : TEXT H: mm
                            // 21 : TEXT H: mm: ss
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "TEXT";
                            break;
                        case 22:
                            // 日付形式に変換
                            // 22 : TEXT  m / d / yyyy H: mm
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "TEXT";
                            break;
                        case 37:
                        case 38:
                            // 値そのまま
                            // 37 : INTEGER    #,##0 ;(#,##0)
                            // 38 : INTEGER    #,##0 ;[Red](#,##0)
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "INTEGER";
                            break;
                        case 39:
                        case 40:
                            // 値そのまま
                            // 39 : REAL   #,##0.00;(#,##0.00)
                            // 40 : REAL   #,##0.00;[Red](#,##0.00)
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "REAL";
                            break;
                        case 45:
                        case 46:
                        case 47:
                            // フォーマット適用後のものを取得
                            // 45 : TEXT   mm:ss
                            // 46 : TEXT   [h]:mm:ss
                            // 47 : TEXT   mmss.0
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "TEXT";
                            break;
                        case 48:
                            // 48 : TEXT   ##0.0E+0
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "TEXT";
                            break;
                        case 49:
                            // 49 : TEXT
                            columnInfo.NumberFormatId = numberFormatId;
                            columnInfo.ColumnType = "TEXT";
                            break;
                        case -1: // 指定なし
                            // NumberFormatId は確定せず。
                            columnInfo.NumberFormatId = numberFormatId;
                            var value = cell.Value;
                            if (value.IsNumber)
                            {
                                columnInfo.ColumnType = "REAL"; // 仮で型を設定
                            }
                            else
                            {
                                columnInfo.ColumnType = "TEXT"; // 仮で型を設定
                            }
                            break;
                        default:
                            Console.WriteLine("Unexpected: TypeCheck: Unknown NumberFormatId: " + columnInfo.NumberFormatId + ", ColumnName: " + columnInfo.ColumnName);
                            columnInfo.ColumnType = "TEXT";
                            // NumberFormatId を上書き
                            columnInfo.NumberFormatId = 0;
                            break;
                    }
                }
            }

            Console.WriteLine($"Table: {tableInfo.TableName}");
            foreach (var columnInfo in tableInfo.Columns)
            {
                Console.WriteLine($"    {columnInfo.ColumnName} ({columnInfo.ColumnType}) : NumberFormatId:{columnInfo.NumberFormatId}");
            }
        }

        private static bool IsColumnNameExists(NyanTableInfo tableInfo, string name)
        {
            foreach (var columnInfo in tableInfo.Columns)
            {
                if (columnInfo.ColumnName == name)
                {
                    return true;
                }
            }
            return false;
        }

        private async static Task LoadData(SqliteConnection _connection, IXLWorksheet worksheet, NyanTableInfo tableInfo)
        {
            int totalRowCount = 0;
            // 2行目以降のデータをテーブルにINSERT
            int rowCount = worksheet.LastRowUsed().RowNumber();
            for (int rowNo = 2; rowNo <= rowCount; rowNo++)
            {
                string insertSql = "INSERT INTO \"" + tableInfo.TableName + "\"";
                var currentRow = worksheet.Row(rowNo);
                var cells = currentRow.Cells();
                string colNameStr = "";
                string colValueStr = "";
                List<SqliteParameter> colValueList = new List<SqliteParameter>();
                bool isFirstCol = true;
                int columnCount = 0;
                foreach (var cell in cells)
                {
                    var tableColumnInfo = tableInfo.FindColumnByIndex(cell.Address.ColumnNumber);
                    if (tableColumnInfo == null)
                    {
                        // 該当する列の列情報がないのでスキップします。
                        continue;
                    }
                    if (isFirstCol)
                    {
                        isFirstCol = false;
                    }
                    else
                    {
                        colNameStr += ", ";
                        colValueStr += ", ";
                    }
                    colNameStr += "\"" + tableColumnInfo.ColumnName + "\"";
                    colValueStr += "$col" + columnCount;

                    var param = new SqliteParameter();
                    param.ParameterName = "$col" + columnCount;
                    switch (tableColumnInfo.NumberFormatId)
                    {
                        case 0:
                            // 0 : TEXT
                            if (cell.TryGetValue(out string str0))
                            {
                                param.Value = str0;
                            }
                            else
                            {
                                param.Value = cell.GetFormattedString();
                            }
                            break;
                        case 1:
                            // 1 : INTEGER 0
                            param.Value = GetIntegerValueString(cell);
                            break;
                        case 2:
                        case 3:
                        case 4:
                            // 2 : REAL    0.00
                            // 3 : REAL    #,##0
                            // 4 : REAL    #,##0.00
                            param.Value = cell.Value.GetNumber().ToString();
                            break;
                        case 9:
                        case 10:
                            // 9 : REAL    0 %
                            // 10 : REAL   0.00 %
                            param.Value = cell.Value.GetNumber().ToString();
                            break;
                        case 11:
                            // 11 : TEXT   0.00E+00
                            param.Value = cell.GetFormattedString();
                            break;
                        case 12:
                        case 13:
                            // 12 : TEXT   # ?/?
                            // 13 : TEXT   # ??/??
                            param.Value = cell.GetFormattedString();
                            break;
                        case 14:
                        case 15:
                            // 14 : TEXT  d / m / yyyy : Format yyyy-MM-dd
                            // 15 : TEXT  d - mmm - yy : Format yyyy-MM-dd
                            if (cell.TryGetValue(out DateTime dateTimeValue14))
                            {
                                param.Value = dateTimeValue14.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                param.Value = cell.GetFormattedString();
                            }
                            break;
                        case 16:
                            // 16 : TEXT d - mmm : Format MM-dd
                            if (cell.TryGetValue(out DateTime dateTimeValue16))
                            {
                                param.Value = dateTimeValue16.ToString("MM-dd", CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                param.Value = cell.GetFormattedString();
                            }
                            break;
                        case 17:
                            // 17 : TEXT mmm - yy : Format yyyy-MM
                            if (cell.TryGetValue(out DateTime dateTimeValue17))
                            {
                                param.Value = dateTimeValue17.ToString("yyyy-MM", CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                param.Value = cell.GetFormattedString();
                            }
                            break;
                        case 18:
                            // 18 : TEXT h: mm tt
                            param.Value = cell.GetFormattedString();
                            break;
                        case 19:
                            // 19 : TEXT h: mm: ss tt
                            param.Value = cell.GetFormattedString();
                            break;
                        case 20:
                            // 20 : TEXT H: mm : Format HH:mm
                            if (cell.TryGetValue(out DateTime dateTimeValue20))
                            {
                                param.Value = dateTimeValue20.ToString("HH:mm", CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                param.Value = cell.GetFormattedString();
                            }
                            break;
                        case 21:
                            // 21 : TEXT H: mm: ss : Format HH:mm:ss
                            if (cell.TryGetValue(out DateTime dateTimeValue21))
                            {
                                param.Value = dateTimeValue21.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                param.Value = cell.GetFormattedString();
                            }
                            break;
                        case 22:
                            // 22 : TEXT  m / d / yyyy H: mm : Format yyyy-MM-dd HH:mm
                            if (cell.TryGetValue(out DateTime dateTimeValue22))
                            {
                                param.Value = dateTimeValue22.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                param.Value = cell.GetFormattedString();
                            }
                            break;
                        case 37:
                            // 37 : INTEGER    #,##0 ;(#,##0)
                            param.Value = GetIntegerValueString(cell);
                            break;
                        case 38:
                            // 38 : INTEGER    #,##0 ;[Red](#,##0)
                            param.Value = GetIntegerValueString(cell);
                            break;
                        case 39:
                        case 40:
                            // 値そのまま。フォーマットは当てない
                            // 39 : REAL   #,##0.00;(#,##0.00)
                            // 40 : REAL   #,##0.00;[Red](#,##0.00)
                            param.Value = cell.Value.GetNumber().ToString();
                            break;
                        case 45:
                            // 45 : TEXT   mm:ss : Format mm:ss
                            if (cell.TryGetValue(out DateTime dateTimeValue45))
                            {
                                param.Value = dateTimeValue45.ToString("mm:ss", CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                param.Value = cell.GetFormattedString();
                            }
                            break;
                        case 46:
                            // 46 : TEXT   [h]:mm:ss : Format HH:mm:ss
                            if (cell.TryGetValue(out DateTime dateTimeValue46))
                            {
                                param.Value = dateTimeValue46.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
                            }
                            else
                            {
                                param.Value = cell.GetFormattedString();
                            }
                            break;
                        case 47:
                            // 47 : TEXT   mmss.0
                            param.Value = cell.GetFormattedString();
                            break;
                        case 48:
                            // 48 : TEXT   ##0.0E+0
                            param.Value = cell.GetFormattedString();
                            break;
                        case 49:
                            // 49 : TEXT
                            if (cell.TryGetValue(out string str49))
                            {
                                param.Value = str49;
                            }
                            else
                            {
                                param.Value = cell.Value.ToString();
                            }
                            break;
                        case -1: // 指定なし
                            var value = cell.Value;
                            if (value.IsNumber)
                            {
                                param.Value = cell.Value.GetNumber().ToString();
                            }
                            else
                            {
                                // フォーマット付きで文字列化
                                param.Value = cell.GetFormattedString();
                            }
                            break;
                        default:
                            throw new Exception("Unexpected: GetValue: Unknown NumberFormatId: " + tableColumnInfo.NumberFormatId);
                    }
                    colValueList.Add(param);
                    columnCount++;
                }

                insertSql += " (" + colNameStr + ") VALUES (" + colValueStr + ")";

                var insertIntoCommand = _connection.CreateCommand();
                insertIntoCommand.CommandText = insertSql;
                foreach (var param in colValueList)
                {
                    insertIntoCommand.Parameters.Add(param);
                }
                await insertIntoCommand.ExecuteNonQueryAsync();
                totalRowCount++;

                if (totalRowCount <= 5)
                {
                    // 最初の数行のINSERTについてログ
                    Console.WriteLine(insertSql);
                }
            }
            Console.WriteLine($"  Record count: {tableInfo.TableName}: {totalRowCount}");
        }

        private static string GetIntegerValueString(IXLCell cell) {
            if (cell == null)
            {
                return null;
            }
            if (cell.Value.IsBlank)
            {
                return null;
            }
            var value = cell.Value.GetNumber().ToString();
            if (value.EndsWith(".0"))
            {
                return value.Substring(0, value.Length - 2);
            }
            return value;
        }
    }
}
