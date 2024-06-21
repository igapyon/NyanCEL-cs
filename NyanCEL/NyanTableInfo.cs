// Copyright (c) 2024 Toshiki Iga
//
// Released under the MIT license
// https://opensource.org/license/mit

using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace NyanCEL
{
    public class NyanTableInfo
    {
        public string TableName { get; set; }

        public List<NyanTableColumnInfo> Columns = new List<NyanTableColumnInfo>();

        public NyanTableColumnInfo FindColumnByIndex(int index)
        {
            foreach (NyanTableColumnInfo column in Columns)
            {
                if (column.ColumnIndex == index)
                {
                    return column;
                }
            }
            // TODO Exception
            return null;
        }

        public string GetCreateTableSql()
        {
            string result = "CREATE TABLE \"" + TableName + "\" (";
            for (int i = 0; i < Columns.Count; i++)
            {
                var col = Columns[i];
                if (i > 0)
                {
                    result += ", ";
                }
                result += "\"" + col.ColumnName + "\"" + " " + col.ColumnType;
            }
            result += ", NyanRowId INTEGER PRIMARY KEY)";
            return result;
        }
//        public void LogTableInfo()
//        {
//            NyanLog.Info($"TableName: {TableName}");
//            foreach ( NyanTableColumnInfo column in Columns)
//            {
//                NyanLog.Info($"  ColumnName: {column.ColumnName}");
//                NyanLog.Info($"    ColumnIndex: {column.ColumnIndex}");
//                NyanLog.Info($"    NumberFormatId: {column.NumberFormatId}");
//                NyanLog.Info($"    ColumnType: {column.ColumnType}");
//            }
//        }
    }

    public class NyanTableColumnInfo
    {
        public string ColumnName { get; set; }
        public string ColumnType { get; set; }
        public int ColumnIndex { get; set; }
        public int NumberFormatId { get; set; }

        public NyanTableColumnInfo() {
            NumberFormatId = -1; // 初期値は -1 : "指定なし"
        }
    }
}
