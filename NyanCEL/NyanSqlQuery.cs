// Copyright (c) 2024 Toshiki Iga
//
// Released under the MIT license
// https://opensource.org/license/mit

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;

namespace NyanCEL
{
    public class NyanSqlQuery
    {
        public async static Task<List<Dictionary<string, object>>> Sql2ListDictionary(SqliteConnection connection, string sql)
        {
            try
            {
                using (var readonlyTransaction = connection.BeginTransaction())
                {
                    var results = new List<Dictionary<string, object>>();
                    var selectCommand = connection.CreateCommand();
                    selectCommand.Transaction = readonlyTransaction; // READONLY
                    selectCommand.CommandText = sql;
                    using (var reader = await selectCommand.ExecuteReaderAsync())
                    {
                        while (await reader.ReadAsync())
                        {
                            var row = new Dictionary<string, object>();
                            for (int colNo = 0; colNo < reader.FieldCount; colNo++)
                            {
                                row[reader.GetName(colNo)] = reader.GetValue(colNo);
                            }
                            results.Add(row);
                        }
                    }

                    return results;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("NyanSqlQuery: " + ex.Message);
            }
        }
    }
}
