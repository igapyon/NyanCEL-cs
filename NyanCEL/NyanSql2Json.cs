// Copyright (c) 2024 Toshiki Iga
//
// Released under the MIT license
// https://opensource.org/license/mit

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace NyanCEL
{
    public class NyanSql2Json
    {
        public async static Task<string> Sql2Json(SqliteConnection connection, string sql)
        {
            var results = await NyanSqlQuery.Sql2ListDictionary(connection, sql);
            try
            {
                string jsonResult = JsonConvert.SerializeObject(results, Formatting.Indented);
                return jsonResult;
            }
            catch (Exception ex)
            {
                throw new Exception("NyanSql2Json: " + ex.Message);
            }
        }

        public async static Task<string> Sql2JsonWithJSONPath(SqliteConnection connection, string sql, string jsonpath)
        {
            string original = await Sql2Json(connection, sql);
            try
            {
                JArray jsonArray = JArray.Parse(original);

                var selected = jsonArray.SelectTokens(jsonpath);
                if (selected.Count() == 1)
                {
                    JToken singleNode = selected.First();
                    return singleNode.ToString();
                }
                else if (selected.Count() > 1)
                {
                    string result = "";
                    foreach (var node in selected)
                    {
                        result+= node.ToString();
                    }
                    return result;
                }
                else
                {
                    return "";
                }
            }
            catch (Exception ex)
            {
                throw new Exception("NyanSql2Json: " + ex.Message);
            }
        }

        public async static Task<string> Sql2JsonWithTarget(SqliteConnection connection, string sql, string target)
        {
            try
            {
                if (string.IsNullOrEmpty(target))
                {
                    // 正常系
                    return await Sql2Json(connection, sql);
                }

                if (target.Trim().Equals("status"))
                {
                    string original = await Sql2Json(connection, sql);
                    // 例外出ずに進んだらOK
                    return "OK";
                }
                if (!target.Trim().StartsWith("data."))
                {
                    return "NG";
                }

                string targetData = target.Substring(5);

                int dot = targetData.IndexOf('.');
                if (dot < 0) {
                    try
                    {
                        int rowId = int.Parse(targetData);
                        return await Sql2JsonWithJSONPath(connection, sql, $"[{rowId}]");
                    }
                    catch (Exception)
                    {
                        return "";
                    }
                }

                string left = targetData.Substring(0, dot);
                string right = targetData.Substring(dot + 1);
                try
                {
                    int rowId = int.Parse(left);
                    return await Sql2JsonWithJSONPath(connection, sql, $"[{rowId}].{right}");
                }
                catch (Exception)
                {
                    if (target.Trim().Equals("status"))
                    {
                        return "NG";
                    }
                    else
                    {
                        return "";
                    }
                }
            }
            catch (Exception)
            {
                if (target.Trim().Equals("status"))
                {
                    return "NG";
                }
                else
                {
                    return "";
                }
            }
        }
    }
}
