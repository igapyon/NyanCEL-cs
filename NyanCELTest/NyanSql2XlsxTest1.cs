// Copyright (c) 2024 Toshiki Iga
//
// Released under the MIT license
// https://opensource.org/license/mit

using System;
using System.IO;
using Microsoft.Data.Sqlite;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NuGet.Frameworks;
using NyanCEL;

namespace NyanCELTest
{
    [TestClass]
    public class NyanSql2XlsxTest1
    {
        [TestMethod]
        public async Task Sql2XlsxTest1()
        {
            SqliteConnection connection = new SqliteConnection("Data Source=:memory:");
            await connection.OpenAsync();

            byte[] data = NyanCELUtil.ReadBinaryFile("./TestData/Book1.xlsx");
            var memoryStream = await NyanCELUtil.ByteArray2MemoryStream(data);

            List<NyanTableInfo> tableInfoList = new List<NyanTableInfo>();
            await NyanXlsx2Sqlite.LoadExcelFileAsync(connection, memoryStream, tableInfoList);

            byte[] resultByteArray = await NyanSql2Xlsx.Sql2Xlsx(connection, "SELECT * FROM sqlite_master");
            // TODO Check binary result data.
        }
    }
}