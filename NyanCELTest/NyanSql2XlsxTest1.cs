// Copyright (c) 2024 Toshiki Iga
//
// Released under the MIT license
// https://opensource.org/license/mit

using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using NyanCEL;

namespace NyanCELTest
{
    [TestClass]
    public class NyanSql2XlsxTest1
    {
        [TestMethod]
        public async Task Sql2XlsxTest1()
        {
            using (var connection = await NyanCELUtil.CreateXlsxDatabase())
            {
                var memoryStream = await NyanCELUtil.ReadBinaryFile2MemoryStream("./TestData/Book1.xlsx");
                List<NyanTableInfo> tableInfoList = await NyanXlsx2Sqlite.LoadExcelFile(connection, memoryStream);

                byte[] resultByteArray = await NyanSql2Xlsx.Sql2Xlsx(connection, "SELECT * FROM sqlite_master");
                // TODO Check binary result data.
            }
        }
    }
}