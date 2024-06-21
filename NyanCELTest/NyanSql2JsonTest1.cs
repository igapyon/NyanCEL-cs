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
    public class NyanSql2JsonTest1
    {
        [TestMethod]
        public async Task Sql2JsonTest1()
        {
            SqliteConnection connection = new SqliteConnection("Data Source=:memory:");
            await connection.OpenAsync();

            byte[] data = NyanCELUtil.ReadBinaryFile("./TestData/Book1.xlsx");
            var memoryStream = await NyanCELUtil.ByteArray2MemoryStream(data);

            List<NyanTableInfo> tableInfoList = new List<NyanTableInfo>();
            await NyanXlsx2Sqlite.LoadExcelFileAsync(connection, memoryStream, tableInfoList);

            string resultString = await NyanSql2Json.Sql2Json(connection, "SELECT * FROM sqlite_master");
            Assert.AreEqual(@"[
  {
    ""type"": ""table"",
    ""name"": ""11SAITAM"",
    ""tbl_name"": ""11SAITAM"",
    ""rootpage"": 2,
    ""sql"": ""CREATE TABLE \""11SAITAM\"" (\""JISX0401\"" TEXT, \""OLDZIPCODE\"" TEXT, \""ZIPCODE\"" TEXT, \""TODOFUKEN_ZENKANA\"" TEXT, \""SIKUTYOUSON_ZENKANA\"" TEXT, \""TYOUIKI_ZENKANA\"" TEXT, \""TODOFUKEN_KANJI\"" TEXT, \""SIKUTYOSON_KANJI\"" TEXT, \""TYOUIKI_KANJI\"" TEXT, \""DUPZIP\"" REAL, \""KOAZAKIBAN\"" REAL, \""HAS_TYOUME\"" REAL, \""DUP_TYOUIKI\"" REAL, \""UPDATED\"" REAL, \""UPDATE_REASON\"" REAL, NyanRowId INTEGER PRIMARY KEY)""
  },
  {
    ""type"": ""table"",
    ""name"": ""TypeCheck"",
    ""tbl_name"": ""TypeCheck"",
    ""rootpage"": 13,
    ""sql"": ""CREATE TABLE \""TypeCheck\"" (\""STANDARD\"" TEXT, \""NUMBER\"" REAL, \""CURRENCY\"" REAL, \""FINANCE\"" REAL, \""DATE\"" TEXT, \""TIME\"" TEXT, \""PERCENT\"" REAL, \""BUNSU\"" TEXT, \""EXP\"" REAL, \""STRING\"" TEXT, \""OTHER\"" TEXT, \""USERDEF\"" TEXT, NyanRowId INTEGER PRIMARY KEY)""
  },
  {
    ""type"": ""table"",
    ""name"": ""SameName"",
    ""tbl_name"": ""SameName"",
    ""rootpage"": 14,
    ""sql"": ""CREATE TABLE \""SameName\"" (\""SameName\"" TEXT, \""SameName_2\"" TEXT, \""SameName_3\"" TEXT, \""OtherName\"" TEXT, \""SameName_4\"" REAL, \""OtherName_2\"" TEXT, \""SameName_5\"" TEXT, NyanRowId INTEGER PRIMARY KEY)""
  },
  {
    ""type"": ""table"",
    ""name"": ""CellFormat"",
    ""tbl_name"": ""CellFormat"",
    ""rootpage"": 15,
    ""sql"": ""CREATE TABLE \""CellFormat\"" (\""Field1\"" TEXT, \""Field2\"" TEXT, \""Field3\"" REAL, \""Field4\"" TEXT, \""Field5\"" REAL, \""Field6\"" TEXT, \""Field7\"" REAL, NyanRowId INTEGER PRIMARY KEY)""
  }
]", resultString);
            Console.WriteLine(resultString);
        }

        [TestMethod]
        public async Task Sql2JsonWithJSONPathTest1()
        {
            SqliteConnection connection = new SqliteConnection("Data Source=:memory:");
            await connection.OpenAsync();

            byte[] data = NyanCELUtil.ReadBinaryFile("./TestData/Book1.xlsx");
            var memoryStream = await NyanCELUtil.ByteArray2MemoryStream(data);

            List<NyanTableInfo> tableInfoList = new List<NyanTableInfo>();
            await NyanXlsx2Sqlite.LoadExcelFileAsync(connection, memoryStream, tableInfoList);

            string resultString = await NyanSql2Json.Sql2JsonWithJSONPath(connection, "SELECT * FROM sqlite_master", "[1].name");
            Assert.AreEqual(@"TypeCheck", resultString);
            Console.WriteLine(resultString);
        }

        [TestMethod]
        public async Task Sql2JsonWithTargetTest1()
        {
            SqliteConnection connection = new SqliteConnection("Data Source=:memory:");
            await connection.OpenAsync();

            byte[] data = NyanCELUtil.ReadBinaryFile("./TestData/Book1.xlsx");
            var memoryStream = await NyanCELUtil.ByteArray2MemoryStream(data);

            List<NyanTableInfo> tableInfoList = new List<NyanTableInfo>();
            await NyanXlsx2Sqlite.LoadExcelFileAsync(connection, memoryStream, tableInfoList);

            string resultString = await NyanSql2Json.Sql2JsonWithTarget(connection, "SELECT * FROM sqlite_master", "data.1.name");
            Assert.AreEqual(@"TypeCheck", resultString);
            Console.WriteLine(resultString);
        }
    }
}