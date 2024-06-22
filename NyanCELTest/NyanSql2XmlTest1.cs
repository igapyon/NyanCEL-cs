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
    public class NyanSql2XmlTest1
    {
        [TestMethod]
        public async Task Sql2XmlTest1()
        {
            using (var connection = await NyanCELUtil.CreateXlsxDatabase())
            {
              var memoryStream = await NyanCELUtil.ReadBinaryFile2MemoryStream("./TestData/Book1.xlsx");
              List<NyanTableInfo> tableInfoList = await NyanXlsx2Sqlite.LoadExcelFile(connection, memoryStream);

              string resultString = await NyanSql2Xml.Sql2Xml(connection, "SELECT * FROM sqlite_master");
              Assert.AreEqual(@"<?xml version=""1.0"" encoding=""utf-16""?>
<ArrayOfRow xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
  <Row>
    <Column Name=""type"">table</Column>
    <Column Name=""name"">11SAITAM</Column>
    <Column Name=""tbl_name"">11SAITAM</Column>
    <Column Name=""rootpage"">2</Column>
    <Column Name=""sql"">CREATE TABLE ""11SAITAM"" (""JISX0401"" TEXT, ""OLDZIPCODE"" TEXT, ""ZIPCODE"" TEXT, ""TODOFUKEN_ZENKANA"" TEXT, ""SIKUTYOUSON_ZENKANA"" TEXT, ""TYOUIKI_ZENKANA"" TEXT, ""TODOFUKEN_KANJI"" TEXT, ""SIKUTYOSON_KANJI"" TEXT, ""TYOUIKI_KANJI"" TEXT, ""DUPZIP"" REAL, ""KOAZAKIBAN"" REAL, ""HAS_TYOUME"" REAL, ""DUP_TYOUIKI"" REAL, ""UPDATED"" REAL, ""UPDATE_REASON"" REAL, NyanRowId INTEGER PRIMARY KEY)</Column>
  </Row>
  <Row>
    <Column Name=""type"">table</Column>
    <Column Name=""name"">TypeCheck</Column>
    <Column Name=""tbl_name"">TypeCheck</Column>
    <Column Name=""rootpage"">13</Column>
    <Column Name=""sql"">CREATE TABLE ""TypeCheck"" (""STANDARD"" TEXT, ""NUMBER"" REAL, ""CURRENCY"" REAL, ""FINANCE"" REAL, ""DATE"" TEXT, ""TIME"" TEXT, ""PERCENT"" REAL, ""BUNSU"" TEXT, ""EXP"" REAL, ""STRING"" TEXT, ""OTHER"" TEXT, ""USERDEF"" TEXT, NyanRowId INTEGER PRIMARY KEY)</Column>
  </Row>
  <Row>
    <Column Name=""type"">table</Column>
    <Column Name=""name"">SameName</Column>
    <Column Name=""tbl_name"">SameName</Column>
    <Column Name=""rootpage"">14</Column>
    <Column Name=""sql"">CREATE TABLE ""SameName"" (""SameName"" TEXT, ""SameName_2"" TEXT, ""SameName_3"" TEXT, ""OtherName"" TEXT, ""SameName_4"" REAL, ""OtherName_2"" TEXT, ""SameName_5"" TEXT, NyanRowId INTEGER PRIMARY KEY)</Column>
  </Row>
  <Row>
    <Column Name=""type"">table</Column>
    <Column Name=""name"">CellFormat</Column>
    <Column Name=""tbl_name"">CellFormat</Column>
    <Column Name=""rootpage"">15</Column>
    <Column Name=""sql"">CREATE TABLE ""CellFormat"" (""Field1"" TEXT, ""Field2"" TEXT, ""Field3"" REAL, ""Field4"" TEXT, ""Field5"" REAL, ""Field6"" TEXT, ""Field7"" REAL, NyanRowId INTEGER PRIMARY KEY)</Column>
  </Row>
</ArrayOfRow>", resultString);
          }
        }

        [TestMethod]
        public async Task Sql2XmlWithXPathTest1()
        {
            using (var connection = await NyanCELUtil.CreateXlsxDatabase())
            {
              var memoryStream = await NyanCELUtil.ReadBinaryFile2MemoryStream("./TestData/Book1.xlsx");
              List<NyanTableInfo> tableInfoList = await NyanXlsx2Sqlite.LoadExcelFile(connection, memoryStream);

              string resultString = await NyanSql2Xml.Sql2XmlWithXPath(connection, "SELECT * FROM sqlite_master", "/ArrayOfRow/Row[1]");
              Assert.AreEqual(@"<?xml version=""1.0"" encoding=""utf-16""?><Row><Column Name=""type"">table</Column><Column Name=""name"">11SAITAM</Column><Column Name=""tbl_name"">11SAITAM</Column><Column Name=""rootpage"">2</Column><Column Name=""sql"">CREATE TABLE ""11SAITAM"" (""JISX0401"" TEXT, ""OLDZIPCODE"" TEXT, ""ZIPCODE"" TEXT, ""TODOFUKEN_ZENKANA"" TEXT, ""SIKUTYOUSON_ZENKANA"" TEXT, ""TYOUIKI_ZENKANA"" TEXT, ""TODOFUKEN_KANJI"" TEXT, ""SIKUTYOSON_KANJI"" TEXT, ""TYOUIKI_KANJI"" TEXT, ""DUPZIP"" REAL, ""KOAZAKIBAN"" REAL, ""HAS_TYOUME"" REAL, ""DUP_TYOUIKI"" REAL, ""UPDATED"" REAL, ""UPDATE_REASON"" REAL, NyanRowId INTEGER PRIMARY KEY)</Column></Row>", resultString);
            }
        }
    }
}