// Copyright (c) 2024 Toshiki Iga
//
// Released under the MIT license
// https://opensource.org/license/mit

using System;
using System.IO;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using Microsoft.Data.Sqlite;


namespace NyanCEL
{
    public class SerializableDictionaryItem
    {
        [XmlAttribute("Name")]
        public string Key { get; set; }
        [XmlText]
        public string Value { get; set; }
    }

    [XmlRoot("Row")]
    public class Row
    {
        [XmlElement("Column")]
        public List<SerializableDictionaryItem> Items { get; set; } = new List<SerializableDictionaryItem>();
    }

    public class NyanSql2Xml
    {
        public async static Task<string> Sql2Xml(SqliteConnection connection, string sql)
        {
            var results = await NyanSqlQuery.Sql2ListDictionary(connection, sql);
            try
            {
                var serializableDicList = new List<Row>();
                foreach (var dict in results)
                {
                    var serializableDict = new Row();
                    foreach (var kvp in dict)
                    {
                        serializableDict.Items.Add(new SerializableDictionaryItem { Key = kvp.Key, Value = kvp.Value.ToString() });
                    }
                    serializableDicList.Add(serializableDict);
                }

                XmlSerializer serializer = new XmlSerializer(typeof(List<Row>));
                using (StringWriter stringWriter = new StringWriter())
                {
                    serializer.Serialize(stringWriter, serializableDicList);
                    stringWriter.Close();
                    return stringWriter.ToString();
                }
            }
            catch (Exception ex)
            {
                throw new Exception("NyanSql2Xml: " + ex.Message);
            }
        }

        public async static Task<string> Sql2XmlWithXPath(SqliteConnection connection, string sql, string xpath)
        {
            string original = await Sql2Xml(connection, sql);
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(original);
                try
                {
                    XmlNodeList nodes = doc.SelectNodes(xpath);
                    using (StringWriter sw = new StringWriter())
                    {
                        using (XmlWriter xmlWriter = XmlWriter.Create(sw))
                        {
                            foreach (XmlNode node in nodes)
                            {
                                node.WriteTo(xmlWriter);
                            }
                        }
                        return sw.ToString();
                    }
                }
                catch (Exception)
                {
                    XmlNode node = doc.SelectSingleNode(xpath);
                    using (StringWriter sw = new StringWriter())
                    {
                        using (XmlWriter xmlWriter = XmlWriter.Create(sw))
                        {
                            node.WriteTo(xmlWriter);
                        }
                        return sw.ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("NyanSql2Xml: " + ex.Message);
            }
        }
    }
}
