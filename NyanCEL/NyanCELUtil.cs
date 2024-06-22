// Copyright (c) 2024 Toshiki Iga
//
// Released under the MIT license
// https://opensource.org/license/mit

using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;

namespace NyanCEL
{
    public class NyanCELUtil
    {
        private static SqliteConnection _connection = null;

        public static async Task<SqliteConnection> GetXlsxDatabaseInstance() {
            if (_connection != null) 
            {
                return _connection;
            }

            _connection = await CreateXlsxDatabase();
            return _connection;
        }

        public static async Task<SqliteConnection> CreateXlsxDatabase() {
            SqliteConnection connection = new SqliteConnection("Data Source=:memory:");
            await connection.OpenAsync();
            return connection;
        }

        public static async Task<MemoryStream> ReadBinaryFile2MemoryStream(string filePath)
        {
              byte[] data = NyanCELUtil.ReadBinaryFile("./TestData/Book1.xlsx");
              return await NyanCELUtil.ByteArray2MemoryStream(data);
        }

        public static async Task<MemoryStream> ByteArray2MemoryStream(byte[] data)
        {
            var memoryStream = new MemoryStream();
            await memoryStream.WriteAsync(data.ToArray(), 0, (int)data.Length);
            memoryStream.Seek(0, SeekOrigin.Begin);
            return memoryStream;
        }

        public static byte[] ReadBinaryFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("File not found.", filePath);
            }

            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                using (BinaryReader br = new BinaryReader(fs))
                {
                    byte[] data = new byte[fs.Length];
                    br.Read(data, 0, data.Length);
                    return data;
                }
            }
        }
    }
}
