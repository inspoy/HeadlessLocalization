using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Mono.Data;
using Mono.Data.Sqlite;

namespace HeadlessL10nExporter
{
    public sealed class SqliteHelper : IDisposable
    {
        /// <summary>
        /// 数据库连接
        /// </summary>
        private SqliteConnection m_dbConnection;

        /// <summary>
        /// 数据库文件名
        /// </summary>
        private readonly string m_dbPath;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="connect">数据库文件名</param>
        /// <param name="useMemory">事先把数据库加载到内存</param>
        public SqliteHelper(string connect, bool useMemory = false)
        {
            m_dbPath = connect;
            try
            {
                if (useMemory)
                {
                    m_dbConnection = new SqliteConnection("URI=file::memory:");
                    m_dbConnection.Open();
                    var command = m_dbConnection.CreateCommand();
                    command.CommandText = string.Format("ATTACH '{0}' AS filedb", connect);
                    command.ExecuteNonQuery();
                    command.Dispose();
                    Console.WriteLine("数据库[内存]已连接:" + m_dbPath);
                }
                else
                {
                    m_dbConnection = new SqliteConnection("data source=" + connect);
                    m_dbConnection.Open();
                    Console.WriteLine("数据库已连接:" + m_dbPath);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("数据库连接失败" + e);
            }
        }

        /// <summary>
        /// 执行SQL查询
        /// </summary>
        /// <param name="sql">SQL语句</param>
        /// <returns>查询结果</returns>
        public SqliteDataReader Query(string sql)
        {
            SqliteCommand dbCommand = m_dbConnection.CreateCommand();
            dbCommand.CommandText = sql;
            var ret = dbCommand.ExecuteReader();
            dbCommand.Dispose();
            return ret;
        }

        #region IDisposable Support
        private bool disposedValue = false; // 要检测冗余调用

        private void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing && m_dbConnection != null)
                {
                    m_dbConnection.Close();
                    m_dbConnection.Dispose();
                    Console.WriteLine("数据库连接已关闭:" + m_dbPath);
                }
                m_dbConnection = null;

                disposedValue = true;
            }
        }

        ~SqliteHelper()
        {
            // 请勿更改此代码。将清理代码放入以上 Dispose(bool disposing) 中。
            Dispose(false);
        }

        // 添加此代码以正确实现可处置模式。
        public void Dispose()
        {
            // 请勿更改此代码。将清理代码放入以上 Dispose(bool disposing) 中。
            Dispose(true);
            // 通知GC不再调用终结器
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
