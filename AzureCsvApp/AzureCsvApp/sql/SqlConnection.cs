using System;
using System.Data;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;


using System.Threading.Tasks;

namespace AzureCsvApp.sql
{
    public class SqlConection
    {
        public static SqlConnection GetConnection(string settingName)
        {
            SqlConnection con = null;

            // 接続文字列をApp.configから取得します
            string connectionString = ConfigurationManager.ConnectionStrings[settingName].ConnectionString;

            if (String.IsNullOrEmpty(connectionString))
            {
                throw new NullReferenceException("接続文字列が設定されていません。");
            }
            else
            {
                con = new SqlConnection(connectionString);
            }

            return con;
        }
    }
}
