using System;
using System.Data;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;

using AzureCsvApp.sql;

using System.Threading.Tasks;

namespace AzureCsvApp
{
    public class Unit
    {

        public static void Output()
        {
            try
            {
                using (var connection = SqlConection.GetConnection("customer"))
                {
                    // データベースの接続開始
                    connection.Open();

                    {
                        try
                        {
                            {
                                string query = "SELECT SUM((PurchasePrice)) AS sumprice " +
                                                "FROM Azure_details " +
                                                "WHERE DateOfAcquisition = DateOfAcquisition ";

                                SqlCommand com = new SqlCommand(query, connection);
                                SqlDataReader sdr = com.ExecuteReader();
                                //int uId_Or = com.GetOrdinal("PurchasePrice");

                                while (sdr.Read() == true)
                                {
                                    //Decimal uId = sdr.GetDecimal(PurchasePrice);
                                    var SumPrice = sdr["sumprice"].ToString();

                                    MessageBox.Show(SumPrice);

                                    MessageBox.Show("全体の合計の値を実行しました");
                                }


                                MessageBox.Show("処理が終了しました。");
                            }
                        }
                        finally
                        {
                            connection.Close();
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }
        }



    }
}
