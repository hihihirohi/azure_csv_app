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
    public class SumPrice
    {
        public static void Output()
        {
            try
            {
                using (var conection = SqlConection.GetConnection("customer"))
                {
                    // データベースの接続開始
                    conection.Open();
                    {
                        try
                        {
                            {
                                string query = "SELECT SUM((PurchasePrice)) AS sumprice " +
                                                "FROM Azure_details " +
                                                "WHERE DateOfAcquisition = DateOfAcquisition ";

                                SqlCommand com = new SqlCommand(query, conection);
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
                            conection.Close();
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