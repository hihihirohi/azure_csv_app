using System;
using System.Data;
using System.Windows.Forms;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections.Generic;
using CsvHelper;
using CsvHelper.Configuration;

namespace AzureCsvApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            int count = customer_Update_Sql();
        }

        public class CsvData
        {
            public string Name { get; set; }
            public string Name1 { get; set; }
            public string Name2 { get; set; }
            public string Name3 { get; set; }
            public string Name4 { get; set; }
            public string Name5 { get; set; }
            public string Name6 { get; set; }
            public string Name7 { get; set; }
            public string Name8 { get; set; }
            public string Name9 { get; set; }
            public string Name10 { get; set; }
            public string Name11 { get; set; }
            public string Name12 { get; set; }
            public string Name13 { get; set; }
            public string Name14 { get; set; }
            public string Name15 { get; set; }
            public string Name16 { get; set; }
            public string Name17 { get; set; }
            public string Name18 { get; set; }
            public string Name19 { get; set; }
            public string Name20 { get; set; }
            public string Name21 { get; set; }
            public string Name22 { get; set; }
            public string Name23 { get; set; }
        }

        class CsvDataMap : CsvHelper.Configuration.ClassMap<CsvData>
        {
            public CsvDataMap()
            {
                Map(x => x.Name).Index(0);
                Map(x => x.Name1).Index(1);
                Map(x => x.Name2).Index(2);
                Map(x => x.Name3).Index(3);
                Map(x => x.Name4).Index(4);
                Map(x => x.Name5).Index(5);
                Map(x => x.Name6).Index(6);
                Map(x => x.Name7).Index(7);
                Map(x => x.Name8).Index(8);
                Map(x => x.Name9).Index(9);
                Map(x => x.Name10).Index(10);
                Map(x => x.Name11).Index(11);
                Map(x => x.Name12).Index(12);
                Map(x => x.Name13).Index(13);
                Map(x => x.Name14).Index(14);
                Map(x => x.Name15).Index(15);
                Map(x => x.Name16).Index(16);
                Map(x => x.Name17).Index(17);
                Map(x => x.Name18).Index(18);
                Map(x => x.Name19).Index(19);
                Map(x => x.Name20).Index(20);
                Map(x => x.Name21).Index(21);
                Map(x => x.Name22).Index(22);
                Map(x => x.Name23).Index(23);
            }
        }

        public int customer_Update_Sql()
        {
            int Count = 0;
            try
            {
                using (var connection = GetConnection("customer"))
                {
                    // データベースの接続開始
                    connection.Open();

                    // トランザクション開始
                    using (var transaction = connection.BeginTransaction())
                    {
                        try
                        {
                            using (var adapter = new SqlDataAdapter())
                            {
                                DateTime dtNow = DateTime.Now;
                                SqlCommand Cmd = new SqlCommand("", connection) { Transaction = transaction };
                                string query = "";

                                // クエリーを作成します。
                                query =
                                         "UPDATE Azure_details " +
                                         "SET SpCompanyName = @Name, " +
                                             "SpDomain = @Name1, " +
                                             "SpSubDomain = @Name2, " +
                                             "PurchaseCompanyCode = @Name3, " +
                                             "PurchaseCompanyName = @Name4, " +
                                             "StartUsagePeriod = @Name6, " +
                                             "FinishUsagePeriod = @Name7, " +
                                             "ElementName = @Name8, " +
                                             "CategoyName = @Name9, " +
                                             "SubCategoryName = @Name10, " +
                                             "Area = @Name11, " +
                                             "ResouceId = @Name12, " +
                                             "IdmSubsucriptionId = @Name13, " +
                                             "Tag = @Name14, " +
                                             "UnitName = @Name15, " +
                                             "InstanceDataResourceUri = @Name16, " +
                                             "InstanceDataPartNumber = @Name17, " +
                                             "InStanceDataOrderNumber = @Name18, " +
                                             "Domain = @Name19, " +
                                             "SubscriptionId = @Name20, " +
                                             "PurchaseValue = @Name21, " +
                                             "UsafeFee = @Name22, " +
                                             "PurchasePrice = @Name23 " +
                                             "WHERE DateOfAcquisition = @Name5 ";

                                System.IO.StreamReader sr = new System.IO.StreamReader(
                                                        @"C:\\csv\\azure.csv",
                                                        System.Text.Encoding.GetEncoding("shift_jis"));

                                using (var csv = new CsvReader(sr))

                                {
                                    var config = csv.Configuration;

                                    csv.Configuration.HasHeaderRecord = true;

                                    csv.Configuration.RegisterClassMap<CsvDataMap>();

                                    var records = csv.GetRecords<CsvData>();

                                    foreach (var n in records)
                                    {
                                        //Console.WriteLine("{0}|{1}|{2}|{3}|{4}", n.Name, n.Name1, n.Name2, n.Name3, n.Name4);
                                        //command.CommandType = CommandType.StoredProcedure;;
                                        #region パラメータ
                                        SqlParameter param1 = new SqlParameter();
                                        param1 = Cmd.CreateParameter();
                                        param1.ParameterName = "@Name";
                                        param1.SqlDbType = SqlDbType.VarChar;
                                        param1.Direction = ParameterDirection.Input;
                                        param1.Value = n.Name;
                                        Cmd.Parameters.Add(param1);

                                        SqlParameter param2 = new SqlParameter();
                                        param2 = Cmd.CreateParameter();
                                        param2.ParameterName = "@Name1";
                                        param2.SqlDbType = SqlDbType.VarChar;
                                        param2.Direction = ParameterDirection.Input;
                                        param2.Value = n.Name1;
                                        Cmd.Parameters.Add(param2);

                                        SqlParameter param3 = new SqlParameter();
                                        param3 = Cmd.CreateParameter();
                                        param3.ParameterName = "@Name2";
                                        param3.SqlDbType = SqlDbType.VarChar;
                                        param3.Direction = ParameterDirection.Input;
                                        param3.Value = n.Name2;
                                        Cmd.Parameters.Add(param3);

                                        SqlParameter param4 = new SqlParameter();
                                        param4 = Cmd.CreateParameter();
                                        param4.ParameterName = "@Name3";
                                        param4.SqlDbType = SqlDbType.VarChar;
                                        param4.Direction = ParameterDirection.Input;
                                        param4.Value = n.Name3;
                                        Cmd.Parameters.Add(param4);

                                        SqlParameter param5 = new SqlParameter();
                                        param5 = Cmd.CreateParameter();
                                        param5.ParameterName = "@Name4";
                                        param5.SqlDbType = SqlDbType.VarChar;
                                        param5.Direction = ParameterDirection.Input;
                                        param5.Value = n.Name4;
                                        Cmd.Parameters.Add(param5);

                                        SqlParameter param6 = new SqlParameter();
                                        param6 = Cmd.CreateParameter();
                                        param6.ParameterName = "@Name5";
                                        param6.SqlDbType = SqlDbType.VarChar;
                                        param6.Direction = ParameterDirection.Input;
                                        param6.Value = n.Name5;
                                        Cmd.Parameters.Add(param6);

                                        SqlParameter param7 = new SqlParameter();
                                        param7 = Cmd.CreateParameter();
                                        param7.ParameterName = "@Name6";
                                        param7.SqlDbType = SqlDbType.DateTime;
                                        param7.Direction = ParameterDirection.Input;
                                        param7.Value = n.Name6;
                                        Cmd.Parameters.Add(param7);

                                        SqlParameter param8 = new SqlParameter();
                                        param8 = Cmd.CreateParameter();
                                        param8.ParameterName = "@Name7";
                                        param8.SqlDbType = SqlDbType.DateTime;
                                        param8.Direction = ParameterDirection.Input;
                                        param8.Value = n.Name7;
                                        Cmd.Parameters.Add(param8);

                                        SqlParameter param9 = new SqlParameter();
                                        param9 = Cmd.CreateParameter();
                                        param9.ParameterName = "@Name8";
                                        param9.SqlDbType = SqlDbType.VarChar;
                                        param9.Direction = ParameterDirection.Input;
                                        param9.Value = n.Name8;
                                        Cmd.Parameters.Add(param9);

                                        SqlParameter param10 = new SqlParameter();
                                        param10 = Cmd.CreateParameter();
                                        param10.ParameterName = "@Name9";
                                        param10.SqlDbType = SqlDbType.VarChar;
                                        param10.Direction = ParameterDirection.Input;
                                        param10.Value = n.Name9;
                                        Cmd.Parameters.Add(param10);

                                        SqlParameter param11 = new SqlParameter();
                                        param11 = Cmd.CreateParameter();
                                        param11.ParameterName = "@Name10";
                                        param11.SqlDbType = SqlDbType.VarChar;
                                        param11.Direction = ParameterDirection.Input;
                                        param11.Value = n.Name10;
                                        Cmd.Parameters.Add(param11);

                                        SqlParameter param12 = new SqlParameter();
                                        param12 = Cmd.CreateParameter();
                                        param12.ParameterName = "@Name11";
                                        param12.SqlDbType = SqlDbType.VarChar;
                                        param12.Direction = ParameterDirection.Input;
                                        param12.Value = n.Name11;
                                        Cmd.Parameters.Add(param12);

                                        SqlParameter param13 = new SqlParameter();
                                        param13 = Cmd.CreateParameter();
                                        param13.ParameterName = "@Name12";
                                        param13.SqlDbType = SqlDbType.VarChar;
                                        param13.Direction = ParameterDirection.Input;
                                        param13.Value = n.Name12;
                                        Cmd.Parameters.Add(param13);

                                        SqlParameter param14 = new SqlParameter();
                                        param14 = Cmd.CreateParameter();
                                        param14.ParameterName = "@Name13";
                                        param14.SqlDbType = SqlDbType.VarChar;
                                        param14.Direction = ParameterDirection.Input;
                                        param14.Value = n.Name13;
                                        Cmd.Parameters.Add(param14);

                                        SqlParameter param15 = new SqlParameter();
                                        param15 = Cmd.CreateParameter();
                                        param15.ParameterName = "@Name14";
                                        param15.SqlDbType = SqlDbType.VarChar;
                                        param15.Direction = ParameterDirection.Input;
                                        param15.Value = n.Name14;
                                        Cmd.Parameters.Add(param15);

                                        SqlParameter param16 = new SqlParameter();
                                        param16 = Cmd.CreateParameter();
                                        param16.ParameterName = "@Name15";
                                        param16.SqlDbType = SqlDbType.VarChar;
                                        param16.Direction = ParameterDirection.Input;
                                        param16.Value = n.Name15;
                                        Cmd.Parameters.Add(param16);

                                        SqlParameter param17 = new SqlParameter();
                                        param17 = Cmd.CreateParameter();
                                        param17.ParameterName = "@Name16";
                                        param17.SqlDbType = SqlDbType.VarChar;
                                        param17.Direction = ParameterDirection.Input;
                                        param17.Value = n.Name16;
                                        Cmd.Parameters.Add(param17);

                                        SqlParameter param18 = new SqlParameter();
                                        param18 = Cmd.CreateParameter();
                                        param18.ParameterName = "@Name17";
                                        param18.SqlDbType = SqlDbType.VarChar;
                                        param18.Direction = ParameterDirection.Input;
                                        param18.Value = n.Name17;
                                        Cmd.Parameters.Add(param18);

                                        SqlParameter param19 = new SqlParameter();
                                        param19 = Cmd.CreateParameter();
                                        param19.ParameterName = "@Name18";
                                        param19.SqlDbType = SqlDbType.VarChar;
                                        param19.Direction = ParameterDirection.Input;
                                        param19.Value = n.Name18;
                                        Cmd.Parameters.Add(param19);

                                        SqlParameter param20 = new SqlParameter();
                                        param20 = Cmd.CreateParameter();
                                        param20.ParameterName = "@Name19";
                                        param20.SqlDbType = SqlDbType.VarChar;
                                        param20.Direction = ParameterDirection.Input;
                                        param20.Value = n.Name19;
                                        Cmd.Parameters.Add(param20);

                                        SqlParameter param21 = new SqlParameter();
                                        param21 = Cmd.CreateParameter();
                                        param21.ParameterName = "@Name20";
                                        param21.SqlDbType = SqlDbType.VarChar;
                                        param21.Direction = ParameterDirection.Input;
                                        param21.Value = n.Name20;
                                        Cmd.Parameters.Add(param21);

                                        SqlParameter param22 = new SqlParameter();
                                        param22 = Cmd.CreateParameter();
                                        param22.ParameterName = "@Name21";
                                        param22.SqlDbType = SqlDbType.VarChar;
                                        param22.Direction = ParameterDirection.Input;
                                        param22.Value = n.Name21;
                                        Cmd.Parameters.Add(param22);

                                        SqlParameter param23 = new SqlParameter();
                                        param23 = Cmd.CreateParameter();
                                        param23.ParameterName = "@Name22";
                                        param23.SqlDbType = SqlDbType.VarChar;
                                        param23.Direction = ParameterDirection.Input;
                                        param23.Value = n.Name22;
                                        Cmd.Parameters.Add(param23);

                                        SqlParameter param24 = new SqlParameter();
                                        param24 = Cmd.CreateParameter();
                                        param24.ParameterName = "@Name23";
                                        param24.SqlDbType = SqlDbType.VarChar;
                                        param24.Direction = ParameterDirection.Input;
                                        param24.Value = n.Name23;
                                        Cmd.Parameters.Add(param24);
                                        #endregion

                                        Cmd.CommandText = query;
                                        adapter.UpdateCommand = Cmd;

                                        Count = adapter.UpdateCommand.ExecuteNonQuery();

                                        if (Count <  0)
                                        {
                                            // 更新成功
                                        }
                                        else
                                        {
                                            // 追加処理を実行します。
                                            customer_Insert_Sql(n.Name, n.Name1, n.Name2, n.Name3, n.Name4, n.Name5, n.Name6, n.Name7, n.Name8, n.Name9, n.Name10, n.Name11, n.Name12, n.Name13, n.Name14, n.Name15, n.Name16, n.Name17, n.Name18, n.Name19, n.Name20, n.Name21, n.Name22, n.Name23);
                                        }

                                        Cmd.Parameters.Clear();
                                    }
                                    transaction.Commit();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // ロールバック
                            transaction.Rollback();
                            throw (ex);
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
            return Count;
        }

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


                public int customer_Insert_Sql(string Name, string Name1, string Name2, string Name3, string Name4, string Name5, string Name6, string Name7, string Name8, string Name9, string Name10, string Name11, string Name12, string Name13, string Name14, string Name15, string Name16, string Name17, string Name18, string Name19, string Name20, string Name21, string Name22, string Name23)
        {
            int Count = 0;
            try
            {
                using (var connection = GetConnection("customer"))
                {
                    // データベースの接続開始
                    connection.Open();

                    // トランザクション開始
                    using (var transaction = connection.BeginTransaction())
                    {                                                                                                                                                     
                        try
                        {
                            using (var adapter = new SqlDataAdapter())
                            {                                
                                DateTime dtNow = DateTime.Now;
                                SqlCommand Cmd = new SqlCommand("", connection) { Transaction = transaction };
                                string query = "";
                                query = "insert into Azure_details (SpCompanyName,SpDomain,SpSubDomain,PurchaseCompanyCode,PurchaseCompanyName,DateOfAcquisition,StartUsagePeriod,FinishUsagePeriod,ElementName,CategoyName,SubCategoryName,Area,ResouceId,IdmSubsucriptionId,Tag,UnitName,InstanceDataResourceUri,InstanceDataPartNumber,InStanceDataOrderNumber,Domain,SubscriptionId,PurchaseValue,UsafeFee,PurchasePrice) values (@Name,@Name1,@Name2,@Name3,@Name4,@Name5,@Name6,@Name7,@Name8,@Name9,@Name10,@Name11,@Name12,@Name13,@Name14,@Name15,@Name16,@Name17,@Name18,@Name19,@Name20,@Name21,@Name22,@Name23)";
                                #region パラメータ
                                SqlParameter param1 = new SqlParameter();
                                param1 = Cmd.CreateParameter();
                                param1.ParameterName = "@Name";
                                param1.SqlDbType = SqlDbType.VarChar;
                                param1.Direction = ParameterDirection.Input;
                                param1.Value = Name;
                                Cmd.Parameters.Add(param1);

                                SqlParameter param2 = new SqlParameter();
                                param2 = Cmd.CreateParameter();
                                param2.ParameterName = "@Name1";
                                param2.SqlDbType = SqlDbType.VarChar;
                                param2.Direction = ParameterDirection.Input;
                                param2.Value = Name1;
                                Cmd.Parameters.Add(param2);

                                SqlParameter param3 = new SqlParameter();
                                param3 = Cmd.CreateParameter();
                                param3.ParameterName = "@Name2";
                                param3.SqlDbType = SqlDbType.VarChar;
                                param3.Direction = ParameterDirection.Input;
                                param3.Value = Name2;
                                Cmd.Parameters.Add(param3);

                                SqlParameter param4 = new SqlParameter();
                                param4 = Cmd.CreateParameter();
                                param4.ParameterName = "@Name3";
                                param4.SqlDbType = SqlDbType.VarChar;
                                param4.Direction = ParameterDirection.Input;
                                param4.Value = Name3;
                                Cmd.Parameters.Add(param4);

                                SqlParameter param5 = new SqlParameter();
                                param5 = Cmd.CreateParameter();
                                param5.ParameterName = "@Name4";
                                param5.SqlDbType = SqlDbType.VarChar;
                                param5.Direction = ParameterDirection.Input;
                                param5.Value = Name4;
                                Cmd.Parameters.Add(param5);

                                SqlParameter param6 = new SqlParameter();
                                param6 = Cmd.CreateParameter();
                                param6.ParameterName = "@Name5";
                                param6.SqlDbType = SqlDbType.VarChar;
                                param6.Direction = ParameterDirection.Input;
                                param6.Value = Name5;
                                Cmd.Parameters.Add(param6);
 
                                SqlParameter param7 = new SqlParameter();
                                param7 = Cmd.CreateParameter();
                                param7.ParameterName = "@Name6";
                                param7.SqlDbType = SqlDbType.DateTime;
                                param7.Direction = ParameterDirection.Input;
                                param7.Value = Name6;
                                Cmd.Parameters.Add(param7);

                                SqlParameter param8 = new SqlParameter();
                                param8 = Cmd.CreateParameter();
                                param8.ParameterName = "@Name7";
                                param8.SqlDbType = SqlDbType.DateTime;
                                param8.Direction = ParameterDirection.Input;
                                param8.Value = Name7;
                                Cmd.Parameters.Add(param8);

                                SqlParameter param9 = new SqlParameter();
                                param9 = Cmd.CreateParameter();
                                param9.ParameterName = "@Name8";
                                param9.SqlDbType = SqlDbType.VarChar;
                                param9.Direction = ParameterDirection.Input;
                                param9.Value = Name8;
                                Cmd.Parameters.Add(param9);

                                SqlParameter param10 = new SqlParameter();
                                param10 = Cmd.CreateParameter();
                                param10.ParameterName = "@Name9";
                                param10.SqlDbType = SqlDbType.VarChar;
                                param10.Direction = ParameterDirection.Input;
                                param10.Value = Name9;
                                Cmd.Parameters.Add(param10);

                                SqlParameter param11 = new SqlParameter();
                                param11 = Cmd.CreateParameter();
                                param11.ParameterName = "@Name10";
                                param11.SqlDbType = SqlDbType.VarChar;
                                param11.Direction = ParameterDirection.Input;
                                param11.Value = Name10;
                                Cmd.Parameters.Add(param11);

                                SqlParameter param12 = new SqlParameter();
                                param12 = Cmd.CreateParameter();
                                param12.ParameterName = "@Name11";
                                param12.SqlDbType = SqlDbType.VarChar;
                                param12.Direction = ParameterDirection.Input;
                                param12.Value = Name11;
                                Cmd.Parameters.Add(param12);

                                SqlParameter param13 = new SqlParameter();
                                param13 = Cmd.CreateParameter();
                                param13.ParameterName = "@Name12";
                                param13.SqlDbType = SqlDbType.VarChar;
                                param13.Direction = ParameterDirection.Input;
                                param13.Value = Name12;
                                Cmd.Parameters.Add(param13);

                                SqlParameter param14 = new SqlParameter();
                                param14 = Cmd.CreateParameter();
                                param14.ParameterName = "@Name13";
                                param14.SqlDbType = SqlDbType.VarChar;
                                param14.Direction = ParameterDirection.Input;
                                param14.Value = Name13;
                                Cmd.Parameters.Add(param14);

                                SqlParameter param15 = new SqlParameter();
                                param15 = Cmd.CreateParameter();
                                param15.ParameterName = "@Name14";
                                param15.SqlDbType = SqlDbType.VarChar;
                                param15.Direction = ParameterDirection.Input;
                                param15.Value = Name14;
                                Cmd.Parameters.Add(param15);
                                
                                SqlParameter param16 = new SqlParameter();
                                param16 = Cmd.CreateParameter();
                                param16.ParameterName = "@Name15";
                                param16.SqlDbType = SqlDbType.VarChar;
                                param16.Direction = ParameterDirection.Input;
                                param16.Value = Name15;
                                Cmd.Parameters.Add(param16);

                                SqlParameter param17 = new SqlParameter();
                                param17 = Cmd.CreateParameter();
                                param17.ParameterName = "@Name16";
                                param17.SqlDbType = SqlDbType.VarChar;
                                param17.Direction = ParameterDirection.Input;
                                param17.Value = Name16;
                                Cmd.Parameters.Add(param17);

                                SqlParameter param18 = new SqlParameter();
                                param18 = Cmd.CreateParameter();
                                param18.ParameterName = "@Name17";
                                param18.SqlDbType = SqlDbType.VarChar;
                                param18.Direction = ParameterDirection.Input;
                                param18.Value = Name17;
                                Cmd.Parameters.Add(param18);

                                SqlParameter param19 = new SqlParameter();
                                param19 = Cmd.CreateParameter();
                                param19.ParameterName = "@Name18";
                                param19.SqlDbType = SqlDbType.VarChar;
                                param19.Direction = ParameterDirection.Input;
                                param19.Value = Name18;
                                Cmd.Parameters.Add(param19);

                                SqlParameter param20 = new SqlParameter();
                                param20 = Cmd.CreateParameter();
                                param20.ParameterName = "@Name19";
                                param20.SqlDbType = SqlDbType.VarChar;
                                param20.Direction = ParameterDirection.Input;
                                param20.Value = Name19;
                                Cmd.Parameters.Add(param20);

                                SqlParameter param21 = new SqlParameter();
                                param21 = Cmd.CreateParameter();
                                param21.ParameterName = "@Name20";
                                param21.SqlDbType = SqlDbType.VarChar;
                                param21.Direction = ParameterDirection.Input;
                                param21.Value = Name20;
                                Cmd.Parameters.Add(param21);

                                SqlParameter param22 = new SqlParameter();
                                param22 = Cmd.CreateParameter();
                                param22.ParameterName = "@Name21";
                                param22.SqlDbType = SqlDbType.VarChar;
                                param22.Direction = ParameterDirection.Input;
                                param22.Value = Name21;
                                Cmd.Parameters.Add(param22);

                                SqlParameter param23 = new SqlParameter();
                                param23 = Cmd.CreateParameter();
                                param23.ParameterName = "@Name22";
                                param23.SqlDbType = SqlDbType.VarChar;
                                param23.Direction = ParameterDirection.Input;
                                param23.Value = Name22;
                                Cmd.Parameters.Add(param23);

                                SqlParameter param24 = new SqlParameter();
                                param24 = Cmd.CreateParameter();
                                param24.ParameterName = "@Name23";
                                param24.SqlDbType = SqlDbType.VarChar;
                                param24.Direction = ParameterDirection.Input;
                                param24.Value = Name23;
                                Cmd.Parameters.Add(param24);
                                #endregion
                                Cmd.CommandText = query;
                                        adapter.InsertCommand = Cmd;

                                        Count += adapter.InsertCommand.ExecuteNonQuery();
                                        
                                        Cmd.Parameters.Clear();                                    
                                    transaction.Commit();
                            }
                        }
                        catch (Exception ex)
                        {
                            // ロールバック
                            transaction.Rollback();
                            throw (ex);
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
            return Count;
        }


    }
}










































//        public int customer_Insert_Sql(string Name, string Name1, string Name2, string Name3, string Name4)
//        {
//            int Count = 0;
//            try
//            {
//                using (var connection = GetConnection("customer"))
//                {
//                    // データベースの接続開始
//                    connection.Open();

//                    // トランザクション開始
//                    using (var transaction = connection.BeginTransaction())
//                    {
//                        try
//                        {
//                            using (var adapter = new SqlDataAdapter())
//                            {
//                                DateTime dtNow = DateTime.Now;
//                                SqlCommand Cmd = new SqlCommand("", connection) { Transaction = transaction };
//                                string query = "";

//                                // クエリーを作成します。
//                                query =
//                                         "insert into customer (SpCompanyName,SpDomain,SpSubDomain,PurchaseCompanyCode,PurchaseCompanyName,DateOfAcquisition,StartUsagePeriod,FinidhUsagePeriod,ElementName,CategoryName,SubCategoryName,Area,ResourceUri,InstanceDataResourcsUri,InstanceDataPartNumber,InstanceDateOrderNumber,Domain,SubscriptionId,PurchaseValue,UsafeFee,PurchaseValue,UsafeFee,PurchasePrice) values (@Name,@Name1,@Name2,@Name3,@Name4,@Name5,@Name6,@Name7,@Name8,@Name9,@Name10,@Name11,@Name12,@Name13,@Name14,@Name15,@Name16,@Name17,@Name18,@Name19,@Name20,@Name21,@Name22,@Name23,@Name24)";
//                                #region パラメータ
//                                SqlParameter param1 = new SqlParameter();
//                                param1 = Cmd.CreateParameter();
//                                param1.ParameterName = "@Name";
//                                param1.SqlDbType = SqlDbType.NVarChar;
//                                param1.Direction = ParameterDirection.Input;
//                                param1.Value = Name;
//                                Cmd.Parameters.Add(param1);

//                                SqlParameter param2 = new SqlParameter();
//                                param2 = Cmd.CreateParameter();
//                                param2.ParameterName = "@Name1";
//                                param2.SqlDbType = SqlDbType.NVarChar;
//                                param2.Direction = ParameterDirection.Input;
//                                param2.Value = Name1;
//                                Cmd.Parameters.Add(param2);

//                                SqlParameter param3 = new SqlParameter();
//                                param3 = Cmd.CreateParameter();
//                                param3.ParameterName = "@Name2";
//                                param3.SqlDbType = SqlDbType.NVarChar;
//                                param3.Direction = ParameterDirection.Input;
//                                param3.Value = Name2;
//                                Cmd.Parameters.Add(param3);

//                                SqlParameter param4 = new SqlParameter();
//                                param4 = Cmd.CreateParameter();
//                                param4.ParameterName = "@Name3";
//                                param4.SqlDbType = SqlDbType.NVarChar;
//                                param4.Direction = ParameterDirection.Input;
//                                param4.Value = Name3;
//                                Cmd.Parameters.Add(param4);

//                                SqlParameter param5 = new SqlParameter();
//                                param5 = Cmd.CreateParameter();
//                                param5.ParameterName = "@Name4";
//                                param5.SqlDbType = SqlDbType.NVarChar;
//                                param5.Direction = ParameterDirection.Input;
//                                param5.Value = Name4;
//                                Cmd.Parameters.Add(param5);

//                                #region パラメータ
//                                SqlParameter param6 = new SqlParameter();
//                                param1 = Cmd.CreateParameter();
//                                param1.ParameterName = "@Name";
//                                param1.SqlDbType = SqlDbType.NVarChar;
//                                param1.Direction = ParameterDirection.Input;
//                                param1.Value = Name;
//                                Cmd.Parameters.Add(param1);

//                                SqlParameter param7 = new SqlParameter();
//                                param2 = Cmd.CreateParameter();
//                                param2.ParameterName = "@Name1";
//                                param2.SqlDbType = SqlDbType.NVarChar;
//                                param2.Direction = ParameterDirection.Input;
//                                param2.Value = Name1;
//                                Cmd.Parameters.Add(param2);

//                                SqlParameter param8 = new SqlParameter();
//                                param3 = Cmd.CreateParameter();
//                                param3.ParameterName = "@Name2";
//                                param3.SqlDbType = SqlDbType.NVarChar;
//                                param3.Direction = ParameterDirection.Input;
//                                param3.Value = Name2;
//                                Cmd.Parameters.Add(param3);

//                                SqlParameter param9 = new SqlParameter();
//                                param4 = Cmd.CreateParameter();
//                                param4.ParameterName = "@Name3";
//                                param4.SqlDbType = SqlDbType.NVarChar;
//                                param4.Direction = ParameterDirection.Input;
//                                param4.Value = Name3;
//                                Cmd.Parameters.Add(param4);

//                                SqlParameter param10 = new SqlParameter();
//                                param5 = Cmd.CreateParameter();
//                                param5.ParameterName = "@Name4";
//                                param5.SqlDbType = SqlDbType.NVarChar;
//                                param5.Direction = ParameterDirection.Input;
//                                param5.Value = Name4;
//                                Cmd.Parameters.Add(param5);

//                                #region パラメータ
//                                SqlParameter param11= new SqlParameter();
//                                param1 = Cmd.CreateParameter();
//                                param1.ParameterName = "@Name";
//                                param1.SqlDbType = SqlDbType.NVarChar;
//                                param1.Direction = ParameterDirection.Input;
//                                param1.Value = Name;
//                                Cmd.Parameters.Add(param1);

//                                SqlParameter param12 = new SqlParameter();
//                                param2 = Cmd.CreateParameter();
//                                param2.ParameterName = "@Name1";
//                                param2.SqlDbType = SqlDbType.NVarChar;
//                                param2.Direction = ParameterDirection.Input;
//                                param2.Value = Name1;
//                                Cmd.Parameters.Add(param2);

//                                SqlParameter param13 = new SqlParameter();
//                                param3 = Cmd.CreateParameter();
//                                param3.ParameterName = "@Name2";
//                                param3.SqlDbType = SqlDbType.NVarChar;
//                                param3.Direction = ParameterDirection.Input;
//                                param3.Value = Name2;
//                                Cmd.Parameters.Add(param3);

//                                SqlParameter param14= new SqlParameter();
//                                param4 = Cmd.CreateParameter();
//                                param4.ParameterName = "@Name3";
//                                param4.SqlDbType = SqlDbType.NVarChar;
//                                param4.Direction = ParameterDirection.Input;
//                                param4.Value = Name3;
//                                Cmd.Parameters.Add(param4);

//                                SqlParameter param15 = new SqlParameter();
//                                param5 = Cmd.CreateParameter();
//                                param5.ParameterName = "@Name4";
//                                param5.SqlDbType = SqlDbType.NVarChar;
//                                param5.Direction = ParameterDirection.Input;
//                                param5.Value = Name4;
//                                Cmd.Parameters.Add(param5);

//                                #region パラメータ
//                                SqlParameter param16 = new SqlParameter();
//                                param1 = Cmd.CreateParameter();
//                                param1.ParameterName = "@Name";
//                                param1.SqlDbType = SqlDbType.NVarChar;
//                                param1.Direction = ParameterDirection.Input;
//                                param1.Value = Name;
//                                Cmd.Parameters.Add(param1);

//                                SqlParameter param17 = new SqlParameter();
//                                param2 = Cmd.CreateParameter();
//                                param2.ParameterName = "@Name1";
//                                param2.SqlDbType = SqlDbType.NVarChar;
//                                param2.Direction = ParameterDirection.Input;
//                                param2.Value = Name1;
//                                Cmd.Parameters.Add(param2);

//                                SqlParameter param18 = new SqlParameter();
//                                param3 = Cmd.CreateParameter();
//                                param3.ParameterName = "@Name2";
//                                param3.SqlDbType = SqlDbType.NVarChar;
//                                param3.Direction = ParameterDirection.Input;
//                                param3.Value = Name2;
//                                Cmd.Parameters.Add(param3);

//                                SqlParameter param19 = new SqlParameter();
//                                param4 = Cmd.CreateParameter();
//                                param4.ParameterName = "@Name3";
//                                param4.SqlDbType = SqlDbType.NVarChar;
//                                param4.Direction = ParameterDirection.Input;
//                                param4.Value = Name3;
//                                Cmd.Parameters.Add(param4);

//                                SqlParameter param20 = new SqlParameter();
//                                param5 = Cmd.CreateParameter();
//                                param5.ParameterName = "@Name4";
//                                param5.SqlDbType = SqlDbType.NVarChar;
//                                param5.Direction = ParameterDirection.Input;
//                                param5.Value = Name4;
//                                Cmd.Parameters.Add(param5);

//                                SqlParameter param21 = new SqlParameter();
//                                param1 = Cmd.CreateParameter();
//                                param1.ParameterName = "@Name";
//                                param1.SqlDbType = SqlDbType.NVarChar;
//                                param1.Direction = ParameterDirection.Input;
//                                param1.Value = Name;
//                                Cmd.Parameters.Add(param1);

//                                SqlParameter param22 = new SqlParameter();
//                                param2 = Cmd.CreateParameter();
//                                param2.ParameterName = "@Name1";
//                                param2.SqlDbType = SqlDbType.NVarChar;
//                                param2.Direction = ParameterDirection.Input;
//                                param2.Value = Name1;
//                                Cmd.Parameters.Add(param2);

//                                SqlParameter param23 = new SqlParameter();
//                                param3 = Cmd.CreateParameter();
//                                param3.ParameterName = "@Name2";
//                                param3.SqlDbType = SqlDbType.NVarChar;
//                                param3.Direction = ParameterDirection.Input;
//                                param3.Value = Name2;
//                                Cmd.Parameters.Add(param3);

//                                SqlParameter param24 = new SqlParameter();
//                                param4 = Cmd.CreateParameter();
//                                param4.ParameterName = "@Name3";
//                                param4.SqlDbType = SqlDbType.NVarChar;
//                                param4.Direction = ParameterDirection.Input;
//                                param4.Value = Name3;
//                                Cmd.Parameters.Add(param4);
//                                #endregion

//                                Cmd.CommandText = query;
//                                adapter.InsertCommand = Cmd;

//                                Count += adapter.InsertCommand.ExecuteNonQuery();

//                                Cmd.Parameters.Clear();
//                                transaction.Commit();

//                            }
//                        }
//                        catch (Exception ex)
//                        {
//                            // ロールバック
//                            transaction.Rollback();
//                            throw (ex);
//                        }
//                        finally
//                        {
//                            connection.Close();
//                        }
//                    }

//                }
//            }
//            catch (Exception ex)
//            {
//                throw (ex);
//            }
//            return Count;
//        }

//        public static SqlConnection GetConnection(string settingName)
//        {
//            SqlConnection con = null;

//            // 接続文字列をApp.configから取得します
//            string connectionString = ConfigurationManager.ConnectionStrings[settingName].ConnectionString;

//            if (String.IsNullOrEmpty(connectionString))
//            {
//                throw new NullReferenceException("接続文字列が設定されていません。");
//            }
//            else
//            {
//                con = new SqlConnection(connectionString);
//            }

//            return con;
//        }
//    }

//}