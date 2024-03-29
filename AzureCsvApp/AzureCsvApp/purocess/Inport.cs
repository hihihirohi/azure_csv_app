﻿using System;
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

using AzureCsvApp.sql;

namespace AzureCsvApp
{
    public class InportData
    {
        public static void Inport()
        {
            int count = customer_Update_Sql();
        

            int customer_Update_Sql()
            {
                int Count = 0;
                try
                {
                    using (var connection = SqlConection.GetConnection("customer"))
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
                                             "SELECT * " +
                                             "FROM  Azure_details " +
                                             "WHERE DateOfAcquisition = @DateofAcquisition  ";

                                    System.IO.StreamReader sr = new System.IO.StreamReader(
                                                            @"C:\\csv\\azure.csv",
                                                            System.Text.Encoding.GetEncoding("shift_jis"));

                                    using (var csv = new CsvReader(sr))

                                    {
                                        var config = csv.Configuration;

                                        csv.Configuration.HasHeaderRecord = true;

                                        csv.Configuration.RegisterClassMap<CsvDataMap>();

                                        var records = csv.GetRecords<CsvData>();
                                        int loop_Count = 0;
                                        foreach (var n in records)
                                        {
                                            if (loop_Count == 0)
                                            {
                                                #region パラメータ
                                                SqlParameter param1 = new SqlParameter();
                                                param1 = Cmd.CreateParameter();
                                                param1.ParameterName = "@SpCompanyName";
                                                param1.SqlDbType = SqlDbType.NVarChar;
                                                param1.Direction = ParameterDirection.Input;
                                                param1.Value = n.SpCompanyName;
                                                Cmd.Parameters.Add(param1);

                                                SqlParameter param2 = new SqlParameter();
                                                param2 = Cmd.CreateParameter();
                                                param2.ParameterName = "@SpDomain";
                                                param2.SqlDbType = SqlDbType.VarChar;
                                                param2.Direction = ParameterDirection.Input;
                                                param2.Value = n.SpDomain;
                                                Cmd.Parameters.Add(param2);

                                                SqlParameter param3 = new SqlParameter();
                                                param3 = Cmd.CreateParameter();
                                                param3.ParameterName = "@SpSubDomain";
                                                param3.SqlDbType = SqlDbType.VarChar;
                                                param3.Direction = ParameterDirection.Input;
                                                param3.Value = n.SpSubDomain;
                                                Cmd.Parameters.Add(param3);

                                                SqlParameter param4 = new SqlParameter();
                                                param4 = Cmd.CreateParameter();
                                                param4.ParameterName = "@PurchaseCompanyCode";
                                                param4.SqlDbType = SqlDbType.VarChar;
                                                param4.Direction = ParameterDirection.Input;
                                                param4.Value = n.PurchaseCompanyCode;
                                                Cmd.Parameters.Add(param4);

                                                SqlParameter param5 = new SqlParameter();
                                                param5 = Cmd.CreateParameter();
                                                param5.ParameterName = "@PurchaseCompanyName";
                                                param5.SqlDbType = SqlDbType.VarChar;
                                                param5.Direction = ParameterDirection.Input;
                                                param5.Value = n.PurchaseCompanyName;
                                                Cmd.Parameters.Add(param5);

                                                SqlParameter param6 = new SqlParameter();
                                                param6 = Cmd.CreateParameter();
                                                param6.ParameterName = "@DateofAcquisition";
                                                param6.SqlDbType = SqlDbType.VarChar;
                                                param6.Direction = ParameterDirection.Input;
                                                param6.Value = n.DateofAcquisition;
                                                Cmd.Parameters.Add(param6);

                                                SqlParameter param7 = new SqlParameter();
                                                param7 = Cmd.CreateParameter();
                                                param7.ParameterName = "@StartUsagePeriod";
                                                param7.SqlDbType = SqlDbType.DateTime;
                                                param7.Direction = ParameterDirection.Input;
                                                param7.Value = n.StartUsagePeriod;
                                                Cmd.Parameters.Add(param7);

                                                SqlParameter param8 = new SqlParameter();
                                                param8 = Cmd.CreateParameter();
                                                param8.ParameterName = "@FinishUsagePeriod";
                                                param8.SqlDbType = SqlDbType.DateTime;
                                                param8.Direction = ParameterDirection.Input;
                                                param8.Value = n.FinishUsagePeriod;
                                                Cmd.Parameters.Add(param8);

                                                SqlParameter param9 = new SqlParameter();
                                                param9 = Cmd.CreateParameter();
                                                param9.ParameterName = "@ElementName";
                                                param9.SqlDbType = SqlDbType.VarChar;
                                                param9.Direction = ParameterDirection.Input;
                                                param9.Value = n.ElementName;
                                                Cmd.Parameters.Add(param9);

                                                SqlParameter param10 = new SqlParameter();
                                                param10 = Cmd.CreateParameter();
                                                param10.ParameterName = "@CategoyName";
                                                param10.SqlDbType = SqlDbType.VarChar;
                                                param10.Direction = ParameterDirection.Input;
                                                param10.Value = n.CategoyName;
                                                Cmd.Parameters.Add(param10);

                                                SqlParameter param11 = new SqlParameter();
                                                param11 = Cmd.CreateParameter();
                                                param11.ParameterName = "@SubCategoryName";
                                                param11.SqlDbType = SqlDbType.VarChar;
                                                param11.Direction = ParameterDirection.Input;
                                                param11.Value = n.SubCategoryName;
                                                Cmd.Parameters.Add(param11);

                                                SqlParameter param12 = new SqlParameter();
                                                param12 = Cmd.CreateParameter();
                                                param12.ParameterName = "@Area";
                                                param12.SqlDbType = SqlDbType.VarChar;
                                                param12.Direction = ParameterDirection.Input;
                                                param12.Value = n.Area;
                                                Cmd.Parameters.Add(param12);

                                                SqlParameter param13 = new SqlParameter();
                                                param13 = Cmd.CreateParameter();
                                                param13.ParameterName = "@ResouceId";
                                                param13.SqlDbType = SqlDbType.VarChar;
                                                param13.Direction = ParameterDirection.Input;
                                                param13.Value = n.ResouceId;
                                                Cmd.Parameters.Add(param13);

                                                SqlParameter param14 = new SqlParameter();
                                                param14 = Cmd.CreateParameter();
                                                param14.ParameterName = "@IdmSubsucriptionId";
                                                param14.SqlDbType = SqlDbType.VarChar;
                                                param14.Direction = ParameterDirection.Input;
                                                param14.Value = n.IdmSubsucriptionId;
                                                Cmd.Parameters.Add(param14);

                                                SqlParameter param15 = new SqlParameter();
                                                param15 = Cmd.CreateParameter();
                                                param15.ParameterName = "@Tag";
                                                param15.SqlDbType = SqlDbType.VarChar;
                                                param15.Direction = ParameterDirection.Input;
                                                param15.Value = n.Tag;
                                                Cmd.Parameters.Add(param15);

                                                SqlParameter param16 = new SqlParameter();
                                                param16 = Cmd.CreateParameter();
                                                param16.ParameterName = "@UnitName";
                                                param16.SqlDbType = SqlDbType.VarChar;
                                                param16.Direction = ParameterDirection.Input;
                                                param16.Value = n.UnitName;
                                                Cmd.Parameters.Add(param16);

                                                SqlParameter param17 = new SqlParameter();
                                                param17 = Cmd.CreateParameter();
                                                param17.ParameterName = "@InstanceDataResourceUri";
                                                param17.SqlDbType = SqlDbType.VarChar;
                                                param17.Direction = ParameterDirection.Input;
                                                param17.Value = n.InstanceDataResourceUri;
                                                Cmd.Parameters.Add(param17);

                                                SqlParameter param18 = new SqlParameter();
                                                param18 = Cmd.CreateParameter();
                                                param18.ParameterName = "@InstanceDataLocation";
                                                param18.SqlDbType = SqlDbType.VarChar;
                                                param18.Direction = ParameterDirection.Input;
                                                param18.Value = n.InstanceDataLocation;
                                                Cmd.Parameters.Add(param18);

                                                SqlParameter param19 = new SqlParameter();
                                                param19 = Cmd.CreateParameter();
                                                param19.ParameterName = "@InstanceDataPartNumber";
                                                param19.SqlDbType = SqlDbType.VarChar;
                                                param19.Direction = ParameterDirection.Input;
                                                param19.Value = n.InstanceDataPartNumber;
                                                Cmd.Parameters.Add(param19);

                                                SqlParameter param20 = new SqlParameter();
                                                param20 = Cmd.CreateParameter();
                                                param20.ParameterName = "@InStanceDataOrderNumber";
                                                param20.SqlDbType = SqlDbType.VarChar;
                                                param20.Direction = ParameterDirection.Input;
                                                param20.Value = n.InStanceDataOrderNumber;
                                                Cmd.Parameters.Add(param20);

                                                SqlParameter param21 = new SqlParameter();
                                                param21 = Cmd.CreateParameter();
                                                param21.ParameterName = "@Domain";
                                                param21.SqlDbType = SqlDbType.VarChar;
                                                param21.Direction = ParameterDirection.Input;
                                                param21.Value = n.Domain;
                                                Cmd.Parameters.Add(param21);

                                                SqlParameter param22 = new SqlParameter();
                                                param22 = Cmd.CreateParameter();
                                                param22.ParameterName = "@SubscriptionId";
                                                param22.SqlDbType = SqlDbType.VarChar;
                                                param22.Direction = ParameterDirection.Input;
                                                param22.Value = n.SubscriptionId;
                                                Cmd.Parameters.Add(param22);

                                                SqlParameter param23 = new SqlParameter();
                                                param23 = Cmd.CreateParameter();
                                                param23.ParameterName = "@PurchaseValue";
                                                param23.SqlDbType = SqlDbType.Decimal;
                                                param23.Direction = ParameterDirection.Input;
                                                param23.Value = n.PurchaseValue;
                                                Cmd.Parameters.Add(param23);

                                                SqlParameter param24 = new SqlParameter();
                                                param24 = Cmd.CreateParameter();
                                                param24.ParameterName = "@UsafeFee";
                                                param24.SqlDbType = SqlDbType.Decimal;
                                                param24.Direction = ParameterDirection.Input;
                                                param24.Value = n.UsafeFee;
                                                Cmd.Parameters.Add(param24);

                                                SqlParameter param25 = new SqlParameter();
                                                param25 = Cmd.CreateParameter();
                                                param25.ParameterName = "@PurchasePrice";
                                                param25.SqlDbType = SqlDbType.Decimal;
                                                param25.Direction = ParameterDirection.Input;
                                                param25.Value = n.PurchasePrice;
                                                Cmd.Parameters.Add(param25);
                                                #endregion

                                                Cmd.CommandText = query;
                                                adapter.SelectCommand = Cmd;

                                                DataTable dt = new DataTable();

                                                Count = adapter.Fill(dt);

                                                loop_Count = 1;
                                            }

                                            if (Count > 0)
                                            {
                                                MessageBox.Show("既にデータが存在している為、処理できません。");
                                                return Count;
                                            }
                                            else
                                            {
                                                // 追加処理を実行します。
                                                customer_Insert_Sql(n.SpCompanyName, n.SpDomain, n.SpSubDomain, n.PurchaseCompanyCode, n.PurchaseCompanyName, n.DateofAcquisition, n.StartUsagePeriod, n.FinishUsagePeriod, n.ElementName, n.CategoyName, n.SubCategoryName, n.Area, n.ResouceId, n.IdmSubsucriptionId, n.Tag, n.UnitName, n.InstanceDataResourceUri, n.InstanceDataLocation, n.InstanceDataPartNumber, n.InStanceDataOrderNumber, n.Domain, n.SubscriptionId, n.PurchaseValue, n.UsafeFee, n.PurchasePrice);
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
                                MessageBox.Show("処理が終了しました。");
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

           int customer_Insert_Sql(string SpCompanyName, string SpDomain, string SpSubDomain, string PurchaseCompanyCode, string PurchaseCompanyName, string DateofAcquisition, string StartUsagePeriod, string FinishUsagePeriod, string ElementName, string CategoyName, string SubCategoryName, string Area, string ResouceId, string IdmSubsucriptionId, string Tag, string UnitName, string InstanceDataResourceUri, string InstanceDataLocation, string InstanceDataPartNumber, string InStanceDataOrderNumber, string Domain, string SubscriptionId, decimal PurchaseValue, decimal UsafeFee, decimal PurchasePrice)
            {
                int Count = 0;
                try
                {
                    using (var connection = SqlConection.GetConnection("customer"))
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
                                    query = "insert into Azure_details (SpCompanyName,SpDomain,SpSubDomain,PurchaseCompanyCode,PurchaseCompanyName,DateOfAcquisition,StartUsagePeriod,FinishUsagePeriod,ElementName,CategoyName,SubCategoryName,Area,ResouceId,IdmSubsucriptionId,Tag,UnitName,InstanceDataResourceUri,InstanceDataLocation,InstanceDataPartNumber,InStanceDataOrderNumber,Domain,SubscriptionId,PurchaseValue,UsafeFee,PurchasePrice) values (@SpCompanyName,@SpDomain,@SpSubDomain,@PurchaseCompanyCode,@PurchaseCompanyName,@DateofAcquisition,@StartUsagePeriod,@FinishUsagePeriod,@ElementName,@CategoyName,@SubCategoryName,@Area,@ResouceId,@IdmSubsucriptionId,@Tag,@UnitName,@InstanceDataResourceUri,@InstanceDataLocation,@InstanceDataPartNumber,@InStanceDataOrderNumber,@Domain,@SubscriptionId,@PurchaseValue,@UsafeFee,@PurchasePrice)";
                                    #region パラメータ
                                    SqlParameter param1 = new SqlParameter();
                                    param1 = Cmd.CreateParameter();
                                    param1.ParameterName = "@SpCompanyName";
                                    param1.SqlDbType = SqlDbType.VarChar;
                                    param1.Direction = ParameterDirection.Input;
                                    param1.Value = SpCompanyName;
                                    Cmd.Parameters.Add(param1);

                                    SqlParameter param2 = new SqlParameter();
                                    param2 = Cmd.CreateParameter();
                                    param2.ParameterName = "@SpDomain";
                                    param2.SqlDbType = SqlDbType.VarChar;
                                    param2.Direction = ParameterDirection.Input;
                                    param2.Value = SpDomain;
                                    Cmd.Parameters.Add(param2);

                                    SqlParameter param3 = new SqlParameter();
                                    param3 = Cmd.CreateParameter();
                                    param3.ParameterName = "@SpSubDomain";
                                    param3.SqlDbType = SqlDbType.VarChar;
                                    param3.Direction = ParameterDirection.Input;
                                    param3.Value = SpSubDomain;
                                    Cmd.Parameters.Add(param3);

                                    SqlParameter param4 = new SqlParameter();
                                    param4 = Cmd.CreateParameter();
                                    param4.ParameterName = "@PurchaseCompanyCode";
                                    param4.SqlDbType = SqlDbType.VarChar;
                                    param4.Direction = ParameterDirection.Input;
                                    param4.Value = PurchaseCompanyCode;
                                    Cmd.Parameters.Add(param4);

                                    SqlParameter param5 = new SqlParameter();
                                    param5 = Cmd.CreateParameter();
                                    param5.ParameterName = "@PurchaseCompanyName";
                                    param5.SqlDbType = SqlDbType.VarChar;
                                    param5.Direction = ParameterDirection.Input;
                                    param5.Value = PurchaseCompanyName;
                                    Cmd.Parameters.Add(param5);

                                    SqlParameter param6 = new SqlParameter();
                                    param6 = Cmd.CreateParameter();
                                    param6.ParameterName = "@DateofAcquisition";
                                    param6.SqlDbType = SqlDbType.VarChar;
                                    param6.Direction = ParameterDirection.Input;
                                    param6.Value = DateofAcquisition;
                                    Cmd.Parameters.Add(param6);

                                    SqlParameter param7 = new SqlParameter();
                                    param7 = Cmd.CreateParameter();
                                    param7.ParameterName = "@StartUsagePeriod";
                                    param7.SqlDbType = SqlDbType.DateTime;
                                    param7.Direction = ParameterDirection.Input;
                                    param7.Value = StartUsagePeriod;
                                    Cmd.Parameters.Add(param7);

                                    SqlParameter param8 = new SqlParameter();
                                    param8 = Cmd.CreateParameter();
                                    param8.ParameterName = "@FinishUsagePeriod";
                                    param8.SqlDbType = SqlDbType.DateTime;
                                    param8.Direction = ParameterDirection.Input;
                                    param8.Value = FinishUsagePeriod;
                                    Cmd.Parameters.Add(param8);

                                    SqlParameter param9 = new SqlParameter();
                                    param9 = Cmd.CreateParameter();
                                    param9.ParameterName = "@ElementName";
                                    param9.SqlDbType = SqlDbType.VarChar;
                                    param9.Direction = ParameterDirection.Input;
                                    param9.Value = ElementName;
                                    Cmd.Parameters.Add(param9);

                                    SqlParameter param10 = new SqlParameter();
                                    param10 = Cmd.CreateParameter();
                                    param10.ParameterName = "@CategoyName";
                                    param10.SqlDbType = SqlDbType.VarChar;
                                    param10.Direction = ParameterDirection.Input;
                                    param10.Value = CategoyName;
                                    Cmd.Parameters.Add(param10);

                                    SqlParameter param11 = new SqlParameter();
                                    param11 = Cmd.CreateParameter();
                                    param11.ParameterName = "@SubCategoryName";
                                    param11.SqlDbType = SqlDbType.VarChar;
                                    param11.Direction = ParameterDirection.Input;
                                    param11.Value = SubCategoryName;
                                    Cmd.Parameters.Add(param11);

                                    SqlParameter param12 = new SqlParameter();
                                    param12 = Cmd.CreateParameter();
                                    param12.ParameterName = "@Area";
                                    param12.SqlDbType = SqlDbType.VarChar;
                                    param12.Direction = ParameterDirection.Input;
                                    param12.Value = Area;
                                    Cmd.Parameters.Add(param12);

                                    SqlParameter param13 = new SqlParameter();
                                    param13 = Cmd.CreateParameter();
                                    param13.ParameterName = "@ResouceId";
                                    param13.SqlDbType = SqlDbType.VarChar;
                                    param13.Direction = ParameterDirection.Input;
                                    param13.Value = ResouceId;
                                    Cmd.Parameters.Add(param13);

                                    SqlParameter param14 = new SqlParameter();
                                    param14 = Cmd.CreateParameter();
                                    param14.ParameterName = "@IdmSubsucriptionId";
                                    param14.SqlDbType = SqlDbType.VarChar;
                                    param14.Direction = ParameterDirection.Input;
                                    param14.Value = IdmSubsucriptionId;
                                    Cmd.Parameters.Add(param14);

                                    SqlParameter param15 = new SqlParameter();
                                    param15 = Cmd.CreateParameter();
                                    param15.ParameterName = "@Tag";
                                    param15.SqlDbType = SqlDbType.VarChar;
                                    param15.Direction = ParameterDirection.Input;
                                    param15.Value = Tag;
                                    Cmd.Parameters.Add(param15);

                                    SqlParameter param16 = new SqlParameter();
                                    param16 = Cmd.CreateParameter();
                                    param16.ParameterName = "@UnitName ";
                                    param16.SqlDbType = SqlDbType.VarChar;
                                    param16.Direction = ParameterDirection.Input;
                                    param16.Value = UnitName;
                                    Cmd.Parameters.Add(param16);

                                    SqlParameter param17 = new SqlParameter();
                                    param17 = Cmd.CreateParameter();
                                    param17.ParameterName = "@InstanceDataResourceUri";
                                    param17.SqlDbType = SqlDbType.VarChar;
                                    param17.Direction = ParameterDirection.Input;
                                    param17.Value = InstanceDataResourceUri;
                                    Cmd.Parameters.Add(param17);

                                    SqlParameter param18 = new SqlParameter();
                                    param18 = Cmd.CreateParameter();
                                    param18.ParameterName = "@InstanceDataLocation";
                                    param18.SqlDbType = SqlDbType.VarChar;
                                    param18.Direction = ParameterDirection.Input;
                                    param18.Value = InstanceDataLocation;
                                    Cmd.Parameters.Add(param18);

                                    SqlParameter param19 = new SqlParameter();
                                    param19 = Cmd.CreateParameter();
                                    param19.ParameterName = "@InstanceDataPartNumber";
                                    param19.SqlDbType = SqlDbType.VarChar;
                                    param19.Direction = ParameterDirection.Input;
                                    param19.Value = InstanceDataPartNumber;
                                    Cmd.Parameters.Add(param19);

                                    SqlParameter param20 = new SqlParameter();
                                    param20 = Cmd.CreateParameter();
                                    param20.ParameterName = "@InStanceDataOrderNumber";
                                    param20.SqlDbType = SqlDbType.VarChar;
                                    param20.Direction = ParameterDirection.Input;
                                    param20.Value = InStanceDataOrderNumber;
                                    Cmd.Parameters.Add(param20);

                                    SqlParameter param21 = new SqlParameter();
                                    param21 = Cmd.CreateParameter();
                                    param21.ParameterName = "@Domain";
                                    param21.SqlDbType = SqlDbType.VarChar;
                                    param21.Direction = ParameterDirection.Input;
                                    param21.Value = Domain;
                                    Cmd.Parameters.Add(param21);

                                    SqlParameter param22 = new SqlParameter();
                                    param22 = Cmd.CreateParameter();
                                    param22.ParameterName = "@SubscriptionId";
                                    param22.SqlDbType = SqlDbType.VarChar;
                                    param22.Direction = ParameterDirection.Input;
                                    param22.Value = SubscriptionId;
                                    Cmd.Parameters.Add(param22);

                                    SqlParameter param23 = new SqlParameter();
                                    param23 = Cmd.CreateParameter();
                                    param23.ParameterName = "@PurchaseValue";
                                    param23.SqlDbType = SqlDbType.Decimal;
                                    param23.Direction = ParameterDirection.Input;
                                    param23.Value = PurchaseValue;
                                    Cmd.Parameters.Add(param23);

                                    SqlParameter param24 = new SqlParameter();
                                    param24 = Cmd.CreateParameter();
                                    param24.ParameterName = "@UsafeFee";
                                    param24.SqlDbType = SqlDbType.Decimal;
                                    param24.Direction = ParameterDirection.Input;
                                    param24.Value = UsafeFee;
                                    Cmd.Parameters.Add(param24);

                                    SqlParameter param25 = new SqlParameter();
                                    param25 = Cmd.CreateParameter();
                                    param25.ParameterName = "@PurchasePrice";
                                    param25.SqlDbType = SqlDbType.Decimal;
                                    param25.Direction = ParameterDirection.Input;
                                    param25.Value = PurchasePrice;
                                    Cmd.Parameters.Add(param25);
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
}
