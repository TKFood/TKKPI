using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using FastReport;
using FastReport.Data;
using TKITDLL;
using System.Data.SQLite;

namespace TKKPI
{
    public partial class frmREPORTVISITORS : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;

        SqlTransaction tran;
        int result;

        public frmREPORTVISITORS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();
            StringBuilder SQL3 = new StringBuilder();
            StringBuilder SQL4 = new StringBuilder();
            StringBuilder SQL5 = new StringBuilder();
            StringBuilder SQL6 = new StringBuilder();

            SQL1 = SETSQL(); 
            SQL2 = SETSQL2();
            SQL3 = SETSQL3();
            SQL4 = SETSQL4();
            SQL5 = SETSQL5();
            SQL6 = SETSQL6();

            Report report1 = new Report();
            report1.Load(@"REPORT\營銷來客報表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

             
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL2.ToString();
            TableDataSource table2= report1.GetDataSource("Table2") as TableDataSource;
            table2.SelectCommand = SQL3.ToString();
            TableDataSource table3 = report1.GetDataSource("Table3") as TableDataSource;
            table3.SelectCommand = SQL4.ToString();
            TableDataSource table4 = report1.GetDataSource("Table4") as TableDataSource;
            table4.SelectCommand = SQL5.ToString();
            TableDataSource table5 = report1.GetDataSource("Table5") as TableDataSource;
            table5.SelectCommand = SQL6.ToString();


            report1.Preview = previewControl1;
            report1.Show();
        }
        /// <summary>
        /// 查詢各門市的各月
        /// </summary>
        /// <returns></returns>
        public StringBuilder SETSQL()
        {
            
            StringBuilder SB = new StringBuilder();
               
            SB.AppendFormat(@" 
                            SELECT  ME001,ME002,YEARS,MONTHS,SUM(TT008) SUMTT008,SUM(TT018)/SUM(TT008) AS 'AVGTT018',SUM(TT018) SUMTT018
                            FROM 
                            (
                            SELECT ME001,ME002,TT001,SUBSTRING(TT001,1,4) AS 'YEARS',SUBSTRING(TT001,5,2)  AS 'MONTHS',TT018,TT008
                            FROM [TK].dbo.POSTT,[TK].dbo.CMSME
                            WHERE TT002=ME001
                            AND TT001 LIKE '{0}%'
                            ) AS TEMP
                            WHERE ME001 LIKE '106%'
                            GROUP BY ME001,ME002,YEARS,MONTHS
                            ORDER BY ME001,ME002,YEARS,MONTHS 
                            ", dateTimePicker1.Value.ToString("yyyy"));

            return SB;

        }
        /// <summary>
        /// 查詢各門市的各週
        /// </summary>
        /// <returns></returns>
        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 

                            SELECT TT002,STORESNAME,YEARS,WEEKS,SUM(NUMS) NUMS,SUM(SUMTT011) SUMTT011,SUM(SUMTT008) SUMTT008
                            ,(SUM(SUMTT008)/SUM(NUMS)) AS 'PCTS',(SUM(SUMTT011)/SUM(SUMTT008)) AS 'AVGTT011'
                            FROM (
                            SELECT View_t_visitors.TT002,STORESNAME,YEARS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fin_data+Fout_data)/2 AS NUMS
                            ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                            ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704') 
                            AND YEARS='{0}'
                            GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,WEEKS,Fdate1,DAYOFWEEK
   
                            UNION ALL
                            SELECT View_t_visitors.TT002,STORESNAME,YEARS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fout_data) AS NUMS
                            ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                            ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  TT002 IN ('106701') 
                            AND YEARS='{0}'
                            GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,WEEKS,Fdate1,DAYOFWEEK

                            ) AS TEMP
                            GROUP BY TT002,STORESNAME,YEARS,WEEKS
                            ORDER BY TT002,STORESNAME,YEARS,CONVERT(INT,WEEKS)
                            ", dateTimePicker1.Value.ToString("yyyy"));

            return SB;

        }
        /// <summary>
        /// 查詢各門市的各月的交易時間 
        /// </summary>
        /// <returns></returns>

        public StringBuilder SETSQL3()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"  
                           SELECT TT002,STORESNAME,YEARS,MONTHS,HOURS,DAYSS,SUM(NUMS) NUMS,SUM(SUMTA026) SUMTA026,SUM(COUNTSTA026) COUNTSTA026
                            ,(CASE WHEN SUM(NUMS)>0 AND SUM(COUNTSTA026)>0 THEN  ROUND(CONVERT(decimal,SUM(COUNTSTA026),2)/CONVERT(decimal,SUM(NUMS),2),4) ELSE 0 END) AS 'PCTS'
                            ,(CASE WHEN SUM(COUNTSTA026)>0 AND SUM(SUMTA026)>0 THEN  SUM(SUMTA026)/SUM(COUNTSTA026) ELSE 0 END )AS 'AVGTA026'
                            FROM (
                            SELECT 

                            TT002,STORESNAME,YEARS,MONTHS,[Fdate1],HOURS,SUM(Fin_data+Fout_data)/2 AS NUMS, day(dateadd(ms,-3,DATEADD(m, DATEDIFF(m,0,CONVERT(datetime,CONVERT(NVARCHAR(4),YEARS)+'/'+CONVERT(NVARCHAR(4),MONTHS)+'/1'))+1,0))) AS DAYSS
                            ,(SELECT ISNULL(SUM(TA026),0) FROM [TK].[dbo].[POSTA] WITH(NOLOCK)  WHERE [POSTA].TA002=[View_t_visitors].TT002 AND [POSTA].TA004=[View_t_visitors].[Fdate1] AND [POSTA].HHS= Right('00' + Cast([View_t_visitors].HOURS as varchar),2)) AS 'SUMTA026'
                            ,(SELECT ISNULL(COUNT(TA026),0) FROM [TK].[dbo].[POSTA] WITH(NOLOCK)  WHERE [POSTA].TA002=[View_t_visitors].TT002 AND [POSTA].TA004=[View_t_visitors].[Fdate1] AND [POSTA].HHS=Right('00' + Cast([View_t_visitors].HOURS as varchar),2)) AS 'COUNTSTA026'
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704') 
                            AND YEARS='{0}'
                            GROUP BY  TT002,STORESNAME,YEARS,MONTHS,[Fdate1],HOURS


                            UNION ALL
                            SELECT TT002,STORESNAME,YEARS,MONTHS,[Fdate1],HOURS,SUM(Fout_data) AS NUMS, day(dateadd(ms,-3,DATEADD(m, DATEDIFF(m,0,CONVERT(datetime,CONVERT(NVARCHAR(4),YEARS)+'/'+CONVERT(NVARCHAR(4),MONTHS)+'/1'))+1,0))) AS DAYSS
                            ,(SELECT ISNULL(SUM(TA026),0) FROM [TK].[dbo].[POSTA] WITH(NOLOCK)  WHERE [POSTA].TA002=[View_t_visitors].TT002 AND [POSTA].TA004=[View_t_visitors].[Fdate1] AND [POSTA].HHS=Right('00' + Cast([View_t_visitors].HOURS as varchar),2)) AS 'SUMTA026'
                            ,(SELECT ISNULL(COUNT(TA026),0) FROM [TK].[dbo].[POSTA] WITH(NOLOCK)  WHERE [POSTA].TA002=[View_t_visitors].TT002 AND [POSTA].TA004=[View_t_visitors].[Fdate1] AND [POSTA].HHS=Right('00' + Cast([View_t_visitors].HOURS as varchar),2)) AS 'COUNTSTA026'
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  TT002 IN ('106701') 
                            AND YEARS='{0}'
                            GROUP BY  TT002,STORESNAME,YEARS,MONTHS,[Fdate1],HOURS
                            ) AS TEMP
                            GROUP BY TT002,STORESNAME,YEARS,MONTHS,HOURS,DAYSS
                            ORDER BY TT002,STORESNAME,YEARS,MONTHS,CONVERT(INT,HOURS)
                            ", dateTimePicker1.Value.ToString("yyyy"));

            return SB;

        }
        /// <summary>
        /// 找最近的前10筆
        /// </summary>
        /// <returns></returns>
        public StringBuilder SETSQL4()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            
                            
                            SELECT TOP 10 
                            [TT002]
                            ,[Fdevice_sn]
                            ,[STORESNAME]
                            ,[Fdate1]
                            ,[Fdate2]
                            ,[Fin_data]
                            ,[Fout_data]
                            ,[id]
                            ,[Fdate]
                            ,[YEARS]
                            ,[MONTHS]
                            ,[DAYS]
                            ,[DAYOFWEEK]
                            ,[WEEKS]
                            ,[HOURS]
                            FROM [TKMK].[dbo].[View_t_visitors]
                            ORDER BY [Fdate] DESC  

                            ");

            return SB;

        }
        /// <summary>
        /// 查詢各門市的各月的交易星期
        /// </summary>
        /// <returns></returns>
        public StringBuilder SETSQL5()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"  
                            SELECT TT002,STORESNAME,YEARS,MONTHS
                            ,COUNT(WEEKS) WEEKSCOUNTS,DAYOFWEEK,SUM(NUMS) NUMS,SUM(SUMTT011) SUMTT011,SUM(SUMTT008) SUMTT008
                            ,CASE WHEN SUM(NUMS)>0 AND COUNT(WEEKS) >0 THEN  SUM(NUMS)/COUNT(WEEKS)  ELSE 0 END AS 'NUMSAVGS'
                            ,CASE WHEN SUM(SUMTT008)>0 AND SUM(NUMS) >0 THEN SUM(SUMTT008)/SUM(NUMS) ELSE 0 END  AS 'PCTS'
                            ,CASE WHEN SUM(SUMTT011)>0 AND SUM(SUMTT008) >0 THEN  SUM(SUMTT011)/SUM(SUMTT008) ELSE 0 END  AS 'AVGTT011',WEEKDAY

                            FROM (
                            SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fin_data+Fout_data)/2 AS NUMS,DATEPART(WEEKDAY,Fdate1) AS 'WEEKDAY'
                            ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                            ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  View_t_visitors.TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704') 
                            AND YEARS='{0}'
                            GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,DATEPART(WEEKDAY,Fdate1)
 
                            UNION ALL
                            SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fout_data) AS NUMS,DATEPART(WEEKDAY,Fdate1) AS 'WEEKDAY'
                            ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                            ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  View_t_visitors.TT002 IN ('106701') 
                            AND YEARS='{0}'
                            GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,DATEPART(WEEKDAY,Fdate1)

                            ) AS TEMP
                            GROUP BY TT002,STORESNAME,YEARS,MONTHS,WEEKDAY,DAYOFWEEK
                            ORDER BY TT002,STORESNAME,YEARS,MONTHS,WEEKDAY,DAYOFWEEK
 
                            ", dateTimePicker1.Value.ToString("yyyy")); 

            return SB;

        }
        /// <summary>
        /// 查詢各門市的各月
        /// </summary>
        /// <returns></returns>
        public StringBuilder SETSQL6()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                           SELECT TT002,STORESNAME,YEARS,MONTHS,SUM(SUMNUMS) SUMNUMS,SUM(SUMTT011) SUMTT011,SUM(SUMTT008) SUMTT008
                            ,SUM(SUMTT008)/SUM(SUMNUMS) AS PCTS,SUM(SUMTT011)/SUM(SUMTT008) AS AVGTT011
                            
                            ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WHERE POSTT.TT002=TEMP.TT002 AND POSTT.TT001 LIKE TEMP.YEARS+Right('00' + Cast(TEMP.MONTHS as varchar),2)+'%' ) AS 'REALSUMTT008'
                            ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WHERE POSTT.TT002=TEMP.TT002 AND POSTT.TT001 LIKE TEMP.YEARS+Right('00' + Cast(TEMP.MONTHS as varchar),2)+'%' ) AS 'REALSUMTT018'
                            ,((SELECT SUM(TT018) FROM [TK].dbo.POSTT WHERE POSTT.TT002=TEMP.TT002 AND POSTT.TT001 LIKE TEMP.YEARS+Right('00' + Cast(TEMP.MONTHS as varchar),2)+'%' )/(SELECT SUM(TT008) FROM [TK].dbo.POSTT WHERE POSTT.TT002=TEMP.TT002 AND POSTT.TT001 LIKE TEMP.YEARS+Right('00' + Cast(TEMP.MONTHS as varchar),2)+'%' )) AS 'REALAVGTT018'
                            FROM (
                            SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fin_data+Fout_data)/2 AS SUMNUMS
                            ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                            ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704') 
                            AND YEARS='{0}'
                            GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK
   
                            UNION ALL
                            SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fout_data) AS SUMNUMS
                            ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                            ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  TT002 IN ('106701') 
                            AND YEARS='{0}'
                            GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK
                            ) AS TEMP
                            GROUP BY TT002,STORESNAME,YEARS,MONTHS
                            ORDER BY TT002,STORESNAME,YEARS,MONTHS





                            ", dateTimePicker1.Value.ToString("yyyy"));

            return SB;

        }


        public void ADDTKMKt_visitors()
        {
            SQLiteConnection SQLiteConnection = new SQLiteConnection();
            string MAXID = null;
           

            try
            {
                MAXID = FINDTKMKt_visitorsMAXID();

                if(!string.IsNullOrEmpty(MAXID))
                {
                    //SQLite的檔案要先copy到 F:\kldatabase.db
                    string path = @"E:\kldatabase.db";
                    SQLiteConnection = new SQLiteConnection("data source=" + path);
                    SQLiteConnection.Open();

                    SQLiteCommand cmd = SQLiteConnection.CreateCommand();

                    sbSql.Clear();
                    sbSql.AppendFormat(@"  
                                        SELECT *
                                        FROM t_visitors
                                        WHERE ID>'{0}'
                                     ", MAXID);

                    cmd.CommandText = sbSql.ToString();

                    // 用DataAdapter和DataTable類，記得要 using System.Data
                    SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd);
                    DataTable table = new DataTable();
                    adapter.Fill(table);

                    if(table.Rows.Count>0)
                    {
                        ADDTOTKMKt_visitors(table);
                        UPDATEt_visitors();
                    }

                    else
                    {
                        MessageBox.Show("沒有新資料，請更新kldatabasepri 到E:");
                    }

                    SQLiteConnection.Close();


                }
                else
                {
                    MessageBox.Show("沒有新資料，請更新kldatabasepri 到E:");
                }
               
            }
            catch
            {
                MessageBox.Show("有錯誤");
            }
            finally
            {
                
            }
           
        }

        public string FINDTKMKt_visitorsMAXID()
        {
            string MAXID = null;

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT TOP 1 MAX([id])  id  FROM [TKMK].[dbo].[t_visitors]
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    MAXID = ds1.Tables["TEMPds1"].Rows[0]["id"].ToString();                    
                }
                else
                {

                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }

            return MAXID;
        }

        public void ADDTOTKMKt_visitors(DataTable dtt_visitors)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconnTKMK"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            using (SqlConnection connection = sqlConn)
            {
                connection.Open();
                SqlTransaction sqlTrans = connection.BeginTransaction();
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(connection, SqlBulkCopyOptions.KeepIdentity, sqlTrans))
                {
                    DataTable dt = dtt_visitors;
                    bulkCopy.DestinationTableName = "t_visitors";

                    //對應資料行
                    //bulkCopy.ColumnMappings.Add("DataTable的欄位A", "資料庫裡的資料表的的欄位A");
                    bulkCopy.ColumnMappings.Add("Fuid", "Fuid");
                    bulkCopy.ColumnMappings.Add("Fvisit_md5", "Fvisit_md5");
                    bulkCopy.ColumnMappings.Add("Fdevice_sn", "Fdevice_sn");
                    bulkCopy.ColumnMappings.Add("Fdate", "Fdate");
                    bulkCopy.ColumnMappings.Add("Fin_data", "Fin_data");
                    bulkCopy.ColumnMappings.Add("Fout_data", "Fout_data");
                    bulkCopy.ColumnMappings.Add("Fcreate_time", "Fcreate_time");
                    bulkCopy.ColumnMappings.Add("Fdata_version", "Fdata_version");
                    bulkCopy.ColumnMappings.Add("Fbatvoltage", "Fbatvoltage");
                    bulkCopy.ColumnMappings.Add("Fbatpercent", "Fbatpercent");
                    bulkCopy.ColumnMappings.Add("Flosefocus", "Flosefocus");
                    bulkCopy.ColumnMappings.Add("Fcharge", "Fcharge");
                    bulkCopy.ColumnMappings.Add("Ftemperature", "Ftemperature");
                    bulkCopy.ColumnMappings.Add("id", "id");

                    bulkCopy.BatchSize = 1000;
                    bulkCopy.BulkCopyTimeout = 60;

                    try
                    {
                        bulkCopy.WriteToServer(dt);
                        sqlTrans.Commit();

                        MessageBox.Show("完成");
                    }

                    catch (Exception)
                    {
                        sqlTrans.Rollback();                       
                    }

                   


                }

            }
        }


        public void SETFASTREPORT2(string SDATES,string EDATES)
        {
            StringBuilder SQL1 = new StringBuilder();


            SQL1.AppendFormat(@"
                                SELECT  TA001 AS '日期',TA002 AS '門市代',MA002 AS '門市',NUMS AS '成交筆數',MMS AS '交易金額',CLINETS AS '來客數',CARS AS '團車數'
                                FROM 
                                (
                                SELECT TA001,TA002,MA002,COUNT(TA001) AS NUMS,SUM(TA026) AS MMS
                                ,(SELECT SUM(Fout_data) FROM [TKMK].[dbo].[View_t_visitors] WHERE TT002=TA002 AND Fdate1=TA001) AS 'CLINETS'
                                ,(SELECT SUM([CARNUM]) FROM [TKMK].[dbo].[GROUPSALES] WHERE CONVERT(NVARCHAR,[CREATEDATES],112)= TA001) AS 'CARS'
                                FROM [TK].dbo.POSTA,[TK].dbo.WSCMA
                                WHERE 1=1
                                AND TA002=MA001
                                AND TA002 IN ('106701') 
                                AND TA001>='{0}' AND TA001<='{1}'
                                GROUP BY TA001,TA002,MA002
                                UNION ALL
                                SELECT TA001,TA002,MA002,COUNT(TA001) AS NUMS,SUM(TA026) AS MMS
                                ,(SELECT SUM(Fin_data+Fout_data)/2 FROM [TKMK].[dbo].[View_t_visitors] WHERE TT002=TA002 AND Fdate1=TA001) AS 'CLINETS'
                                ,0 AS 'CARS'
                                FROM [TK].dbo.POSTA,[TK].dbo.WSCMA
                                WHERE 1=1
                                AND TA002=MA001
                                AND TA002 IN ('106501','106502','106503','106504') 
                                AND TA001>='{0}' AND TA001<='{1}'
                                GROUP BY TA001,TA002,MA002
                                ) 
                                AS TEMP
                                ORDER BY TA002,MA002,TA001
                                ", SDATES, EDATES);


            Report report1 = new Report();
            report1.Load(@"REPORT\每日來客報表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
   


            report1.Preview = previewControl2;
            report1.Show();
        }


        public void SETFASTREPORT3(string YY,string MM)
        {
            StringBuilder SQL1 = new StringBuilder();



            SQL1.AppendFormat(@"
                               
                                SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fin_data+Fout_data)/2 AS SUMNUMS
                                ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                                ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                                FROM [TKMK].[dbo].[View_t_visitors]
                                WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704') 
                                 AND YEARS='{0}'
                                AND MONTHS='{1}'
                                GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK

                                UNION ALL
                                SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fout_data) AS SUMNUMS
                                ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                                ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                                FROM [TKMK].[dbo].[View_t_visitors]
                                WHERE  TT002 IN ('106701') 
                                AND YEARS='{0}'
                                AND MONTHS='{1}'
              
                                GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK
                                ORDER BY View_t_visitors.TT002,Fdate1

                                ", YY,MM); 


            Report report1 = new Report();
            report1.Load(@"REPORT\來客數-每月.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();



            report1.Preview = previewControl3;
            report1.Show();
        }


        public void SETFASTREPORT4(string YY, string MM)
        { 
            StringBuilder SQL1 = new StringBuilder();

               

            SQL1.AppendFormat(@"
                                SELECT TT002 AS '門代'
                                ,STORESNAME AS '門店'
                                ,YEARS AS '年'
                                ,MONTHS AS '月'
                                ,WEEKS AS '週'
                                ,Fdate1 AS '日'
                                ,DAYOFWEEK AS '星期'
                                ,SUMNUMS AS '來客數'
                                ,SUMTT011 AS '銷售金額'
                                ,COUNTSTA001 AS '交易筆數'
                                ,SUMSTB019 AS '交易商品數'
                                ,(CASE WHEN SUMNUMS>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,2),((CONVERT(DECIMAL(16,4),COUNTSTA001)/CONVERT(DECIMAL(16,4),SUMNUMS))*100)) ELSE 0 END ) AS '提袋率'
                                ,(CASE WHEN SUMNUMS>0 AND SUMTT011>0 THEN CONVERT(DECIMAL(16,0),SUMTT011/SUMNUMS) ELSE 0 END ) AS '客單價'
                                ,(CASE WHEN SUMTT011>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,0),SUMTT011/COUNTSTA001) ELSE 0 END ) AS '每單單價'
                                ,(CASE WHEN SUMSTB019>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,2),SUMSTB019/COUNTSTA001) ELSE 0 END ) AS '每單平均商品數'

                                FROM 
                                (
                                SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fin_data+Fout_data)/2 AS SUMNUMS
                                ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                                ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT  WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                                ,(SELECT COUNT(TA001) FROM [TK].dbo.POSTA WITH(NOLOCK)  WHERE  TA002=View_t_visitors.TT002 AND TA004=View_t_visitors.Fdate1) AS 'COUNTSTA001'
                                ,(SELECT SUM(TB019) FROM [TK].dbo.POSTB  WITH(NOLOCK) WHERE  TB002=View_t_visitors.TT002 AND TB004=View_t_visitors.Fdate1 AND TB010 NOT LIKE '1%'  AND TB010 NOT LIKE '2%'  AND TB010 NOT LIKE '3%') AS 'SUMSTB019'
                                FROM [TKMK].[dbo].[View_t_visitors]
                                WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704') 
                                AND YEARS='{0}'
                                AND MONTHS='{1}'
                                GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK

                                UNION ALL
                                SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fout_data) AS SUMNUMS
                                ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                                ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WITH(NOLOCK)  WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                                ,(SELECT COUNT(TA001) FROM [TK].dbo.POSTA WITH(NOLOCK)  WHERE  TA002=View_t_visitors.TT002 AND TA004=View_t_visitors.Fdate1) AS 'COUNTSTA001'
                                ,(SELECT SUM(TB019) FROM [TK].dbo.POSTB  WITH(NOLOCK) WHERE  TB002=View_t_visitors.TT002 AND TB004=View_t_visitors.Fdate1 AND TB010 NOT LIKE '1%'  AND TB010 NOT LIKE '2%'  AND TB010 NOT LIKE '3%') AS 'SUMSTB019'
                                FROM [TKMK].[dbo].[View_t_visitors]
                                WHERE  TT002 IN ('106701') 
                                AND YEARS='{0}'
                                AND MONTHS='{1}'
              
                                GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK
                                ) AS TEMP
                                ORDER BY TT002,Fdate1

                                

                                ", YY, MM);


            Report report1 = new Report();
            report1.Load(@"REPORT\營銷-每日明細.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();



            report1.Preview = previewControl4;
            report1.Show();
        }


        public void UPDATEt_visitors()
        {

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"
                                     UPDATE [TKMK].[dbo].[t_visitors]
                                    SET [TT002]= [t_STORESNAME].[TT002],[STORESNAME]=[t_STORESNAME].[STORESNAME]
                                    FROM [TKMK].[dbo].[t_STORESNAME]
                                    WHERE [t_STORESNAME].[Fdevice_sn]=[t_visitors].[Fdevice_sn]
                                    AND ISNULL([t_visitors].[TT002],'')=''

                                    ");



                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消


                }
                else
                {
                    tran.Commit();      //執行交易                    

                }

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }


        }

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            ADDTKMKt_visitors();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3(dateTimePicker4.Value.Year.ToString(), dateTimePicker4.Value.Month.ToString());
        }
        private void button5_Click(object sender, EventArgs e)
        {
            SETFASTREPORT4(dateTimePicker5.Value.Year.ToString(), dateTimePicker5.Value.Month.ToString());
        }

        #endregion


    }
}
