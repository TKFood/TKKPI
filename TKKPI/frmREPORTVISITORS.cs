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
using System.Web;

namespace TKKPI
{
    public partial class frmREPORTVISITORS : Form
    {
        int TIMEOUT = 240;

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

            SETDATE();
        }

        #region FUNCTION
        public void SETDATE()
        {
            DateTime today = DateTime.Today;
            int diff = (7 + (today.DayOfWeek - DayOfWeek.Monday)) % 7;

            DateTime monday = today.AddDays(-1 * diff).AddDays(-7);
            DateTime sunday = monday.AddDays(6);

            dateTimePicker5.Value = monday;
            dateTimePicker8.Value = sunday;
        }
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
            report1.Dictionary.Connections[0].CommandTimeout = TIMEOUT;

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
                            WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
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
                            WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
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
                            ,CASE WHEN SUM(SUMTT011)>0 AND SUM(SUMTT008) >0 THEN  SUM(SUMTT011)/SUM(SUMTT008) ELSE 0 END  AS 'AVGTT011'
                            ,(CASE WHEN WEEKDAY=1 THEN 99 ELSE WEEKDAY END ) AS WEEKDAY

                            FROM (
                            SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fin_data+Fout_data)/2 AS NUMS,DATEPART(WEEKDAY,Fdate1) AS 'WEEKDAY'
                            ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                            ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  View_t_visitors.TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
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
                            WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
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
            //string MAXID = null;
            string MAX_Fdate = null;



            try
            {
                //MAXID = FINDTKMKt_visitorsMAXID();
                MAX_Fdate = FIND_TKMKt_visitorsMAX_Fdate();

                if (!string.IsNullOrEmpty(MAX_Fdate))
                {
                    //SQLite的檔案要先copy到 F:\kldatabase.db
                   // string path = @"data source=E:\kldatabase.db";
                    string path = @"data source=X:\kldatabase.db";
                    //string path = @"data source=\\192.168.1.101\Users\Administrator\AppData\Roaming\CounterServerData\kldatabase.db";

                    //string filePath = @"\\192.168.1.101\Users\Administrator\AppData\Roaming\CounterServerData\kldatabase.db";
                    //if (File.Exists(filePath))
                    //{
                    //    MessageBox.Show("存在！");
                    //}
                    //else
                    //{
                    //    MessageBox.Show("檔案不存在！");
                    //}

                    SQLiteConnection = new SQLiteConnection(path);
                    SQLiteConnection.Open();

                    SQLiteCommand cmd = SQLiteConnection.CreateCommand();

                    sbSql.Clear();
                    sbSql.AppendFormat(@"  
                                        SELECT *
                                        FROM t_visitors
                                        WHERE Fdate>='{0}'
                                     ", MAX_Fdate);

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

        public string FIND_TKMKt_visitorsMAX_Fdate()
        {
            string Fdate = null;

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
                                    SELECT 
                                    CONVERT(VARCHAR, MAX([Fdate]), 120) AS [Fdate]
                                    FROM  [TKMK].[dbo].[t_visitors]
                                    ");

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    Fdate = ds1.Tables["TEMPds1"].Rows[0]["Fdate"].ToString();
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

            return Fdate;
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
                                WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
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
                                ,CONVERT(INT,SUMTT018) AS '銷售未稅總金額'
                                ,COUNTSTA001 AS '結帳單量'
                                ,CONVERT(INT,SUMSTB019) AS '結帳交易商品數'
                                ,(CASE WHEN SUMNUMS>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,2),((CONVERT(DECIMAL(16,4),COUNTSTA001)/CONVERT(DECIMAL(16,4),SUMNUMS)))) ELSE 0 END ) AS '每日結帳單量/來客數(提袋率)'
                                ,(CASE WHEN SUMTT018>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,0),SUMTT018/COUNTSTA001) ELSE 0 END ) AS '平均每單單價(客單價)'
                                ,(CASE WHEN SUMSTB019>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,2),SUMSTB019/COUNTSTA001) ELSE 0 END ) AS '每單平均商品數'

                                FROM 
                                (
                                SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fin_data+Fout_data)/2 AS SUMNUMS
                                ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT018'
                                ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT  WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                                ,(SELECT COUNT(TA001) FROM [TK].dbo.POSTA WITH(NOLOCK)  WHERE  TA002=View_t_visitors.TT002 AND TA004=View_t_visitors.Fdate1) AS 'COUNTSTA001'
                                ,(SELECT SUM(TB019) FROM [TK].dbo.POSTB  WITH(NOLOCK) WHERE  TB002=View_t_visitors.TT002 AND TB004=View_t_visitors.Fdate1 AND TB010 NOT LIKE '1%'  AND TB010 NOT LIKE '2%'  AND TB010 NOT LIKE '3%') AS 'SUMSTB019'
                                FROM [TKMK].[dbo].[View_t_visitors]
                                WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
                                AND CONVERT(NVARCHAR,Fdate1,112)>='{0}'
                                AND CONVERT(NVARCHAR,Fdate1,112)<='{1}'
                                GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK

                                UNION ALL
                                SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fout_data) AS SUMNUMS
                                ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                                ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WITH(NOLOCK)  WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                                ,(SELECT COUNT(TA001) FROM [TK].dbo.POSTA WITH(NOLOCK)  WHERE  TA002=View_t_visitors.TT002 AND TA004=View_t_visitors.Fdate1) AS 'COUNTSTA001'
                                ,(SELECT SUM(TB019) FROM [TK].dbo.POSTB  WITH(NOLOCK) WHERE  TB002=View_t_visitors.TT002 AND TB004=View_t_visitors.Fdate1 AND TB010 NOT LIKE '1%'  AND TB010 NOT LIKE '2%'  AND TB010 NOT LIKE '3%') AS 'SUMSTB019'
                                FROM [TKMK].[dbo].[View_t_visitors]
                                WHERE  TT002 IN ('106701') 
                                AND CONVERT(NVARCHAR,Fdate1,112)>='{0}'
                                AND CONVERT(NVARCHAR,Fdate1,112)<='{1}'
              
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

        public void SETFASTREPORT7(string SDAYS, string EDAYS)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();
            StringBuilder SQL3 = new StringBuilder();

            SQL1.AppendFormat(@" 
                               
                                SELECT TT002 AS '門代'
                                ,STORESNAME AS '門店'
                                ,YEARS AS '年'
                                ,MONTHS AS '月'
                                ,WEEKS AS '週'
                                ,Fdate1 AS '日'
                                ,DAYOFWEEK AS '星期'
                                ,SUMNUMS AS '來客數'
                                ,CONVERT(INT,SUMTT018) AS '銷售未稅總金額'
                                ,COUNTSTA001 AS '結帳單量'
                                ,CONVERT(INT,SUMSTB019) AS '結帳交易商品數'
                                ,(CASE WHEN SUMNUMS>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,2),((CONVERT(DECIMAL(16,4),COUNTSTA001)/CONVERT(DECIMAL(16,4),SUMNUMS)))) ELSE 0 END ) AS '每日結帳單量/來客數(提袋率)'
                                ,(CASE WHEN SUMTT018>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,0),SUMTT018/COUNTSTA001) ELSE 0 END ) AS '平均每單單價(客單價)'
                                ,(CASE WHEN SUMSTB019>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,2),SUMSTB019/COUNTSTA001) ELSE 0 END ) AS '每單平均商品數'

                                FROM 
                                (
                                SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fin_data+Fout_data)/2 AS SUMNUMS
                                ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT018'
                                ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT  WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                                ,(SELECT COUNT(TA001) FROM [TK].dbo.POSTA WITH(NOLOCK)  WHERE  TA002=View_t_visitors.TT002 AND TA004=View_t_visitors.Fdate1) AS 'COUNTSTA001'
                                ,(SELECT SUM(TB019) FROM [TK].dbo.POSTB  WITH(NOLOCK) WHERE  TB002=View_t_visitors.TT002 AND TB004=View_t_visitors.Fdate1 AND TB010 NOT LIKE '1%'  AND TB010 NOT LIKE '2%'  AND TB010 NOT LIKE '3%') AS 'SUMSTB019'
                                FROM [TKMK].[dbo].[View_t_visitors]
                                WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
                                AND CONVERT(NVARCHAR,Fdate1,112)>='{0}'
                                AND CONVERT(NVARCHAR,Fdate1,112)<='{1}'
                                GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK

                                UNION ALL
                                SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fout_data) AS SUMNUMS
                                ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                                ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WITH(NOLOCK)  WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                                ,(SELECT COUNT(TA001) FROM [TK].dbo.POSTA WITH(NOLOCK)  WHERE  TA002=View_t_visitors.TT002 AND TA004=View_t_visitors.Fdate1) AS 'COUNTSTA001'
                                ,(SELECT SUM(TB019) FROM [TK].dbo.POSTB  WITH(NOLOCK) WHERE  TB002=View_t_visitors.TT002 AND TB004=View_t_visitors.Fdate1 AND TB010 NOT LIKE '1%'  AND TB010 NOT LIKE '2%'  AND TB010 NOT LIKE '3%') AS 'SUMSTB019'
                                FROM [TKMK].[dbo].[View_t_visitors]
                                WHERE  TT002 IN ('106701') 
                                AND CONVERT(NVARCHAR,Fdate1,112)>='{0}'
                                AND CONVERT(NVARCHAR,Fdate1,112)<='{1}'
              
                                GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK
                                ) AS TEMP
                                ORDER BY TT002,Fdate1
                                

                                ", SDAYS, EDAYS);

            SQL2.AppendFormat(@" 
                               
                               SELECT 門代,門店,SUM(來客數) 來客數,SUM(銷售未稅總金額) 銷售未稅總金額,SUM(結帳單量) 結帳單量,SUM(結帳交易商品數) 結帳交易商品數
                                ,(SUM(結帳單量) * 1.0 / NULLIF(SUM(來客數), 0)) AS 提袋率
                                ,(SUM(結帳交易商品數) * 1.0 / NULLIF(SUM(結帳單量), 0)) AS 結帳交易商品數
                                ,(SUM(銷售未稅總金額) * 1.0 / NULLIF(SUM(結帳單量), 0)) AS 客單價
                                ,(SUM(結帳交易商品數) * 1.0 / NULLIF(SUM(結帳單量), 0)) AS 每單平均商品數
                                FROM 
                                (
                                SELECT 
                                TT002 AS '門代'
                                ,STORESNAME AS '門店'
                                ,YEARS AS '年'
                                ,MONTHS AS '月'
                                ,WEEKS AS '週'
                                ,Fdate1 AS '日'
                                ,DAYOFWEEK AS '星期'
                                ,SUMNUMS AS '來客數'
                                ,CONVERT(INT,SUMTT018) AS '銷售未稅總金額'
                                ,COUNTSTA001 AS '結帳單量'
                                ,CONVERT(INT,SUMSTB019) AS '結帳交易商品數'
                                ,(CASE WHEN SUMNUMS>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,2),((CONVERT(DECIMAL(16,4),COUNTSTA001)/CONVERT(DECIMAL(16,4),SUMNUMS)))) ELSE 0 END ) AS '提袋率'
                                ,(CASE WHEN SUMTT018>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,0),SUMTT018/COUNTSTA001) ELSE 0 END ) AS '客單價'
                                ,(CASE WHEN SUMSTB019>0 AND COUNTSTA001>0 THEN CONVERT(DECIMAL(16,2),SUMSTB019/COUNTSTA001) ELSE 0 END ) AS '每單平均商品數'

                                FROM 
                                (
                                SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fin_data+Fout_data)/2 AS SUMNUMS
                                ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT018'
                                ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT  WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                                ,(SELECT COUNT(TA001) FROM [TK].dbo.POSTA WITH(NOLOCK)  WHERE  TA002=View_t_visitors.TT002 AND TA004=View_t_visitors.Fdate1) AS 'COUNTSTA001'
                                ,(SELECT SUM(TB019) FROM [TK].dbo.POSTB  WITH(NOLOCK) WHERE  TB002=View_t_visitors.TT002 AND TB004=View_t_visitors.Fdate1 AND TB010 NOT LIKE '1%'  AND TB010 NOT LIKE '2%'  AND TB010 NOT LIKE '3%') AS 'SUMSTB019'
                                FROM [TKMK].[dbo].[View_t_visitors]
                                WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
                                AND CONVERT(NVARCHAR,Fdate1,112)>='{0}'
                                AND CONVERT(NVARCHAR,Fdate1,112)<='{1}'
                                GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK

                                UNION ALL
                                SELECT View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK,SUM(Fout_data) AS SUMNUMS
                                ,(SELECT SUM(TT018) FROM [TK].dbo.POSTT WITH(NOLOCK) WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT011'
                                ,(SELECT SUM(TT008) FROM [TK].dbo.POSTT WITH(NOLOCK)  WHERE View_t_visitors.TT002=POSTT.TT002 AND View_t_visitors.Fdate1=POSTT.TT001) AS 'SUMTT008'
                                ,(SELECT COUNT(TA001) FROM [TK].dbo.POSTA WITH(NOLOCK)  WHERE  TA002=View_t_visitors.TT002 AND TA004=View_t_visitors.Fdate1) AS 'COUNTSTA001'
                                ,(SELECT SUM(TB019) FROM [TK].dbo.POSTB  WITH(NOLOCK) WHERE  TB002=View_t_visitors.TT002 AND TB004=View_t_visitors.Fdate1 AND TB010 NOT LIKE '1%'  AND TB010 NOT LIKE '2%'  AND TB010 NOT LIKE '3%') AS 'SUMSTB019'
                                FROM [TKMK].[dbo].[View_t_visitors]
                                WHERE  TT002 IN ('106701') 
                                AND CONVERT(NVARCHAR,Fdate1,112)>='{0}'
                                AND CONVERT(NVARCHAR,Fdate1,112)<='{1}'
              
                                GROUP BY View_t_visitors.TT002,STORESNAME,YEARS,MONTHS,WEEKS,Fdate1,DAYOFWEEK
                                ) AS TEMP1
                                ) AS TEMP
                                GROUP BY 門代,門店
                                ORDER BY 門代,門店
                                

                                ", SDAYS, EDAYS);

            SQL3.AppendFormat(@"
                                SELECT 
                                    MIN(TA001) 起日,
                                    MAX(TA001) 迄日,
                                    TA002 門市ID,
                                    ME002 門市,  
                                    COUNT(TA001) AS '總交易筆數',
                                    -- 核心修正：不符合條件就不給值 (預設為 NULL)，這樣 COUNT 才會準
                                    COUNT(CASE WHEN (TA008 LIKE '68%' OR TA008 LIKE '69%') THEN 1 END) AS '團客交易筆數',
                                    COUNT(TA001)-COUNT(CASE WHEN (TA008 LIKE '68%' OR TA008 LIKE '69%') THEN 1 END) AS '散客交易筆數'
                                FROM [TK].dbo.POSTA WITH(NOLOCK)
                                INNER JOIN [TK].dbo.CMSME ON ME001=TA002
                                WHERE TA038 IN ('1','2','3')
                                    AND TA002 IN ('106701','106702')
                                    AND TA001 >= '{0}' AND TA001 <= '{1}'
                                GROUP BY TA002,ME002
                                ORDER BY TA002,ME002
                                ", SDAYS, EDAYS);


            Report report1 = new Report();
            report1.Load(@"REPORT\營銷-每日來客明細.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            report1.Dictionary.Connections[0].CommandTimeout = 120;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL2.ToString();
            TableDataSource table2 = report1.GetDataSource("Table2") as TableDataSource;
            table2.SelectCommand = SQL3.ToString();



            report1.Preview = previewControl4;
            report1.Show();
        }
        public void SETFASTREPORT8(string SDAYS, string EDAYS, string SETMONEYS)
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();

            SQL1.AppendFormat(@"   
                                --TA026<0 是含銷退
                            
                                SELECT 
                                *
                                ,(CASE WHEN 單筆消費滿額的筆數>0 AND 總消費筆數>0 THEN CONVERT(DECIMAL,單筆消費滿額的筆數)/ CONVERT(DECIMAL,總消費筆數) ELSE 0 END ) AS 'PCTS'
                                FROM 
                                (
	                                SELECT 
	                                '{0}'+'~'+'{1}'  AS '日期',
	                                TA002 AS '門市代',
	                                ME002 AS '門市',
	                                COUNT(TA002) AS '總消費筆數',
	                                SUM(TA026) AS '總銷售額(含稅)', 
	                                SUM(CASE WHEN POSTA.TA026 >= {2} THEN 1 ELSE 0 END) AS '單筆消費滿額的筆數'
	                                FROM [TK].dbo.POSTA,[TK].dbo.CMSME
	                                WHERE 1=1
	                                AND TA002=ME001 
                                    AND POSTA.TA026 >=0
                                    AND TA038 NOT IN ('4' )                              
	                                AND TA001>='{0}' AND TA001<='{1}'
	                                GROUP BY TA002,ME002
                                ) AS TEMP
                                ORDER BY 門市代,門市
 
                                ", SDAYS, EDAYS, SETMONEYS);
             


            Report report1 = new Report();
            report1.Load(@"REPORT\銷售筆數表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            report1.Dictionary.Connections[0].CommandTimeout = 120;

            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", SETMONEYS);



            report1.Preview = previewControl7;
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
                                    AND [t_STORESNAME].[ISUSED]='Y'
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

        public void SETFASTREPORT5(string SDATE, string EDATE)
        {
            StringBuilder SQL1 = new StringBuilder();



            SQL1.AppendFormat(@"
                                SELECT TA002 AS '門市代號',MA002  AS '門市',TA004 AS '交易日期',TA005 AS '交易時間',TA026 AS '交易金額'
                                FROM [TK].dbo.POSTA WITH(NOLOCK)
                                LEFT JOIN [TK].dbo.WSCMA ON MA001=TA002
                                WHERE 1=1
                                AND TA002 IN ('106501','106502','106503','106504','106701','106702','106704','106705')
                                AND  TA004>='{0}' AND TA004<='{1}'
                                ORDER BY TA002,TA004,TA005,TA006

                                ", SDATE, EDATE);


            Report report1 = new Report();
            report1.Load(@"REPORT\營銷-每日POS明細.frx");

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



            report1.Preview = previewControl5;
            report1.Show();
        }

        public void SETFASTREPORT6(string SDATE, string EDATE)
        {
            StringBuilder SQL1 = new StringBuilder(); 



            SQL1.AppendFormat(@"
                               SELECT 門市代號,門市
                                ,SUM(M500) AS 'M500' 
                                ,SUM(M1000) AS 'M1000' 
                                ,SUM(M1500) AS 'M1500' 
                                ,SUM(M2000) AS 'M2000' 
                                ,SUM(M2500) AS 'M2500' 
                                ,SUM(M3000) AS 'M3000' 
                                ,SUM(M3500) AS 'M3500' 
                                ,SUM(M4000) AS 'M4000' 
                                ,SUM(M4500) AS 'M4500' 
                                ,SUM(M5000) AS 'M5000' 
                                ,SUM(M5500) AS 'M5500' 
                                ,SUM(M6000) AS 'M6000' 
                                ,SUM(M6500) AS 'M6500' 
                                ,SUM(M7000) AS 'M7000' 
                                ,SUM(M7500) AS 'M7500' 
                                ,SUM(M8000) AS 'M8000' 
                                ,SUM(M8500) AS 'M8500' 
                                ,SUM(M9000) AS 'M9000' 
                                ,SUM(M9500) AS 'M9500' 
                                ,SUM(M10000) AS 'M10000' 
                                ,SUM(M20000) AS 'M20000' 
                                ,SUM(M30000) AS 'M30000' 
                                ,SUM(M30001) AS 'M30001' 
                                FROM 
                                (
                                SELECT TA002 AS '門市代號',MA002  AS '門市',TA004 AS '交易日期',TA005 AS '交易時間',TA026 AS '交易金額'
                                , (CASE WHEN TA026>0 AND TA026<500 THEN 1 ELSE 0 END ) AS 'M500'
                                , (CASE WHEN TA026>=500 AND TA026<1000 THEN 1 ELSE 0 END ) AS 'M1000'
                                , (CASE WHEN TA026>=1000 AND TA026<1500 THEN 1 ELSE 0 END ) AS 'M1500'
                                , (CASE WHEN TA026>=1500 AND TA026<2000 THEN 1 ELSE 0 END ) AS 'M2000'
                                , (CASE WHEN TA026>=2000 AND TA026<2500 THEN 1 ELSE 0 END ) AS 'M2500'
                                , (CASE WHEN TA026>=2500 AND TA026<3000 THEN 1 ELSE 0 END ) AS 'M3000'
                                , (CASE WHEN TA026>=3000 AND TA026<3500 THEN 1 ELSE 0 END ) AS 'M3500'
                                , (CASE WHEN TA026>=3500 AND TA026<4000 THEN 1 ELSE 0 END ) AS 'M4000'
                                , (CASE WHEN TA026>=4000 AND TA026<4500 THEN 1 ELSE 0 END ) AS 'M4500'
                                , (CASE WHEN TA026>=4500 AND TA026<5000 THEN 1 ELSE 0 END ) AS 'M5000'
                                , (CASE WHEN TA026>=5000 AND TA026<5500 THEN 1 ELSE 0 END ) AS 'M5500'
                                , (CASE WHEN TA026>=5500 AND TA026<6000 THEN 1 ELSE 0 END ) AS 'M6000'
                                , (CASE WHEN TA026>=6000 AND TA026<6500 THEN 1 ELSE 0 END ) AS 'M6500'
                                , (CASE WHEN TA026>=6500 AND TA026<7000 THEN 1 ELSE 0 END ) AS 'M7000'
                                , (CASE WHEN TA026>=7000 AND TA026<7500 THEN 1 ELSE 0 END ) AS 'M7500'
                                , (CASE WHEN TA026>=7500 AND TA026<8000 THEN 1 ELSE 0 END ) AS 'M8000'
                                , (CASE WHEN TA026>=8000 AND TA026<8500 THEN 1 ELSE 0 END ) AS 'M8500'
                                , (CASE WHEN TA026>=8500 AND TA026<9000 THEN 1 ELSE 0 END ) AS 'M9000'
                                , (CASE WHEN TA026>=9000 AND TA026<9500 THEN 1 ELSE 0 END ) AS 'M9500'
                                , (CASE WHEN TA026>=9500 AND TA026<10000 THEN 1 ELSE 0 END ) AS 'M10000'
                                , (CASE WHEN TA026>=10000 AND TA026<20000 THEN 1 ELSE 0 END ) AS 'M20000'
                                , (CASE WHEN TA026>=20000 AND TA026<30000 THEN 1 ELSE 0 END ) AS 'M30000'
                                , (CASE WHEN TA026>=30000  THEN 1 ELSE 0 END ) AS 'M30001'


                                FROM [TK].dbo.POSTA WITH(NOLOCK)
                                LEFT JOIN [TK].dbo.WSCMA ON MA001=TA002
                                WHERE 1=1
                                AND TA002 IN ('106501','106502','106503','106504','106701','106702','106704','106705')
                                AND  TA004>='{0}' AND TA004<='{1}'
                                ) AS  TEMP 
                                GROUP BY 門市代號,門市
                                ORDER BY 門市代號,門市

                                ", SDATE, EDATE);


            Report report1 = new Report();
            report1.Load(@"REPORT\營銷-每日POS明細MATRIX.frx");

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

           


            report1.Preview = previewControl6;
            report1.Show();
        }

        public void ADD_Visitors_Monthly(string YEARS,string MONTHS)
        {
            string SDATES = YEARS + "0101";
            string EDATES = YEARS + "1231";
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

                StringBuilder SQLEXE=new StringBuilder();
                SQLEXE.AppendFormat(@"
                DELETE   [TKMK].[dbo].[Visitors_Monthly] WHERE [年度]=@YEARS;

                WITH 
                -- 預先統計人流量 (按門市、年份、月份彙總)
                Visitors_Monthly_CTE AS (
                    SELECT 
                        TT002,
                        STORESNAME,
                        YEARS,
                        MONTHS,
                        SUM(
                            CASE 
                                WHEN TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
                                    THEN (Fin_data + Fout_data) / 2.0
                                WHEN TT002 = '106701' 
                                    THEN Fout_data
                                ELSE 0
                            END
                        ) AS SUMNUMS
                    FROM [TKMK].[dbo].[View_t_visitors] WITH(NOLOCK)
                    WHERE YEARS = @YEARS
                        AND MONTHS<=@MONTHS
                      AND TT002 IN ('106501','106502','106503','106504','106513','106701','106702','106703','106704','106705')
                    GROUP BY TT002, STORESNAME, YEARS, MONTHS
                ),

                -- 預先統計 POS 銷售額 (按門市、年份、月份彙總)
                POSTT_Monthly_CTE AS (
                    SELECT 
                        TT002,
                        LEFT(TT001, 4) AS YEARS,
                        CAST(SUBSTRING(TT001, 5, 2) AS INT) AS MONTHS_INT,
                        SUM(TT008) AS REALSUMTT008, -- 總成交筆數
                        SUM(TT018) AS REALSUMTT018  -- 總銷售金額
                    FROM [TK].dbo.POSTT WITH(NOLOCK)
                    WHERE TT001 >= @SDATES
                      AND TT001 <= @EDATES
                      AND TT002 IN ('106501','106502','106503','106504','106513','106701','106702','106703','106704','106705')
                    GROUP BY TT002, LEFT(TT001, 4), CAST(SUBSTRING(TT001, 5, 2) AS INT)
                )

                -- 執行寫入
                INSERT INTO [TKMK].[dbo].[Visitors_Monthly]
                (
                    [代號],
                    [門市],
                    [年度],
                    [月份],
                    [來客數],
                    [銷售總金額POS機],
                    [成交筆數],
                    [提袋率],
                    [平均客單價],
                    [實際的成交筆數],
                    [實際的銷售總金額POS機],
                    [實際的平均客單價]
                )
                SELECT 
                    V.TT002 AS 代號,
                    V.STORESNAME AS 門市,
                    V.YEARS AS 年度,
                    V.MONTHS AS 月份,
                    ISNULL(V.SUMNUMS, 0) AS 來客數,
                    ISNULL(P.REALSUMTT018, 0) AS 銷售總金額POS機,
                    ISNULL(P.REALSUMTT008, 0) AS 成交筆數,
    
                    -- 平均與比例計算（乘上 1.0 確保轉為浮點數計算）
                    ISNULL(P.REALSUMTT008, 0) * 1.0 / NULLIF(V.SUMNUMS, 0) AS 提袋率,
                    ISNULL(P.REALSUMTT018, 0) * 1.0 / NULLIF(P.REALSUMTT008, 0) AS 平均客單價,
    
                    ISNULL(P.REALSUMTT008, 0) AS 實際的成交筆數,
                    ISNULL(P.REALSUMTT018, 0) AS 實際的銷售總金額POS機,
                    ISNULL(P.REALSUMTT018, 0) * 1.0 / NULLIF(P.REALSUMTT008, 0) AS 實際的平均客單價

                FROM Visitors_Monthly_CTE V
                LEFT JOIN POSTT_Monthly_CTE P 
                       ON V.TT002 = P.TT002 
                      AND V.YEARS = P.YEARS 
                      AND CAST(V.MONTHS AS INT) = P.MONTHS_INT;

                ");

                using (SqlCommand cmd = new SqlCommand(SQLEXE.ToString(), sqlConn))
                {
                    cmd.CommandTimeout = 300;
                    cmd.CommandType = CommandType.Text; // ✅ 關鍵修改：改為 CommandType.Text

                    cmd.Parameters.AddWithValue("@YEARS", YEARS);
                    cmd.Parameters.AddWithValue("@MONTHS", MONTHS);
                    cmd.Parameters.AddWithValue("@SDATES", SDATES);
                    cmd.Parameters.AddWithValue("@EDATES", EDATES);

                    sqlConn.Open();
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void ADD_Visitors_Weeks(string YEARS)
        {
            string SDATES = YEARS + "0101";
            string EDATES = YEARS + "1231";
            try
            {
                // 20210902 密碼解密
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    StringBuilder SQLEXE = new StringBuilder();
                    SQLEXE.AppendFormat(@"
                    DELETE FROM [TKMK].[dbo].[Visitors_Weeks] WHERE [年度] = @YEARS;

                    -- 1. CTE 宣告必須放在最前面（DELETE 後面加分號 ';'）
                    WITH Visitors_Daily AS (
                        SELECT 
                            TT002,
                            STORESNAME,
                            YEARS,
                            WEEKS,
                            Fdate1,
                            SUM(
                                CASE 
                                    WHEN TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
                                        THEN (Fin_data + Fout_data) / 2.0
                                    WHEN TT002 = '106701' 
                                        THEN Fout_data
                                    ELSE 0
                                END
                            ) AS NUMS
                        FROM [TKMK].[dbo].[View_t_visitors] WITH(NOLOCK)
                        WHERE YEARS = @YEARS
                          AND TT002 IN ('106501','106502','106503','106504','106513','106701','106702','106703','106704','106705')
  
                          -- 🔥 核心邏輯：若最新週 > 1 則排除最新週；若最新週 = 1 則保留第 1 週
                          AND CONVERT(INT, WEEKS) < (
                              SELECT 
                                  CASE 
                                      WHEN MAX(CONVERT(INT, WEEKS)) > 1 THEN MAX(CONVERT(INT, WEEKS))
                                      ELSE 2
                                  END
                              FROM [TKMK].[dbo].[View_t_visitors] WITH(NOLOCK)
                              WHERE YEARS = @YEARS
                          )

                        GROUP BY TT002, STORESNAME, YEARS, WEEKS, Fdate1
                    ),

                    POSTT_Daily AS (
                        SELECT 
                            TT002,
                            TT001 AS Fdate,
                            SUM(TT018) AS SUMTT011,
                            SUM(TT008) AS SUMTT008
                        FROM [TK].dbo.POSTT WITH(NOLOCK)
                        WHERE TT001 >= @SDATES AND TT001 <= @EDATES
                          AND TT002 IN ('106501','106502','106503','106504','106513','106701','106702','106703','106704','106705')
                        GROUP BY TT002, TT001
                    )

                    -- 2. INSERT INTO 必須移動到這裡（CTE 與 SELECT 中間）
                    INSERT INTO [TKMK].[dbo].[Visitors_Weeks]
                    (
                        [代號],
                        [門市],
                        [年度],
                        [週次],
                        [來客數],
                        [銷售總金額POS機],
                        [成交筆數],
                        [提袋率],
                        [平均客單價]
                    )
                    SELECT 
                        V.TT002 AS 代號,
                        V.STORESNAME AS 門市,
                        V.YEARS AS 年度,
                        V.WEEKS AS 週次,
                        SUM(V.NUMS) AS 來客數,
                        ISNULL(SUM(P.SUMTT011), 0) AS 銷售總金額POS機,
                        ISNULL(SUM(P.SUMTT008), 0) AS 成交筆數,

                        -- 乘以 1.0 確保轉為浮點數計算，避免小數點被捨去
                        ISNULL(SUM(P.SUMTT008), 0) * 1.0 / NULLIF(SUM(V.NUMS), 0) AS 提袋率,
                        ISNULL(SUM(P.SUMTT011), 0) * 1.0 / NULLIF(SUM(P.SUMTT008), 0) AS 平均客單價

                    FROM Visitors_Daily V
                    LEFT JOIN POSTT_Daily P 
                           ON V.TT002 = P.TT002 
                          AND V.Fdate1 = P.Fdate

                    GROUP BY V.TT002, V.STORESNAME, V.YEARS, V.WEEKS
                    ORDER BY V.TT002, V.STORESNAME, V.YEARS, CONVERT(INT, V.WEEKS);
                    ");

                    using (SqlCommand cmd = new SqlCommand(SQLEXE.ToString(), sqlConn))
                    {
                        cmd.CommandTimeout = 300;
                        cmd.CommandType = CommandType.Text;

                        cmd.Parameters.AddWithValue("@YEARS", YEARS);
                        cmd.Parameters.AddWithValue("@SDATES", SDATES);
                        cmd.Parameters.AddWithValue("@EDATES", EDATES);

                        sqlConn.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void ADD_Visitors_Month_Hours(string YEARS,string MONTHS)
        {
            string SDATES = YEARS + "0101";
            string EDATES = YEARS + "1231";
            try
            {
                // 20210902 密碼解密
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    StringBuilder SQLEXE = new StringBuilder();
                    SQLEXE.AppendFormat(@"
                        DELETE [TKMK].[dbo].[Visitors_Month_Hours] WHERE [年度]=@YEARS AND [月份]=@MONTHS;

                        WITH 
                        -- 1. 預先按「門市 + 日期 + 時段」加總人流量，並處理月份總天數（舊版語法相容）
                        Visitors_Hourly AS (
                            SELECT 
                                TT002,
                                STORESNAME,
                                YEARS,
                                MONTHS,
                                Fdate1,
                                HOURS,
                                -- 相容舊版 SQL Server 的「該月總天數」計算
                                DAY(DATEADD(month, DATEDIFF(month, 0, CONVERT(DATETIME, CONVERT(NVARCHAR(4), YEARS) + '/' + CONVERT(NVARCHAR(4), MONTHS) + '/1')) + 1, -1)) AS DAYSS,
                                SUM(
                                    CASE 
                                        WHEN TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
                                            THEN (Fin_data + Fout_data) / 2.0
                                        WHEN TT002 = '106701' 
                                            THEN Fout_data
                                        ELSE 0
                                    END
                                ) AS NUMS
                            FROM [TKMK].[dbo].[View_t_visitors] WITH(NOLOCK)
                            WHERE YEARS = @YEARS 
                              AND MONTHS = @MONTHS
                              AND TT002 IN ('106501','106502','106503','106504','106513','106701','106702','106703','106704','106705')
                            GROUP BY TT002, STORESNAME, YEARS, MONTHS, Fdate1, HOURS
                        ),

                        -- 2. 預先按「門市 + 日期 + 時段」加總 POS 交易資料 (一次取出 SUM 與 COUNT)
                        POSTA_Hourly AS (
                            SELECT 
                                TA002 AS TT002,
                                TA004 AS Fdate1,
                                HHS,
                                SUM(ISNULL(TA026, 0)) AS SUMTA026,
                                COUNT(TA026) AS COUNTSTA026
                            FROM [TK].[dbo].[POSTA] WITH(NOLOCK)
                            WHERE TA004 >= @SDATES AND TA004 <= @EDATES
                              AND TA002 IN ('106501','106502','106503','106504','106513','106701','106702','106703','106704','106705')
                            GROUP BY TA002, TA004, HHS
                        )

                        INSERT INTO [TKMK].[dbo].[Visitors_Month_Hours]
                        (
                        [代號]
                        ,[門市]
                        ,[年度]
                        ,[月份]
                        ,[時段]
                        ,[天數]
                        ,[來客數]
                        ,[銷售總金額POS機]
                        ,[成交筆數]
                        ,[提袋率]
                        ,[平均客單價]
                        )

                        -- 3. 主查詢：按門市、年月、小時進行最終彙總與指標計算
                        SELECT 
                            V.TT002 代號,
                            V.STORESNAME 門市,
                            V.YEARS 年度,
                            V.MONTHS 月份,
                            V.HOURS 時段,
                            V.DAYSS 天數,
                            SUM(V.NUMS) AS 來客數,
                            ISNULL(SUM(P.SUMTA026), 0) AS 銷售總金額POS機,
                            ISNULL(SUM(P.COUNTSTA026), 0) AS 成交筆數,
    
                            -- 提袋率/轉化率 (PCTS) 計算 (使用 NULLIF 防範除以 0)
                            ROUND(
                                ISNULL(SUM(P.COUNTSTA026), 0) * 1.0 / NULLIF(SUM(V.NUMS), 0), 
                                4
                            ) AS 提袋率,
    
                            -- 平均客單價 (AVGTA026) 計算 (使用 NULLIF 防範除以 0)
                            (CASE WHEN ISNULL(SUM(P.COUNTSTA026), 0)>0 THEN  (ISNULL(SUM(P.SUMTA026), 0) / NULLIF(SUM(P.COUNTSTA026), 0)) ELSE 0 END ) AS 平均客單價

                        FROM Visitors_Hourly V
                        LEFT JOIN POSTA_Hourly P 
                               ON V.TT002 = P.TT002 
                              AND V.Fdate1 = P.Fdate1 
                              AND RIGHT('00' + CAST(V.HOURS AS VARCHAR), 2) = P.HHS

                        GROUP BY V.TT002, V.STORESNAME, V.YEARS, V.MONTHS, V.HOURS, V.DAYSS
                        ORDER BY V.TT002, V.STORESNAME, V.YEARS, V.MONTHS, CONVERT(INT, V.HOURS);

                    ");

                    using (SqlCommand cmd = new SqlCommand(SQLEXE.ToString(), sqlConn))
                    {
                        cmd.CommandTimeout = 300;
                        cmd.CommandType = CommandType.Text;

                        cmd.Parameters.AddWithValue("@YEARS", YEARS);
                        cmd.Parameters.AddWithValue("@MONTHS", MONTHS);
                        cmd.Parameters.AddWithValue("@SDATES", SDATES);
                        cmd.Parameters.AddWithValue("@EDATES", EDATES);

                        sqlConn.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void ADD_Visitors_Month_Days(string YEARS, string MONTHS)
        {
            string SDATES = YEARS + "0101";
            string EDATES = YEARS + "1231";
            try
            {
                // 20210902 密碼解密
                Class1 TKID = new Class1();
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                using (SqlConnection sqlConn = new SqlConnection(sqlsb.ConnectionString))
                {
                    StringBuilder SQLEXE = new StringBuilder();
                    SQLEXE.AppendFormat(@"
                    DELETE FROM [TKMK].[dbo].[Visitors_Month_Days] 
                    WHERE [年度] = @YEARS AND [月份] = @MONTHS;

                    WITH Visitors_Daily AS (
                        SELECT 
                            TT002,
                            STORESNAME,
                            YEARS,
                            MONTHS,
                            WEEKS,
                            Fdate1,
                            DAYOFWEEK,
                            DATEPART(WEEKDAY, Fdate1) AS WEEKDAY_ORIG,
                            SUM(
                                CASE 
                                    WHEN TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704','106705') 
                                        THEN (Fin_data + Fout_data) / 2.0
                                    WHEN TT002 = '106701' 
                                        THEN Fout_data
                                    ELSE 0
                                END
                            ) AS NUMS
                        FROM [TKMK].[dbo].[View_t_visitors] WITH(NOLOCK)
                        WHERE YEARS = @YEARS
                          AND MONTHS = @MONTHS
                          AND TT002 IN ('106501','106502','106503','106504','106513','106701','106702','106703','106704','106705')
                        GROUP BY TT002, STORESNAME, YEARS, MONTHS, WEEKS, Fdate1, DAYOFWEEK
                    ),

                    POSTT_Daily AS (
                        SELECT 
                            TT002,
                            TT001 AS Fdate1,
                            SUM(ISNULL(TT018, 0)) AS SUMTT011,
                            SUM(ISNULL(TT008, 0)) AS SUMTT008
                        FROM [TK].[dbo].[POSTT] WITH(NOLOCK)
                        WHERE TT001 >= @SDATES AND TT001 <= @EDATES
                          AND TT002 IN ('106501','106502','106503','106504','106513','106701','106702','106703','106704','106705')
                        GROUP BY TT002, TT001
                    )

                    -- 🔥 修正 1：補上 INSERT INTO 陳述式
                    INSERT INTO [TKMK].[dbo].[Visitors_Month_Days]
                    (
                        [代號],
                        [門市],
                        [年度],
                        [月份],
                        [週數],
                        [星期],
                        [總來客數],
                        [銷售總金額POS機],
                        [成交筆數],
                        [每週來客數],
                        [提袋率],
                        [平均客單價],
                        [WEEKDAY]
                    )
                    SELECT 
                        V.TT002 AS 代號,
                        V.STORESNAME AS 門市,
                        V.YEARS AS 年度,
                        V.MONTHS AS 月份,
    
                        COUNT(DISTINCT V.WEEKS) AS 週數,
                        V.DAYOFWEEK AS 星期,
    
                        SUM(V.NUMS) AS 總來客數,
                        ISNULL(SUM(P.SUMTT011), 0) AS 銷售總金額POS機,
                        ISNULL(SUM(P.SUMTT008), 0) AS 成交筆數,
    
                        -- 每週平均人流
                        ROUND(SUM(V.NUMS) * 1.0 / NULLIF(COUNT(DISTINCT V.WEEKS), 0), 2) AS 每週來客數,
    
                        -- 🔥 修正 2：乘以 1.0 確保轉為小數計算，並做 ROUND 四捨五入
                        ROUND(ISNULL(SUM(P.SUMTT008), 0) * 1.0 / NULLIF(SUM(V.NUMS), 0), 4) AS 提袋率,
    
                        -- 🔥 修正 2：乘以 1.0 避免整數相除
                        ROUND(ISNULL(SUM(P.SUMTT011), 0) * 1.0 / NULLIF(SUM(P.SUMTT008), 0), 2) AS 平均客單價,
                        CASE WHEN V.WEEKDAY_ORIG = 1 THEN 99 ELSE V.WEEKDAY_ORIG END AS WEEKDAY

                    FROM Visitors_Daily V
                    LEFT JOIN POSTT_Daily P 
                           ON V.TT002 = P.TT002 
                          AND V.Fdate1 = P.Fdate1

                    GROUP BY 
                        V.TT002, 
                        V.STORESNAME, 
                        V.YEARS, 
                        V.MONTHS, 
                        V.DAYOFWEEK, 
                        -- 🔥 修正 3：統一 GROUP BY 的 CASE WHEN 寫法
                        CASE WHEN V.WEEKDAY_ORIG = 1 THEN 99 ELSE V.WEEKDAY_ORIG END
                    ORDER BY 
                        V.TT002, 
                        V.STORESNAME, 
                        V.YEARS, 
                        V.MONTHS, 
                        CASE WHEN V.WEEKDAY_ORIG = 1 THEN 99 ELSE V.WEEKDAY_ORIG END, 
                        V.DAYOFWEEK;    
                    ");

                    using (SqlCommand cmd = new SqlCommand(SQLEXE.ToString(), sqlConn))
                    {
                        cmd.CommandTimeout = 300;
                        cmd.CommandType = CommandType.Text;

                        cmd.Parameters.AddWithValue("@YEARS", YEARS);
                        cmd.Parameters.AddWithValue("@MONTHS", MONTHS);
                        cmd.Parameters.AddWithValue("@SDATES", SDATES);
                        cmd.Parameters.AddWithValue("@EDATES", EDATES);

                        sqlConn.Open();
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void SETFASTREPORT9(string YEARS, string MONTHS)
        {
            Report report1 = new Report();
            report1.Load(@"REPORT\營銷來客報表v3.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            //TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            //table.SelectCommand = SQL1.ToString();



            report1.Preview = previewControl8;
            report1.Show();
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
            //SETFASTREPORT4(dateTimePicker5.Value.ToString("yyyyMMdd"),dateTimePicker8.Value.ToString("yyyyMMdd"));
            SETFASTREPORT7(dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SETFASTREPORT5(dateTimePicker6.Value.ToString("yyyyMMdd"), dateTimePicker7.Value.ToString("yyyyMMdd"));
            SETFASTREPORT6(dateTimePicker6.Value.ToString("yyyyMMdd"), dateTimePicker7.Value.ToString("yyyyMMdd"));
        }
        private void button7_Click(object sender, EventArgs e)
        {
            SETFASTREPORT8(dateTimePicker9.Value.ToString("yyyyMMdd"), dateTimePicker10.Value.ToString("yyyyMMdd"),textBox1.Text);
        }
        private void button8_Click(object sender, EventArgs e)
        {
            string YEARS=dateTimePicker11.Value.Year.ToString();
            string MONTHS = dateTimePicker11.Value.Month.ToString();

            ADD_Visitors_Monthly(YEARS, MONTHS);
            ADD_Visitors_Weeks(YEARS);
            ADD_Visitors_Month_Hours(YEARS, MONTHS);
            ADD_Visitors_Month_Days(YEARS, MONTHS);

            MessageBox.Show("完成");
        }
        private void button9_Click(object sender, EventArgs e)
        {
            string YEARS = dateTimePicker11.Value.Year.ToString();
            string MONTHS = dateTimePicker11.Value.Month.ToString();

            SETFASTREPORT9(YEARS, MONTHS);
        }
        #endregion


    }
}
