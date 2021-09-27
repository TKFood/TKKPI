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
using System.Net.Mail;
using TKITDLL;


namespace TKKPI
{
    public partial class frmRESALES : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        SqlTransaction tran;
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string talbename = null;
        int rownum = 0;
        int result;
        string DATES = null;
        string DirectoryNAME = null;
        string pathFileSALESMONEYS = null;

        public object ID1 { get; private set; }

        public frmRESALES()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();

            report1.Load(@"REPORT\國內、外業務部業績日報表.frx");

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

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            string FirstDay = dateTimePicker1.Value.ToString("yyyyMM")+"01";
            string LastDay = dateTimePicker1.Value.ToString("yyyyMM") + "31";

            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                          --20210910 業務員日 報表
                            --200050 張釋予
                            --140078 蔡顏鴻
                            --100005 何姍怡
                            --160155 洪櫻芬
                            --170007 林杏育
                            --120003 葉枋俐
                            SELECT 
                            DATES
                            ,國內張釋予銷貨
                            ,國內張釋予銷退
                            ,國內蔡顏鴻銷貨
                            ,國內蔡顏鴻銷退
                            ,國內何姍怡銷貨
                            ,國內何姍怡銷退
                            ,國內洪櫻芬銷貨
                            ,國內洪櫻芬銷退
                            ,國內林杏育銷貨
                            ,國內林杏育銷退
                            ,全聯銷貨
                            ,國外洪櫻芬銷貨
                            ,國外洪櫻芬銷退
                            ,國外葉枋俐銷貨
                            ,國外葉枋俐銷退
                            ,(國內張釋予銷貨+國內張釋予銷退+國內蔡顏鴻銷貨+國內蔡顏鴻銷退+國內何姍怡銷貨+國內何姍怡銷退+國內洪櫻芬銷貨+國內洪櫻芬銷退+國內林杏育銷貨+國內林杏育銷退+全聯銷貨) AS '國內業務合計'
                            ,(國外洪櫻芬銷貨+國外洪櫻芬銷退+國外葉枋俐銷貨+國外葉枋俐銷退) AS '國外業務合計'
                            ,(國內張釋予銷貨+國內張釋予銷退+國內蔡顏鴻銷貨+國內蔡顏鴻銷退+國內何姍怡銷貨+國內何姍怡銷退+國內洪櫻芬銷貨+國內洪櫻芬銷退+國內林杏育銷貨+國內林杏育銷退+全聯銷貨+國外洪櫻芬銷貨+國外洪櫻芬銷退+國外葉枋俐銷貨+國外葉枋俐銷退) AS '總計'
                            ,(SELECT ISNULL(INTARGETMONEYS,0) FROM [TK].[dbo].[ZTARGETMONEYS] WHERE YEARSMOTNS=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6)) AS '國內月目標業績'
                            ,(SELECT ISNULL([OUTTARGETMONEYS],0) FROM [TK].[dbo].[ZTARGETMONEYS] WHERE YEARSMOTNS=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6)) AS '國外月目標業績'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND SUBSTRING(TG003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%'))  AND TG006 IN ('200050','140078','100005','160155','170007') ) AS '國內月總銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND SUBSTRING(TI003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006 IN ('200050','140078','100005','160155','170007') ) AS '國內月總銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND SUBSTRING(TG003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TG023='Y' AND (TG004 LIKE '3%' OR TG004 LIKE 'A3%' OR TG004 LIKE 'B3%') AND TG006 IN ('160155','120003')) AS '國外月總銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND SUBSTRING(TI003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6)  AND TI019='Y'AND (TI004 LIKE '3%' OR TI004 LIKE 'A3%' OR TI004 LIKE 'B3%') AND TI006 IN ('160155','120003')) AS '國外月總銷退'
                            ,(((SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND SUBSTRING(TG003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%'))  AND TG006 IN ('200050','140078','100005','160155','170007') )+(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND SUBSTRING(TI003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006 IN ('200050','140078','100005','160155','170007') ))/(SELECT ISNULL(INTARGETMONEYS,0) FROM [TK].[dbo].[ZTARGETMONEYS] WHERE YEARSMOTNS=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6))) AS '國內月累績達成率'
                            ,(((SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND SUBSTRING(TG003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6) AND TG023='Y' AND (TG004 LIKE '3%' OR TG004 LIKE 'A3%' OR TG004 LIKE 'B3%') AND TG006 IN ('160155','120003'))+(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND SUBSTRING(TI003,1,6)=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6)  AND TI019='Y'AND (TI004 LIKE '3%' OR TI004 LIKE 'A3%' OR TI004 LIKE 'B3%') AND TI006 IN ('160155','120003')))/(SELECT ISNULL([OUTTARGETMONEYS],0) FROM [TK].[dbo].[ZTARGETMONEYS] WHERE YEARSMOTNS=SUBSTRING(CONVERT(nvarchar,DATES,112),1,6))) AS '國外月累績達成率'
                            FROM (
                            SELECT CONVERT(nvarchar,DATES,112) AS DATES
                            ,[RTSALEMONEYS]  AS '全聯銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%'))  AND TG006='200050') AS '國內張釋予銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006='200050') AS '國內張釋予銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TG006='140078') AS '國內蔡顏鴻銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006='140078') AS '國內蔡顏鴻銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TG006='100005') AS '國內何姍怡銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006='100005') AS '國內何姍怡銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TG006='160155') AS '國內洪櫻芬銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006='160155') AS '國內洪櫻芬銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '1%' OR TG004 LIKE '2%' OR TG004 LIKE '5%' OR TG004 LIKE 'A2%' OR TG004 LIKE 'B2%') AND (TG004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TG006='170007') AS '國內林杏育銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112) AND TI019='Y' AND (TI004 LIKE '1%' OR TI004 LIKE '2%' OR TI004 LIKE '5%' OR TI004 LIKE 'A2%' OR TI004 LIKE 'B2%') AND (TI004 NOT IN (SELECT MA001 FROM [TK].dbo.COPMA WHERE MA002 LIKE '%全聯%')) AND TI006='170007') AS '國內林杏育銷退'
                            ,'-' AS '-'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND (CASE WHEN ISDATE(COPTG.UDF01)=1 THEN COPTG.UDF01 ELSE TG003 END =CONVERT(nvarchar,DATES,112)) AND TG023='Y' AND (TG004 LIKE '3%' OR TG004 LIKE 'A3%' OR TG004 LIKE 'B3%') AND TG006='160155') AS '國外洪櫻芬銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112)  AND TI019='Y'AND (TI004 LIKE '3%' OR TI004 LIKE 'A3%' OR TI004 LIKE 'B3%') AND TI006='160155') AS '國外洪櫻芬銷退'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TH037),0)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG003=CONVERT(nvarchar,DATES,112) AND TG023='Y' AND (TG004 LIKE '3%' OR TG004 LIKE 'A3%' OR TG004 LIKE 'B3%') AND TG006='120003') AS '國外葉枋俐銷貨'
                            ,(SELECT CONVERT(INT,ISNULL(SUM(TJ033)*-1,0)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI003=CONVERT(nvarchar,DATES,112)  AND TI019='Y'AND (TI004 LIKE '3%' OR TI004 LIKE 'A3%' OR TI004 LIKE 'B3%') AND TI006='120003') AS '國外葉枋俐銷退'
                            FROM [TK].dbo.ZDATES
                            WHERE CONVERT(nvarchar,DATES,112)>='{0}' AND CONVERT(nvarchar,DATES,112)<='{1}'
                            ) AS TEMP
                            ORDER BY DATES 

                            ", FirstDay, LastDay);


            return SB;

        }

        public void SEARCHZDATES(string YYMM)
        {
            string FirstDay = YYMM + "01";
            string LastDay = YYMM + "31";

            try
            {
                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();                
                DataSet ds1 = new DataSet();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT CONVERT(NVARCHAR,[DATES],111) AS '日期',[RTSALEMONEYS] AS '全聯銷售金額'
                                    FROM [TK].[dbo].[ZDATES]
                                    WHERE  CONVERT(NVARCHAR,[DATES],112)>='{0}' AND  CONVERT(NVARCHAR,[DATES],112)<='{1}'
                                    ORDER BY [DATES]
                                    ", FirstDay, LastDay);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();

                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                    dataGridView1.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    dataGridView1.DataSource = null;
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    dateTimePicker3.Value= Convert.ToDateTime(row.Cells["日期"].Value.ToString());
                    textBox1.Text = row.Cells["全聯銷售金額"].Value.ToString();
                   
                }
                else
                {
                    textBox1.Text = "0";
                

                }
            }
        }

        public void UPDATEZDATES(string DATES,string RTSALEMONEYS)
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

             
                sbSql.AppendFormat(@" 
                                    UPDATE [TK].[dbo].[ZDATES]
                                    SET [RTSALEMONEYS]={1}
                                    WHERE CONVERT(NVARCHAR,[DATES],112)='{0}'
                                    ", DATES, RTSALEMONEYS);

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

        public void SEARCHZTARGETMONEYS(string YEARS)
        {
            try
            {
                SqlDataAdapter adapter1 = new SqlDataAdapter();
                SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
                DataSet ds1 = new DataSet();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  
                                    SELECT [YEARSMOTNS] AS '年月'
                                    ,[INTARGETMONEYS] AS '國內月目標業績'
                                    ,[OUTTARGETMONEYS] AS '國外月目標業績'
                                    FROM [TK].[dbo].[ZTARGETMONEYS]
                                    WHERE [YEARSMOTNS] LIKE '{0}%'
                                    ORDER BY [YEARSMOTNS]
                                    ", YEARS);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();

                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    //dataGridView1.Rows.Clear();
                    dataGridView2.DataSource = ds1.Tables["TEMPds1"];
                    dataGridView2.AutoResizeColumns();
                    //dataGridView1.CurrentCell = dataGridView1[0, rownum];

                }
                else if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    dataGridView2.DataSource = null;
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

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    dateTimePicker5.Value = Convert.ToDateTime(row.Cells["年月"].Value.ToString().Substring(0,4)+"/"+ row.Cells["年月"].Value.ToString().Substring(4, 2) + "/01");
                    textBox2.Text = row.Cells["國內月目標業績"].Value.ToString();
                    textBox3.Text = row.Cells["國外月目標業績"].Value.ToString();

                }
                else
                {
                    textBox2.Text = "0";
                    textBox3.Text = "0";


                }
            }
        }

        public void ADDZTARGETMONEYS(string YEARSMOTNS, string INTARGETMONEYS, string OUTTARGETMONEYS)
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


                sbSql.AppendFormat(@" 
                                    INSERT INTO [TK].[dbo].[ZTARGETMONEYS]
                                    ([YEARSMOTNS],[INTARGETMONEYS],[OUTTARGETMONEYS])
                                    VALUES
                                    ('{0}',{1},{2})
                                    ", YEARSMOTNS, INTARGETMONEYS, OUTTARGETMONEYS);

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
        public void UPDATEZTARGETMONEYS(string YEARSMOTNS, string INTARGETMONEYS, string OUTTARGETMONEYS)
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


                sbSql.AppendFormat(@" 
                                   UPDATE  [TK].[dbo].[ZTARGETMONEYS]
                                    SET [INTARGETMONEYS]={1},[OUTTARGETMONEYS]={2}
                                    WHERE [YEARSMOTNS]='{0}'
                                    ", YEARSMOTNS, INTARGETMONEYS, OUTTARGETMONEYS);

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

        public void SENDEMAIL()
        {
            DataSet dsSALESMONEYS = new DataSet();
            StringBuilder SUBJEST = new StringBuilder();
            StringBuilder BODY = new StringBuilder();

            SETPATH();

            SAVEREPORT(pathFileSALESMONEYS);

            dsSALESMONEYS = SERACHMAILSALESMONEYS();

            if (dsSALESMONEYS.Tables["dsSALESMONEYS"].Rows.Count >= 1)
            {
                SUBJEST.Clear();
                BODY.Clear();
                SUBJEST.AppendFormat(@"每日-業務單位業績日報表-" + DateTime.Now.ToString("yyyy/MM/dd"));
                BODY.AppendFormat("Dear SIR" + Environment.NewLine + "附件為每日業務單位業績日報表，請查收" + Environment.NewLine + " ");
                SENDMAIL(SUBJEST, BODY, dsSALESMONEYS, pathFileSALESMONEYS);
            }



        }
        public void SETPATH()
        {          
            DATES = DateTime.Now.ToString("yyyyMMdd");

            DirectoryNAME = @"C:\MQTEMP\" + DATES.ToString() + @"\";
            pathFileSALESMONEYS = @"C:\MQTEMP\" + DATES.ToString() + @"\" + "每日業務單位業績日報表" + DATES.ToString()+".pdf";

            //如果日期資料夾不存在就新增
            if (!Directory.Exists(DirectoryNAME))
            {
                //新增資料夾
                Directory.CreateDirectory(DirectoryNAME);
            }
           
           
        }

        public void SAVEREPORT(string pathFileSALESMONEYS)
        {
            string FILENAME = pathFileSALESMONEYS;
            //string FILENAME = @"C:\MQTEMP\20210915\每日業務單位業績日報表20210915.pdf";
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();

            report1.Load(@"REPORT\國內、外業務部業績日報表.frx");

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

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            

            // prepare a report
            report1.Prepare();
            // create an instance of HTML export filter
            FastReport.Export.Pdf.PDFExport export = new FastReport.Export.Pdf.PDFExport();
            // show the export options dialog and do the export
            report1.Export(export, FILENAME);

        }
        public DataSet SERACHMAILSALESMONEYS()
        {
            SqlDataAdapter adapterSALESMONEYS = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilderSALESMONEYS = new SqlCommandBuilder();
            DataSet dsSALESMONEYS = new DataSet();

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



                sbSql.Clear();
                sbSqlQuery.Clear();

                //sbSql.AppendFormat(@"  WHERE [SENDTO]='COP' AND [MAIL]='tk290@tkfood.com.tw' ");

                sbSql.AppendFormat(@"  
                                    SELECT [SENDTO],[MAIL] 
                                    FROM [TKMQ].[dbo].[MQSENDMAIL] 
                                    WHERE [SENDTO]='SALESMONEYS'  
                                    ");

                adapterSALESMONEYS = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilderSALESMONEYS = new SqlCommandBuilder(adapterSALESMONEYS);
                sqlConn.Open();
                dsSALESMONEYS.Clear();
                adapterSALESMONEYS.Fill(dsSALESMONEYS, "dsSALESMONEYS");
                sqlConn.Close();



                if (dsSALESMONEYS.Tables["dsSALESMONEYS"].Rows.Count >= 1)
                {
                    return dsSALESMONEYS;
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }

        public void SENDMAIL(StringBuilder Subject, StringBuilder Body, DataSet SEND, string Attachments)
        {
            string MySMTPCONFIG = ConfigurationManager.AppSettings["MySMTP"];
            string NAME = ConfigurationManager.AppSettings["NAME"];
            string PW = ConfigurationManager.AppSettings["PW"];

            System.Net.Mail.MailMessage MyMail = new System.Net.Mail.MailMessage();
            MyMail.From = new System.Net.Mail.MailAddress("tk660@tkfood.com.tw");

            //MyMail.Bcc.Add("密件副本的收件者Mail"); //加入密件副本的Mail          
            //MyMail.Subject = "每日訂單-製令追踨表"+DateTime.Now.ToString("yyyy/MM/dd");
            MyMail.Subject = Subject.ToString();
            //MyMail.Body = "<h1>Dear SIR</h1>" + Environment.NewLine + "<h1>附件為每日訂單-製令追踨表，請查收</h1>" + Environment.NewLine + "<h1>若訂單沒有相對的製令則需通知製造生管開立</h1>"; //設定信件內容
            MyMail.Body = Body.ToString();
            //MyMail.IsBodyHtml = true; //是否使用html格式

            System.Net.Mail.SmtpClient MySMTP = new System.Net.Mail.SmtpClient(MySMTPCONFIG, 25);
            MySMTP.Credentials = new System.Net.NetworkCredential(NAME, PW);

            Attachment attch = new Attachment(Attachments);
            MyMail.Attachments.Add(attch);
            

            try
            {
                foreach (DataRow od in SEND.Tables[0].Rows)
                {

                    MyMail.To.Add(od["MAIL"].ToString()); //設定收件者Email，多筆mail
                }

                //MyMail.To.Add("tk290@tkfood.com.tw"); //設定收件者Email

                MySMTP.Send(MyMail);

                MyMail.Dispose(); //釋放資源


            }
            catch (Exception ex)
            {
                ADDLOG(DateTime.Now, Subject.ToString(), ex.ToString());
                //ex.ToString();
            }
        }

        public void ADDLOG(DateTime DATES, string SOURCE, string EX)
        {
            Guid NEWGUID = new Guid();
            NEWGUID = Guid.NewGuid();

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
                                    INSERT INTO [TKMQ].[dbo].[LOG]
                                    ([ID],[DATES],[SOURCE],[EX])
                                    VALUES 
                                    ('{0}','{1}','{2}','{3}')
                                   ", NEWGUID, DATES.ToString("yyyy/MM/dd HH:mm:ss"), SOURCE, EX);



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
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SEARCHZDATES(dateTimePicker2.Value.ToString("yyyyMM"));

        }
        private void button3_Click(object sender, EventArgs e)
        {
            int n;
            if (Int32.TryParse(textBox1.Text.ToString(), out n))
            {
                UPDATEZDATES(dateTimePicker3.Value.ToString("yyyyMMdd"), textBox1.Text.ToString());
            }

            SEARCHZDATES(dateTimePicker2.Value.ToString("yyyyMM"));

        }

        private void button4_Click(object sender, EventArgs e)
        {
            SEARCHZTARGETMONEYS(dateTimePicker4.Value.ToString("yyyy"));
        }


        private void button5_Click(object sender, EventArgs e)
        {
            int n;
            if (Int32.TryParse(textBox2.Text.ToString(), out n)&& Int32.TryParse(textBox3.Text.ToString(), out n))
            {
                ADDZTARGETMONEYS(dateTimePicker5.Value.ToString("yyyyMM"), textBox2.Text, textBox3.Text);
            }

            SEARCHZTARGETMONEYS(dateTimePicker4.Value.ToString("yyyy"));

        }
        private void button6_Click(object sender, EventArgs e)
        {
            int n;
            if (Int32.TryParse(textBox2.Text.ToString(), out n) && Int32.TryParse(textBox3.Text.ToString(), out n))
            {
                UPDATEZTARGETMONEYS(dateTimePicker5.Value.ToString("yyyyMM"), textBox2.Text, textBox3.Text);
            }

            SEARCHZTARGETMONEYS(dateTimePicker4.Value.ToString("yyyy"));
        }

        private void button7_Click(object sender, EventArgs e)
        {
            SENDEMAIL();

            MessageBox.Show("完成");
        }
        #endregion


    }
}
