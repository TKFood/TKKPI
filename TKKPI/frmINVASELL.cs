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

namespace TKKPI
{
    public partial class frmINVASELL : Form
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
        DataTable dt = new DataTable();
        string talbename = null;
        int rownum = 0;



        public frmINVASELL()
        {
            InitializeComponent();

            SETDATES();
        }

        #region FUNCTION
        public void SETDATES()
        {
            DateTime FirstDay = DateTime.Now.AddDays(-DateTime.Now.Day + 1);
            dateTimePicker1.Value = FirstDay;
            dateTimePicker2.Value = FirstDay;

            textBox1.Text = Math.Round(new TimeSpan(DateTime.Now.Ticks - FirstDay.Ticks).TotalDays,0).ToString();
            textBox2.Text = Math.Round(new TimeSpan(DateTime.Now.Ticks - FirstDay.Ticks).TotalDays, 0).ToString();
        }

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();

            report1.Load(@"REPORT\門市銷售預估月份.frx");

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
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                            DECLARE @SDAY nvarchar(10)
                            DECLARE @TOTALDAYS INT
                            SET @SDAY='{0}'
                            SET @TOTALDAYS={1}

                            SELECT LA001 AS '品號',MB002 AS '品名',LA016 AS '批號',NUMS AS '庫存量',有效日期,製造日期,總銷售數量,平均天銷售數量,預計銷售天,預計完銷日
                            ,DATEDIFF (MONTH,製造日期,預計完銷日) AS '生產到完銷的月數'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA LA WITH (NOLOCK) WHERE LA.LA001=TEMP2.LA001 AND LA.LA016=TEMP2.LA016 AND LA009 IN ('30001')) AS '中山一店'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA LA WITH (NOLOCK) WHERE LA.LA001=TEMP2.LA001 AND LA.LA016=TEMP2.LA016 AND LA009 IN ('30002')) AS '概念二店'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA LA WITH (NOLOCK) WHERE LA.LA001=TEMP2.LA001 AND LA.LA016=TEMP2.LA016 AND LA009 IN ('30003')) AS '北港三店'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA LA WITH (NOLOCK) WHERE LA.LA001=TEMP2.LA001 AND LA.LA016=TEMP2.LA016 AND LA009 IN ('30004')) AS '站前四店'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA LA WITH (NOLOCK) WHERE LA.LA001=TEMP2.LA001 AND LA.LA016=TEMP2.LA016 AND LA009 IN ('30012')) AS '微風北車店'
                            ,(SELECT ISNULL(SUM(LA005*LA011),0) FROM [TK].dbo.INVLA LA WITH (NOLOCK) WHERE LA.LA001=TEMP2.LA001 AND LA.LA016=TEMP2.LA016 AND LA009 IN ('30017')) AS '大潤發中崙店'
                            ,@SDAY AS '銷售日起'
                            ,@TOTALDAYS  AS '銷售天數'
                            FROM (
                            SELECT LA001,MB002,LA016,NUMS,有效日期,製造日期,總銷售數量,平均天銷售數量,CASE WHEN 平均天銷售數量>0 THEN (NUMS/平均天銷售數量) ELSE -1 END '預計銷售天'
                            ,CASE WHEN 平均天銷售數量>0 THEN CONVERT(NVARCHAR,DATEADD(DAY,CEILING(NUMS/平均天銷售數量),GETDATE()),112) ELSE '' END AS '預計完銷日'

                            FROM (
                            SELECT LA001,MB002,LA016,SUM(LA005*LA011) AS 'NUMS'
                            ,(SELECT TOP 1 TG018 FROM [TK].dbo.MOCTF WITH (NOLOCK) ,[TK].dbo.MOCTG WITH (NOLOCK)  WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 ORDER BY TG018 ) AS '有效日期'
                            ,(SELECT TOP 1 TG040 FROM [TK].dbo.MOCTF WITH (NOLOCK) ,[TK].dbo.MOCTG WITH (NOLOCK) WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 ORDER BY TG040 ) AS '製造日期'
                            ,(SELECT ISNULL(SUM(TB019),0) FROM [TK].dbo.POSTB WITH (NOLOCK) WHERE TB002 IN ('106501','106502','106503','106504','106513') AND TB010=LA001 AND TB001>=@SDAY) AS '總銷售數量'
                            ,(SELECT ISNULL(SUM(TB019),0) FROM [TK].dbo.POSTB WITH (NOLOCK) WHERE TB002 IN ('106501','106502','106503','106504','106513') AND TB010=LA001 AND TB001>=@SDAY)/@TOTALDAYS AS '平均天銷售數量'
                            FROM [TK].dbo.INVLA WITH (NOLOCK) ,[TK].dbo.INVMB WITH (NOLOCK) 
                            WHERE LA009 IN ('30001','30002','30003','30004','30012','30017')
                            AND LA001=MB001
                            AND LA001 LIKE '40%'
                            AND LA016 LIKE '2%'
                            AND MB002 NOT LIKE '%試吃%'
                            GROUP BY LA001,MB002,LA016
                            HAVING SUM(LA005*LA011)>0

                            ) AS TEMP
                            ) AS TEMP2
                            ORDER BY LA001

                            ", dateTimePicker1.Value.ToString("yyyyMMdd"),textBox1.Text.ToString());
                        

            return SB;

        }

        public void SETFASTREPORT2()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2();
            Report report1 = new Report();

            report1.Load(@"REPORT\營銷銷售預估月份.frx");

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

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"  
                             DECLARE @SDAY nvarchar(10)
                            DECLARE @TOTALDAYS INT
                            SET @SDAY='{0}'
                            SET @TOTALDAYS={1}

                            SELECT LA001 AS '品號',MB002 AS '品名',LA016 AS '批號',NUMS AS '庫存量',有效日期,製造日期,總銷售數量,平均天銷售數量,預計銷售天,預計完銷日
                            ,DATEDIFF (MONTH,製造日期,預計完銷日) AS '生產到完銷的月數'
                            ,@SDAY AS '銷售日起'
                            ,@TOTALDAYS  AS '銷售天數'
                            FROM (
                            SELECT LA001,MB002,LA016,NUMS,有效日期,製造日期,總銷售數量,平均天銷售數量,CASE WHEN 平均天銷售數量>0 THEN (NUMS/平均天銷售數量) ELSE -1 END '預計銷售天'
                            ,CASE WHEN 平均天銷售數量>0 THEN CONVERT(NVARCHAR,DATEADD(DAY,CEILING(NUMS/平均天銷售數量),GETDATE()),112) ELSE '' END AS '預計完銷日'
   
                            FROM (
                            SELECT LA001,MB002,LA016,SUM(LA005*LA011) AS 'NUMS'
                            ,(SELECT TOP 1 TG018 FROM [TK].dbo.MOCTF WITH (NOLOCK) ,[TK].dbo.MOCTG WITH (NOLOCK) WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 ORDER BY TG018 ) AS '有效日期'
                            ,(SELECT TOP 1 TG040 FROM [TK].dbo.MOCTF WITH (NOLOCK) ,[TK].dbo.MOCTG WITH (NOLOCK) WHERE TF001=TG001 AND TF002=TG002 AND TG004=LA001 AND TG017=LA016 ORDER BY TG040 ) AS '製造日期'
                            ,(SELECT ISNULL(SUM(TB019),0) FROM [TK].dbo.POSTB WITH (NOLOCK) WHERE TB002 IN ('106701') AND TB010=LA001 AND TB001>=@SDAY) AS '總銷售數量'
                            ,(SELECT ISNULL(SUM(TB019),0) FROM [TK].dbo.POSTB WITH (NOLOCK) WHERE TB002 IN ('106701') AND TB010=LA001 AND TB001>=@SDAY)/@TOTALDAYS AS '平均天銷售數量'
                            FROM [TK].dbo.INVLA WITH (NOLOCK) ,[TK].dbo.INVMB WITH (NOLOCK) 
                            WHERE LA009 IN ('21001')
                            AND LA001=MB001
                            AND LA001 LIKE '40%'
                            AND LA016 LIKE '2%'
                            AND MB002 NOT LIKE '%試吃%'
                            GROUP BY LA001,MB002,LA016
                            HAVING SUM(LA005*LA011)>0

                            ) AS TEMP
                            ) AS TEMP2
                            ORDER BY LA001


                            ", dateTimePicker2.Value.ToString("yyyyMMdd"), textBox2.Text.ToString());


            return SB;

        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();

        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }


        #endregion


    }
}
