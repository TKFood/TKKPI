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
    public partial class frmREPORTSLASTTHIS : Form
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

        public frmREPORTSLASTTHIS()
        {
            InitializeComponent();
            
            comboBox1load();
        }

        #region FUNCTION
        public void comboBox1load()
        {
            ComboBox CBX = comboBox1;
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[KINDS],[NAMES],[VALUE] FROM [TKKPI].[dbo].[TBPARA] WHERE [KINDS]='frmREPORTSLASTTHIS' ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);

            CBX.DataSource = dt.DefaultView;
            CBX.ValueMember = "NAMES";
            CBX.DisplayMember = "NAMES";
            sqlConn.Close();

            CBX.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }

        public void SETFASTREPORT(string REPORTS,DateTime YEARMONTH)
        {
            StringBuilder SQL1 = new StringBuilder();

            DateTime DTTHISYEARSMONTH = YEARMONTH;
            DateTime DTLASTYEARSMONTH = YEARMONTH.AddMonths(-12);
            string THISYEARSMONTH = DTTHISYEARSMONTH.ToString("yyyyMM");
            string LASTYEARSMONTH= DTLASTYEARSMONTH.ToString("yyyyMM");


            Report report1 = new Report();

            if (REPORTS.Equals("國內差異分析"))
            {
                report1.Load(@"REPORT\國內差異分析.frx");
                report1.SetParameterValue("P1", LASTYEARSMONTH);
                report1.SetParameterValue("P2", THISYEARSMONTH);
                SQL1 = SETSQL1(THISYEARSMONTH, LASTYEARSMONTH);
            }
            else if (REPORTS.Equals("國外差異分析"))
            {
                report1.Load(@"REPORT\國外差異分析.frx");
                report1.SetParameterValue("P1", LASTYEARSMONTH);
                report1.SetParameterValue("P2", THISYEARSMONTH);
                SQL1 = SETSQL2(THISYEARSMONTH, LASTYEARSMONTH);
            }


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
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1(string THISYEARSMONTH, string LASTYEARSMONTH)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"  
                            SELECT LA006 AS '客代',MA002 AS '客戶'
                            ,(SELECT ISNULL(SUM(LA017),0) FROM  [TK].dbo.SASLA WHERE SASLA.LA006=TEMP.LA006 AND  CONVERT(NVARCHAR,LA015,112) LIKE '{1}%') AS 'LASTYEARMONTHMONEY'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND CONVERT(NVARCHAR,TG003,112) LIKE '{0}%' AND (TG004 LIKE '2%' OR TG004 LIKE 'A%') AND TG023='Y' AND TG004=TEMP.LA006) AS 'THISYEARMONTHMONEY'
                            ,(SELECT TOP 1 TG003 FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND CONVERT(NVARCHAR,TG003,112) LIKE '{0}%' AND (TG004 LIKE '2%' OR TG004 LIKE 'A%') AND TG023='Y' ORDER BY TG003 DESC)  AS 'EDAYS'
                            FROM 
                            (
                            SELECT LA006
                            FROM [TK].dbo.SASLA
                            LEFT JOIN [TK].dbo.COPMA ON MA001=LA006
                            WHERE CONVERT(NVARCHAR,LA015,112) LIKE '{1}%'
                            AND (LA006 LIKE '2%' OR LA006 LIKE 'A%')
                            GROUP BY LA006
                            UNION ALL
                            SELECT TG004
                            FROM [TK].dbo.COPTG
                            LEFT JOIN [TK].dbo.COPMA ON TG004=MA001
                            ,[TK].dbo.COPTH
                            WHERE TG001=TH001 AND TG002=TH002
                            AND CONVERT(NVARCHAR,TG003,112) LIKE '{0}%'
                            AND (TG004 LIKE '2%' OR TG004 LIKE 'A%')
                            AND TG023='Y'
                            GROUP BY TG004
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.COPMA ON MA001=LA006
                            GROUP BY LA006,MA002
                            ORDER BY LA006,MA002
                         
                             ", THISYEARSMONTH, LASTYEARSMONTH);

            talbename = "TEMPds1";

            return SB;

        }

        public StringBuilder SETSQL2(string THISYEARSMONTH, string LASTYEARSMONTH)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"  
                            SELECT LA006 AS '客代',MA002 AS '客戶'
                            ,(SELECT ISNULL(SUM(LA017),0) FROM  [TK].dbo.SASLA WHERE SASLA.LA006=TEMP.LA006 AND  CONVERT(NVARCHAR,LA015,112) LIKE '{1}%') AS 'LASTYEARMONTHMONEY'
                            ,(SELECT ISNULL(SUM(TH037),0) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND CONVERT(NVARCHAR,TG003,112) LIKE '{0}%' AND (TG004 LIKE '3%' OR TG004 LIKE 'B%') AND TG023='Y' AND TG004=TEMP.LA006) AS 'THISYEARMONTHMONEY'
                            ,(SELECT TOP 1 TG003 FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND CONVERT(NVARCHAR,TG003,112) LIKE '{0}%' AND (TG004 LIKE '3%' OR TG004 LIKE 'B%') AND TG023='Y' ORDER BY TG003 DESC)  AS 'EDAYS'
                            FROM 
                            (
                            SELECT LA006
                            FROM [TK].dbo.SASLA
                            LEFT JOIN [TK].dbo.COPMA ON MA001=LA006
                            WHERE CONVERT(NVARCHAR,LA015,112) LIKE '{1}%'
                            AND (LA006 LIKE '3%' OR LA006 LIKE 'B%')
                            GROUP BY LA006
                            UNION ALL
                            SELECT TG004
                            FROM [TK].dbo.COPTG
                            LEFT JOIN [TK].dbo.COPMA ON TG004=MA001
                            ,[TK].dbo.COPTH
                            WHERE TG001=TH001 AND TG002=TH002
                            AND CONVERT(NVARCHAR,TG003,112) LIKE '{0}%'
                            AND (TG004 LIKE '32%' OR TG004 LIKE 'B%')
                            AND TG023='Y'
                            GROUP BY TG004
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.COPMA ON MA001=LA006
                            GROUP BY LA006,MA002
                            ORDER BY LA006,MA002
                         
                             ", THISYEARSMONTH, LASTYEARSMONTH);

            talbename = "TEMPds1";

            return SB;

        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(comboBox1.Text.ToString(),dateTimePicker1.Value);
        }
        #endregion

    }
}
