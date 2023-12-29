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
    public partial class frmREPORTSALESMB001 : Form
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

        public frmREPORTSALESMB001()
        {
            InitializeComponent();

            comboBox1load();

        }

        #region FUNCTION
        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [COMMENTS]  FROM [TK].[dbo].[Z_TB_SALESMB001] GROUP BY [COMMENTS]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("COMMENTS", typeof(string));  
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "COMMENTS";
            comboBox1.DisplayMember = "COMMENTS";
            sqlConn.Close();

        }

        public void SETFASTREPORT(string SDAYS, string EDAYS,string COMMENTS)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDAYS, EDAYS, COMMENTS);


            Report report1 = new Report();

            report1.Load(@"REPORT\禮盒銷售.frx");

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

            report1.Preview = previewControl4;
            report1.Show();
        }

        public StringBuilder SETSQL(string SDAYS,string EDAYS,string COMMENTS)
        {


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                            SELECT MB001 AS '品號',MB002 AS '品名'
                            ,ISNULL((SELECT SUM(TB019) FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB002 LIKE '1065%' AND TB001>='{0}' AND TB001<='{1}'),0) AS '門市'
                            ,ISNULL((SELECT SUM(TB019) FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB002 LIKE '1067%' AND TB001>='{0}' AND TB001<='{1}'),0) AS '觀光'
                            ,ISNULL((SELECT SUM(TH008+TH024) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TH004=MB001 AND TG023='Y' AND TG006 IN (SELECT [MV001] FROM [TK].[dbo].[Z_TB_SALESMB001_SETSALES] WHERE [COMMENTS] IN ('電商')) AND TG003>='{0}' AND TG003<='{1}'),0) AS '電商'
                            ,ISNULL((SELECT SUM(TH008+TH024) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TH004=MB001 AND TG023='Y' AND TG006 IN (SELECT [MV001] FROM [TK].[dbo].[Z_TB_SALESMB001_SETSALES] WHERE [COMMENTS] IN ('張釋予')) AND TG003>='{0}' AND TG003<='{1}'),0) AS '張協理'
                            ,ISNULL((SELECT SUM(TH008+TH024) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TH004=MB001 AND TG023='Y' AND TG006 NOT IN (SELECT [MV001] FROM [TK].[dbo].[Z_TB_SALESMB001_SETSALES] WHERE [COMMENTS] IN ('張釋予','電商')) AND TG003>='{0}' AND TG003<='{1}'),0) AS '業務'
                            ,ISNULL((SELECT SUM(LA011*LA005) FROM [TK].dbo.INVLA WHERE LA001=MB001),0) AS '目前庫存'
                            FROM [TK].dbo.INVMB
                            WHERE MB001 IN (SELECT MB001 FROM [TK].[dbo].[Z_TB_SALESMB001] WHERE COMMENTS='{2}' )



                            ", SDAYS, EDAYS, COMMENTS);


            return SB;

        }

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"), comboBox1.Text.ToString());
        }

        #endregion
    }
}
