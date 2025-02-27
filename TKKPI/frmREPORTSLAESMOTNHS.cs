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
    public partial class frmREPORTSLAESMOTNHS : Form
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


        public frmREPORTSLAESMOTNHS()
        {
            InitializeComponent();
        }
        private void frmREPORTSLAESMOTNHS_Load(object sender, EventArgs e)
        {
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
            Sequel.AppendFormat(@"
                                SELECT 
                                 [ID]
                                ,[KINDS]
                                ,[NAMES]
                                ,[VALUE]
                                FROM [TKKPI].[dbo].[TBPARA]
                                WHERE [KINDS]='frmREPORTSLAESMOTNHS'
                                ORDER BY  [ID]
                                ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);

            CBX.DataSource = dt.DefaultView;
            CBX.ValueMember = "NAMES";
            CBX.DisplayMember = "NAMES";
            sqlConn.Close();

            CBX.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }

        public void SETFASTREPORT(string REPORT,string YYMM)
        {
            StringBuilder SQL1 = new StringBuilder();
            Report report1 = new Report();

            report1.Load(@"REPORT\業務-月報-客戶.frx");
            report1.SetParameterValue("P1", YYMM);

            SQL1 = SETSQL1(YYMM);

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
        public StringBuilder SETSQL1(string YYMM)
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"    
                            SELECT *
                            FROM 
                            (
                            SELECT 
                            '1銷貨' KINDS,MA002,SUM(TH037) SUMTH037,SUM(TH038) SUMTH038,SUM(TH037+TH038) SUMMONEYS
                            FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.COPMA
                            WHERE TG001=TH001 AND TG002=TH002
                            AND TG004=MA001
                            AND TG023='Y'
                            AND TG006 IN (
	                            SELECT [MV001]      
	                            FROM [TK].[dbo].[Z_SALES_DAILY_REPORTS]
	                            WHERE [NATIONS]='國內'
                            )
                            AND TG003 LIKE '{0}%'
                            GROUP BY MA002
                            UNION ALL
                            SELECT 
                            '2銷退' KINDS,MA002,TJ033*-1,TJ034*-1,(TJ033+TJ034)*-1
                            FROM [TK].dbo.COPTI,[TK].dbo.COPTJ,[TK].dbo.COPMA
                            WHERE TI001=TJ001 AND TI002=TJ002
                            AND TI004=MA001
                            AND TI019='Y'
                            AND TI006     IN (
	                            SELECT [MV001]      
	                            FROM [TK].[dbo].[Z_SALES_DAILY_REPORTS]
	                            WHERE [NATIONS]='國內'
                            )
                            AND TI003 LIKE '{0}%'
                            ) AS TEMP
                         
                             ", YYMM);

            talbename = "TEMPds1";

            return SB;

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            string REPOSRT = comboBox1.Text.ToString();
            string YYMM = dateTimePicker1.Value.ToString("yyyyMM");
            SETFASTREPORT(REPOSRT, YYMM);
        }
        #endregion

       
    }

}
