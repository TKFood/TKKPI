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
    public partial class frmREPORTSNEWS : Form
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

        public frmREPORTSNEWS()
        {
            InitializeComponent();

            SET();
        }

        #region FUNCTION
        public void SET()
        {
            // 本月的第一天
            DateTime firstDayOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            dateTimePicker1.Value = firstDayOfMonth;

            // 本月的最後一天
            DateTime lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);
            dateTimePicker2.Value = lastDayOfMonth;
        }

        public void SETFASTREPORT(string SDAYS,string EDAYS)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDAYS,EDAYS);
            Report report1 = new Report();

            report1.Load(@"REPORT\新品銷售資料.frx");

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

        public StringBuilder SETSQL(string SDAYS, string EDAYS)
        {     
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"                               
                            SELECT  
                            MB001 AS '品號'
                            ,MB002 AS '品名'
                            ,MB003 AS '規格'
                            ,MB004 AS '單位'
                            ,CREATE_DATE AS '新品建立日期'
                            ,TOPTG003 AS '第1天業務銷貨日'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTH008)), 1), '.00', '') AS '累計-業務銷貨數量'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTH037)), 1), '.00', '') AS '累計-業務銷貨金額'
                            ,TOPTI003 AS '第1天業務銷退日'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTJ007)), 1), '.00', '') AS '累計-業務銷退數量'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTJ033)), 1), '.00', '') AS '累計-業務銷退金額'
                            ,TOPTB001 AS '第1天POS銷售日'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTB019)), 1), '.00', '') AS '累計-POS銷售數量'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,SUMTB031)), 1), '.00', '') AS '累計-POS銷售金額'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(DECIMAL(16,4),PERCOSTS)), 1), '.00', '') AS '平均單位成本'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH008-SUMTJ007+SUMTB019))), 1), '.00', '')  AS '累計-總銷售數量'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH037-SUMTJ033+SUMTB031))), 1), '.00', '')  AS '累計-總銷售未稅金額'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019)))), 1), '.00', '')  AS '累計-總成本'
                            ,REPLACE(CONVERT(VARCHAR(20), CONVERT(MONEY,CONVERT(INT,(SUMTH037-SUMTJ033+SUMTB031-(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019))))), 1), '.00', '')  AS '累計-總毛利'
                            ,CONVERT(NVARCHAR,CONVERT(DECIMAL(16,2),(CASE WHEN (SUMTH037-SUMTJ033+SUMTB031-(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019)))<>0 AND (SUMTH037-SUMTJ007+SUMTB031)<>0  THEN (SUMTH037-SUMTJ033+SUMTB031-(PERCOSTS*(SUMTH008-SUMTJ007+SUMTB019)))/(SUMTH037+SUMTB031) ELSE  0 END )*100))+'%'  AS '累計-毛利率'
                            FROM 
                            (
                            SELECT *
                            ,ISNULL(
                            (SELECT CASE WHEN SUM(LA024)<>0 AND SUM(LA016)<>0 THEN SUM(LA024)/SUM(LA016) ELSE 0 END
                            FROM [TK].dbo.SASLA
                            WHERE LA005=MB001
                            AND CONVERT(NVARCHAR,LA015,112)>='{0}'
                            AND CONVERT(NVARCHAR,LA015,112)<='{1}')
                            ,0) AS PERCOSTS
                            FROM (
                            SELECT '{0}' SDATES,'{1}' AS EDATES,MB001,MB002,MB003,MB004,CREATE_DATE
                            ,ISNULL((SELECT TOP 1 ISNULL(TG003,'') FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}' AND TH004=MB001 ORDER BY TG003 ),'') AS TOPTG003
                            ,ISNULL((SELECT SUM((CASE WHEN TH009=MD002 THEN ((TH008+TH024)*MD004/MD003) ELSE (TH008+TH024) END)) FROM [TK].dbo.COPTG,[TK].dbo.COPTH LEFT JOIN [TK].dbo.INVMD ON MD001=TH004 WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}' AND TH004=MB001),0) AS SUMTH008
                            ,ISNULL((SELECT SUM(TH037) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002 AND TG023='Y' AND TG003>='{0}' AND TH004=MB001),0) AS SUMTH037

                            ,ISNULL((SELECT TOP 1 ISNULL(TI003,'') FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}' AND TJ004=MB001 ORDER BY TI003 ),'') AS TOPTI003
                            ,ISNULL((SELECT SUM((CASE WHEN TJ008=MD002 THEN (TJ007*MD004/MD003) ELSE TJ007 END)) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ LEFT JOIN [TK].dbo.INVMD ON MD001=TJ004 WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}' AND TJ004=MB001),0) AS SUMTJ007
                            ,ISNULL((SELECT SUM(TJ033) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002 AND TI019='Y' AND TI003>='{0}' AND TJ004=MB001),0) AS SUMTJ033

                            ,ISNULL((SELECT TOP 1 ISNULL(TB001,'') FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}' ORDER BY TB001),'') AS TOPTB001
                            ,ISNULL((SELECT SUM(TB019) FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}'),0) AS SUMTB019
                            ,ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB010=MB001 AND TB001>='{0}'),0) AS SUMTB031
                            FROM [TK].dbo.INVMB
                            WHERE 1=1
                            AND MB001 LIKE '4%'
                            AND MB002 NOT LIKE '%試吃%'
                            AND MB002 NOT LIKE '%空%'
                            AND ISNULL(MB002,'')<>''
                            AND CREATE_DATE>='{0}'
                            ) AS TEMP
                            ) AS TEMP2
                            ORDER BY (SUMTH037-SUMTJ033+SUMTB031) DESC,新品建立日期
                            ", SDAYS, EDAYS);


            return SB;

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        #endregion
    }
}
