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
    public partial class frmREPORTSALESCOP : Form
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

        public frmREPORTSALESCOP()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();

            report1.Load(@"REPORT\本月銷售+訂單-同期銷售.frx");

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

            string THISYM = "";
            string LASTYM = "";

            DateTime dt = new DateTime();
            THISYM = dateTimePicker1.Value.ToString("yyyyMM");

            dt = dateTimePicker1.Value;
            dt=dt.AddYears(-1);
            LASTYM= dt.ToString("yyyyMM");

            SB.AppendFormat(@"                              
                                SELECT '{0}' AS '年月',TG004 AS '客戶代號',MA002 AS '客戶',LASTTH037 AS '去年同月銷售金額',THISTH037 AS '本月銷售金額',THISTD012 AS '本月訂單金額',((THISTH037+THISTD012)-LASTTH037)  AS '差異金額(本月銷售金額+本月訂單金額-去年同月銷售金額)'
                                FROM (
                                SELECT TG004,MA002
                                ,(SELECT ISNULL(SUM(TH.TH037),0) FROM [TK].dbo.COPTG TG,[TK].dbo.COPTH TH WHERE TG.TG001=TH.TH001 AND TG.TG002=TH.TH002 AND TG.TG023='Y' AND TG.TG002 LIKE '{1}%' AND TG.TG004=TEMP.TG004) AS 'LASTTH037'
                                ,(SELECT ISNULL(SUM(TH.TH037),0) FROM [TK].dbo.COPTG TG,[TK].dbo.COPTH TH WHERE TG.TG001=TH.TH001 AND TG.TG002=TH.TH002 AND TG.TG023='Y' AND TG.TG002 LIKE '{0}%' AND TG.TG004=TEMP.TG004) AS 'THISTH037'
                                ,(SELECT ISNULL(SUM((TD.TD008-TD.TD009)*TD.TD011),0) FROM [TK].dbo.COPTC TC,[TK].dbo.COPTD TD WHERE TC.TC001=TD.TD001 AND TC.TC002=TD.TD002 AND TC.TC027='Y' AND TD.TD013 LIKE '{0}%' AND TC.TC004=TEMP.TG004) AS 'THISTD012'
                                FROM (
                                SELECT TG004
                                FROM [TK].dbo.COPTG,[TK].dbo.COPTH
                                WHERE TG001=TH001 AND TG002=TH002
                                AND TG002 LIKE '{1}%'
                                AND TG023='Y'
                                GROUP BY TG004
                                HAVING SUM(TH037)>0
                                UNION ALL
                                SELECT TG004
                                FROM [TK].dbo.COPTG,[TK].dbo.COPTH
                                WHERE TG001=TH001 AND TG002=TH002
                                AND TG002 LIKE '{0}%'
                                AND TG023='Y'
                                GROUP BY TG004
                                HAVING SUM(TH037)>0
                                ) AS TEMP
                                LEFT JOIN [TK].dbo.COPMA ON COPMA.MA001=TEMP.TG004
                                WHERE TEMP.TG004 NOT LIKE '1%'
                                AND TEMP.TG004 NOT LIKE '5%'
                                AND TEMP.TG004 NOT LIKE '7%'
                                GROUP BY TG004,MA002
                                ) AS TEMP2
                                ORDER BY ((THISTH037+THISTD012)-LASTTH037) DESC
  
                            ", THISYM,LASTYM);


            return SB;

        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
