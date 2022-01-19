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
    public partial class frmREPORTAVGS : Form
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

        public frmREPORTAVGS()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SETFASTREPORT()
        {
           

            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\門市-平均交物筆數、平均客單價年報.frx");

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

            report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
      
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            string SDATES = null;
            string EDATES = null;
            SDATES = dateTimePicker1.Value.ToString("yyyy") + "0101";
            EDATES = dateTimePicker1.Value.ToString("yyyy") + "1231";

            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT 年月,門市代號,門市,成交筆數,平均客單價,查詢起日,查詢迄日,(門市來客數+觀光銷售來客數) AS 來客數,(CASE WHEN 門市來客數>0 THEN ROUND(成交筆數/門市來客數,4) ELSE 0 END) AS PCTS,ROUND(成交筆數*平均客單價,0) AS '銷售總金額'
                            FROM (
                            SELECT SUBSTRING(TT001,1,6) AS '年月',TT002 AS '門市代號',MA002 AS '門市',SUM(TT008) AS '成交筆數',SUM(TT011)/SUM(TT008) AS '平均客單價'
                            ,(SELECT TOP 1 TT001 FROM [TK].dbo.POSTT WHERE TT001>='{0}' AND TT001<='{1}' ORDER BY TT001)  AS '查詢起日'
                            ,(SELECT TOP 1 TT001 FROM [TK].dbo.POSTT WHERE TT001>='{0}' AND TT001<='{1}' ORDER BY TT001 DESC) AS '查詢迄日'
                            ,(SELECT ISNULL(SUM(Fin_data+Fout_data)/2,0) FROM [TKMK].[dbo].[View_t_visitors] WHERE TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704') AND [View_t_visitors].TT002=POSTT.TT002 AND CONVERT(NVARCHAR,Fdate1,112) LIKE SUBSTRING(TT001,1,6)+'%') AS '門市來客數'
                            ,(SELECT ISNULL(SUM(Fin_data),0) FROM [TKMK].[dbo].[View_t_visitors] WHERE TT002 IN ('106701') AND [View_t_visitors].TT002=POSTT.TT002 AND CONVERT(NVARCHAR,Fdate1,112) LIKE SUBSTRING(TT001,1,6)+'%') AS '觀光銷售來客數'
                            FROM [TK].dbo.POSTT,[TK].dbo.WSCMA
                            WHERE TT002=MA001
                            AND TT002 IN (SELECT  [TT002]  FROM [TKKPI].[dbo].[SALESTORES])
                            AND TT001 LIKE '{0}%'
                            GROUP BY SUBSTRING(TT001,1,6),TT002,MA002
                            ) AS TEMP
                            ORDER BY 年月,門市代號
                            ", dateTimePicker1.Value.ToString("yyyy"), SDATES, EDATES);

            return SB;

        }

        public void SETFASTREPORT2()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2();
            Report report1 = new Report();
            report1.Load(@"REPORT\門市-平均交物筆數、平均客單價明細.frx");

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

            report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL2()
        {


            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                           SELECT TT002 AS '門市代號',MA002 AS '門市',SUM(TT008) AS '成交筆數',SUM(TT011)/SUM(TT008) AS '平均客單價'
                            ,(SELECT TOP 1 TT001 FROM [TK].dbo.POSTT WHERE TT001>='{0}' AND TT001<='{1}' ORDER BY TT001)  AS '查詢起日'
                            ,(SELECT TOP 1 TT001 FROM [TK].dbo.POSTT WHERE TT001>='{0}' AND TT001<='{1}' ORDER BY TT001 DESC) AS '查詢迄日'
                            FROM [TK].dbo.POSTT,[TK].dbo.WSCMA
                            WHERE TT002=MA001
                            AND TT002 IN (SELECT  [TT002]  FROM [TKKPI].[dbo].[SALESTORES])
                            AND TT001>='{0}' AND TT001<='{1}'
                            GROUP BY TT002,MA002
 
                            ", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));

            return SB;

        }

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }
        #endregion


    }
}
