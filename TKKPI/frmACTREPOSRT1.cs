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
using System.Globalization;

namespace TKKPI
{
    public partial class frmACTREPOSRT1 : Form
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
 

        public frmACTREPOSRT1()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT(string YEARS)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(YEARS);
            Report report1 = new Report();

            report1.Load(@"REPORT\產品貢獻度.frx");

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

        public StringBuilder SETSQL(string YEARS)
        {
            string FirstDay = YEARS + "0101";
            string LastDay = YEARS + "1231";

            DateTime firstDate = DateTime.ParseExact(FirstDay, "yyyyMMdd", CultureInfo.InvariantCulture);
            DateTime lastDate = DateTime.ParseExact(LastDay, "yyyyMMdd", CultureInfo.InvariantCulture);

            TimeSpan diff = lastDate - firstDate;
            int DAYS = (int)diff.TotalDays;
            DAYS = DAYS + 1;

            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                                SELECT *
                                ,RANK() OVER (ORDER BY 本期銷貨額 DESC) AS '銷貨排名'
                                ,本期銷貨額/SUM(本期銷貨額) OVER() AS '銷售比重'
                                ,(本期銷貨額-成本總額)/SUM((本期銷貨額-成本總額)) OVER() AS '毛利額比重'
                                ,RANK() OVER (ORDER BY (CASE WHEN  (本期銷貨額-成本總額)>0 AND 本期銷貨額>0 THEN (本期銷貨額-成本總額)/本期銷貨額 ELSE 0 END) DESC) AS '毛利率排名'
                                ,(CASE WHEN 本期銷貨額>0 AND 成本總額>0 THEN 本期銷貨額/成本總額 ELSE 0 END )AS '產品週轉率'
                                ,RANK() OVER (ORDER BY (CASE WHEN 本期銷貨額>0 AND 成本總額>0 THEN (本期銷貨額/成本總額) ELSE 0 END ) DESC) AS '週轉排名'
                                ,((SELECT SUM([SUMLA013])  FROM [TK].[dbo].[ZINVLASUM] WHERE [ZINVLASUM].MB001=TEMP.品號 AND DATES>='{0}' AND DATES<='{1}')/{2}) AS '平均存貨額'
                                ,((本期銷貨額-成本總額)/SUM((本期銷貨額-成本總額)) OVER()*(CASE WHEN 本期銷貨額>0 AND 成本總額>0 THEN 本期銷貨額/成本總額 ELSE 0 END )) AS '交叉比率'
                                FROM 
                                (
                                SELECT  MB001 AS '品號', MB002 AS '商品/類別',MB003 AS '規格',MB004 AS '單位',SUM(LA017) AS '本期銷貨額'
                                ,SUM(LA024) AS '成本總額'
                                ,SUM(LA017-LA024) AS '毛利總額'
                                ,(CASE WHEN SUM(LA017-LA024)>0 AND SUM(LA017)>0 THEN (SUM(LA017-LA024)/SUM(LA017)) ELSE 0 END) AS '毛利率'
                                FROM [TK].dbo.INVMB,[TK].dbo.SASLA
                                WHERE MB001=LA005
                                AND (MB001 LIKE '4%' OR MB001 LIKE '5%')
                                AND MB001 NOT LIKE '49%'
                                AND MB001 NOT LIKE '59%'
                                AND CONVERT(NVARCHAR, LA015, 112) >= '{0}' AND  CONVERT(NVARCHAR, LA015, 112) <= '{1}'
                                GROUP BY MB001, MB002,MB003,MB004
                                ) AS TEMP
                                ORDER BY 本期銷貨額 DESC
                            ", FirstDay, LastDay,DAYS);


            return SB;

        }

        public void SETFASTREPORT2(string SYEARSMONTHS, string EYEARSMONTHS)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2(SYEARSMONTHS, EYEARSMONTHS);
            Report report1 = new Report();

            report1.Load(@"REPORT\產品貢獻度.frx");

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

        public StringBuilder SETSQL2(string SYEARSMONTHS,string EYEARSMONTHS)
        {
            string FirstDay = SYEARSMONTHS + "01";
            string LastDay = EYEARSMONTHS + "31";

            DateTime firstDate = DateTime.ParseExact(FirstDay, "yyyyMMdd", CultureInfo.InvariantCulture);
            DateTime lastDate = DateTime.ParseExact(LastDay, "yyyyMMdd", CultureInfo.InvariantCulture);

            TimeSpan diff = lastDate - firstDate;
            int DAYS = (int)diff.TotalDays;
            DAYS = DAYS + 1;

            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                                SELECT *
                                ,RANK() OVER (ORDER BY 本期銷貨額 DESC) AS '銷貨排名'
                                ,本期銷貨額/SUM(本期銷貨額) OVER() AS '銷售比重'
                                ,(本期銷貨額-成本總額)/SUM((本期銷貨額-成本總額)) OVER() AS '毛利額比重'
                                ,RANK() OVER (ORDER BY (CASE WHEN  (本期銷貨額-成本總額)>0 AND 本期銷貨額>0 THEN (本期銷貨額-成本總額)/本期銷貨額 ELSE 0 END) DESC) AS '毛利率排名'
                                ,(CASE WHEN 本期銷貨額>0 AND 成本總額>0 THEN 本期銷貨額/成本總額 ELSE 0 END )AS '產品週轉率'
                                ,RANK() OVER (ORDER BY (CASE WHEN 本期銷貨額>0 AND 成本總額>0 THEN (本期銷貨額/成本總額) ELSE 0 END ) DESC) AS '週轉排名'
                                ,((SELECT SUM([SUMLA013])  FROM [TK].[dbo].[ZINVLASUM] WHERE [ZINVLASUM].MB001=TEMP.品號 AND DATES>='{0}' AND DATES<='{1}')/{2}) AS '平均存貨額'
                                ,((本期銷貨額-成本總額)/SUM((本期銷貨額-成本總額)) OVER()*(CASE WHEN 本期銷貨額>0 AND 成本總額>0 THEN 本期銷貨額/成本總額 ELSE 0 END )) AS '交叉比率'
                                FROM 
                                (
                                SELECT  MB001 AS '品號', MB002 AS '商品/類別',MB003 AS '規格',MB004 AS '單位',SUM(LA017) AS '本期銷貨額'
                                ,SUM(LA024) AS '成本總額'
                                ,SUM(LA017-LA024) AS '毛利總額'
                                ,(CASE WHEN SUM(LA017-LA024)>0 AND SUM(LA017)>0 THEN (SUM(LA017-LA024)/SUM(LA017)) ELSE 0 END) AS '毛利率'
                                FROM [TK].dbo.INVMB,[TK].dbo.SASLA
                                WHERE MB001=LA005
                                AND (MB001 LIKE '4%' OR MB001 LIKE '5%')
                                AND MB001 NOT LIKE '49%'
                                AND MB001 NOT LIKE '59%'
                                AND CONVERT(NVARCHAR, LA015, 112) >= '{0}' AND  CONVERT(NVARCHAR, LA015, 112) <= '{1}'
                                GROUP BY MB001, MB002,MB003,MB004
                                ) AS TEMP
                                ORDER BY 本期銷貨額 DESC
                            ", FirstDay, LastDay, DAYS);


            return SB;

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyy"));
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(dateTimePicker2.Value.ToString("yyyyMM"), dateTimePicker3.Value.ToString("yyyyMM"));
        }
        #endregion


    }
}
