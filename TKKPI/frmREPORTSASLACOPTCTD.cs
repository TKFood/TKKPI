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
    public partial class frmREPORTSASLACOPTCTD : Form
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

        public frmREPORTSASLACOPTCTD()
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
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKKPI].[dbo].[KINDSTORE] ORDER BY [ID] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "NAME";
            comboBox1.DisplayMember = "NAME";
            sqlConn.Close(); 

        }

        public void SETFASTREPORT()
        {
            //P1 去年上個月  @LASTYEARLASTMONTH
            //P2 今年上個月 @THISYEARLASTMONTH
            //P3 去年本月  @LASTYEARTHISMONTH
            //P4 今年本月 @THISYEARTHISMONTH
            //P5 去下下個月 @LASTYEARNEXTMONTH
            //P6 今年下個月 @THISYEARNEXTMONTH
            string P1 = null;
            string P2 = null;
            string P3 = null;
            string P4 = null;
            string P5 = null;
            string P6 = null;
            string KINDS = null;
            string LA007 = null;
            string TC004 = null;

            DateTime NOWdt = dateTimePicker1.Value;
            P1 = NOWdt.AddYears(-1).AddMonths(-1).ToString("yyyyMM");
            P2 = NOWdt.AddMonths(-1).ToString("yyyyMM");
            P3 = NOWdt.AddYears(-1).ToString("yyyyMM");
            P4 = NOWdt.ToString("yyyyMM");
            P5 = NOWdt.AddYears(-1).AddMonths(1).ToString("yyyyMM");
            P6 = NOWdt.AddMonths(1).ToString("yyyyMM");

            if(comboBox1.Text.Equals("門市"))
            {
                KINDS = "門市";
                LA007 = "1065%";
                TC004 = "91000006";

            }
            else if (comboBox1.Text.Equals("觀光工廠"))
            {
                KINDS = "觀光工廠";
                LA007 = "106701%";
                TC004 = "91000000"; 
            }

            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(KINDS, LA007, TC004,P1,P2,P3,P4,P5,P6);
            Report report1 = new Report();
            report1.Load(@"REPORT\門市營銷-銷貨+預計訂單備貨商品比較表.frx");

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

            report1.SetParameterValue("P1", P1);
            report1.SetParameterValue("P2", P2);
            report1.SetParameterValue("P3", P3);
            report1.SetParameterValue("P4", P4);
            report1.SetParameterValue("P5", P5);
            report1.SetParameterValue("P6", P6);

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string KINDS, string LA007, string TC004, string P1, string P2, string P3, string P4, string P5, string P6)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            DECLARE @KINDS NVARCHAR(12)
                            DECLARE @LA007 NVARCHAR(12)
                            DECLARE @TC004 NVARCHAR(12)
                            DECLARE @LASTYEARLASTMONTH NVARCHAR(12)
                            DECLARE @LASTYEARTHISMONTH NVARCHAR(12)
                            DECLARE @LASTYEARNEXTMONTH NVARCHAR(12)
                            DECLARE @THISYEARLASTMONTH NVARCHAR(12)
                            DECLARE @THISYEARTHISMONTH NVARCHAR(12)
                            DECLARE @THISYEARNEXTMONTH NVARCHAR(12)

                            SET @KINDS='{0}'
                            SET @LA007='{1}'
                            SET @TC004='{2}'

                            SET @LASTYEARLASTMONTH='{3}'
                            SET @THISYEARLASTMONTH='{4}'
                            SET @LASTYEARTHISMONTH='{5}'
                            SET @THISYEARTHISMONTH='{6}'
                            SET @LASTYEARNEXTMONTH='{7}'
                            SET @THISYEARNEXTMONTH='{8}'

                            SELECT @THISYEARTHISMONTH AS 'YM',@KINDS AS 'KINDS',LA005,MB002
                            ,(SELECT ISNULL(SUM(LA016),0) FROM [TK].dbo.SASLA LA WHERE LA.LA007 LIKE @LA007  AND SUBSTRING(CONVERT(nvarchar,LA.LA015,112),1,6)=@LASTYEARLASTMONTH AND LA.LA005=TEMP2.LA005) AS 'LASTMONTHACTNUMS'
                            ,(SELECT ISNULL(SUM(LA017),0) FROM [TK].dbo.SASLA LA WHERE LA.LA007 LIKE @LA007  AND SUBSTRING(CONVERT(nvarchar,LA.LA015,112),1,6)=@LASTYEARLASTMONTH AND LA.LA005=TEMP2.LA005) AS 'LASTMONTHACTMONEYS'
                            ,(SELECT ISNULL(SUM(TD008),0) FROM [TK].dbo.COPTC,[TK].dbo.COPTD WHERE TC001=TD001 AND TC002=TD002 AND TC001='A228' AND TC004=@TC004 AND SUBSTRING(TD013,1,6)=@THISYEARLASTMONTH AND TD004=LA005 ) AS 'LASTMONTHNUMS'
                            ,(SELECT ISNULL(SUM(TD012),0) FROM [TK].dbo.COPTC,[TK].dbo.COPTD WHERE TC001=TD001 AND TC002=TD002  AND TC001='A228' AND TC004=@TC004 AND SUBSTRING(TD013,1,6)=@THISYEARLASTMONTH AND TD004=LA005 ) AS 'LASTMONTHMONEYS'
                            ,(SELECT ISNULL(SUM(LA016),0) FROM [TK].dbo.SASLA LA WHERE LA.LA007 LIKE @LA007  AND SUBSTRING(CONVERT(nvarchar,LA.LA015,112),1,6)= @LASTYEARTHISMONTH AND LA.LA005=TEMP2.LA005) AS 'THISMONTHACTNUMS'
                            ,(SELECT ISNULL(SUM(LA017),0) FROM [TK].dbo.SASLA LA WHERE LA.LA007 LIKE @LA007  AND SUBSTRING(CONVERT(nvarchar,LA.LA015,112),1,6)= @LASTYEARTHISMONTH AND LA.LA005=TEMP2.LA005) AS 'THISMONTHACTMONEYS'
                            ,(SELECT ISNULL(SUM(TD008),0) FROM [TK].dbo.COPTC,[TK].dbo.COPTD WHERE TC001=TD001 AND TC002=TD002  AND TC001='A228' AND TC004=@TC004 AND SUBSTRING(TD013,1,6)=@THISYEARTHISMONTH AND TD004=LA005 ) AS 'THISMONTHNUMS'
                            ,(SELECT ISNULL(SUM(TD012),0) FROM [TK].dbo.COPTC,[TK].dbo.COPTD WHERE TC001=TD001 AND TC002=TD002  AND TC001='A228' AND TC004=@TC004 AND SUBSTRING(TD013,1,6)=@THISYEARTHISMONTH AND TD004=LA005 ) AS 'THISMONTHMONEYS'
                            ,(SELECT ISNULL(SUM(LA016),0) FROM [TK].dbo.SASLA LA WHERE LA.LA007 LIKE @LA007 AND SUBSTRING(CONVERT(nvarchar,LA.LA015,112),1,6)=@LASTYEARNEXTMONTH AND LA.LA005=TEMP2.LA005) AS 'NEXTMONTHACTNUMS'
                            ,(SELECT ISNULL(SUM(LA017),0) FROM [TK].dbo.SASLA LA WHERE LA.LA007 LIKE @LA007  AND SUBSTRING(CONVERT(nvarchar,LA.LA015,112),1,6)=@LASTYEARNEXTMONTH AND LA.LA005=TEMP2.LA005) AS 'NEXTMONTHACTMONEYS'
                            ,(SELECT ISNULL(SUM(TD008),0) FROM [TK].dbo.COPTC,[TK].dbo.COPTD WHERE TC001=TD001 AND TC002=TD002 AND TC001='A228' AND TC004=@TC004 AND SUBSTRING(TD013,1,6)=@THISYEARNEXTMONTH AND TD004=LA005 ) AS 'NEXTMONTHNUMS'
                            ,(SELECT ISNULL(SUM(TD012),0) FROM [TK].dbo.COPTC,[TK].dbo.COPTD WHERE TC001=TD001 AND TC002=TD002 AND TC001='A228' AND TC004=@TC004 AND SUBSTRING(TD013,1,6)=@THISYEARNEXTMONTH AND TD004=LA005 ) AS 'NEXTMONTHMONEYS'
                            FROM (
                            SELECT LA005
                            FROM (
                            --去年的銷售商品
                            SELECT LA005
                            FROM [TK].dbo.SASLA 
                            WHERE LA007 LIKE @LA007
                            AND LA005 LIKE '4%'
                            AND (SUBSTRING(CONVERT(nvarchar,LA015,112),1,6)=@LASTYEARLASTMONTH OR SUBSTRING(CONVERT(nvarchar,LA015,112),1,6)=@LASTYEARTHISMONTH  OR SUBSTRING(CONVERT(nvarchar,LA015,112),1,6)=@LASTYEARNEXTMONTH )
                            GROUP BY LA005
                            UNION ALL 
                            SELECT TD004
                            FROM [TK].dbo.COPTC,[TK].dbo.COPTD
                            WHERE TC001=TD001 
                            AND TC002=TD002
                            AND TC004=@TC004
                            AND (TD004 LIKE '4%' OR TD004 LIKE '5%')
                            AND TC001='A228'
                            AND (TD013 LIKE @THISYEARLASTMONTH+'%' OR TD013 LIKE @THISYEARTHISMONTH+'%'  OR TD013 LIKE @THISYEARNEXTMONTH+'%' )
                            GROUP BY TD004
                            ) AS TEMP
                            GROUP BY LA005
                            ) AS TEMP2
                            LEFT JOIN [TK].dbo.INVMB ON MB001=LA005
                            ORDER BY LA005
 

                            ", KINDS, LA007, TC004,P1,P2,P3,P4,P5,P6);

            return SB;

        }

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
