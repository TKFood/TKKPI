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
    public partial class frmREPROTPOS : Form
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


        public frmREPROTPOS()
        {
            InitializeComponent();

            dateTimePicker1.Value = DateTime.Now;
        }

        #region FUNCTION
        public void SETFASTREPORT(string YEARS)
        {
            StringBuilder SQL1 = new StringBuilder();    

            SQL1 = SETSQL(YEARS);
    

            Report report1 = new Report();

            report1.Load(@"REPORT\營銷-特價報表.frx");

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
    
          
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                           SELECT KIND,活動代號,特價代號,活動名稱,活動起始日期,活動截止日期
                            ,(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTB WHERE TB036=特價代號) AS '未稅金額'
                            ,(SELECT ISNULL(SUM(TB019),0) FROM [TK].dbo.POSTB WHERE TB036=特價代號) AS '數量'
                            FROM (
                            SELECT '特價' AS 'KIND', MB001 AS '活動代號',MB003 AS '特價代號',MB004 AS '活動名稱',MB012 AS '活動起始日期',MB013 AS '活動截止日期'
                            FROM [TK].dbo.POSMB
                            WHERE MB001 LIKE '{0}%'
                            UNION ALL
                            SELECT '組合品搭贈' AS 'KIND',MI001 AS '活動代號',MI003 AS '特價代號',MI004 AS '活動名稱',MI005 AS '活動起始日期',MI006 AS '活動截止日期'
                            FROM [TK].dbo.POSMI
                            WHERE MI001 LIKE '{0}%'
                            UNION ALL
                            SELECT '滿額折價' AS 'KIND',MM001 AS '活動代號',MM003 AS '特價代號',MM004 AS '活動名稱',MM005 AS '活動起始日期',MM006 AS '活動截止日期'
                            FROM [TK].dbo.POSMM
                            WHERE MM001 LIKE '{0}%'
                            UNION ALL
                            SELECT '配對搭贈' AS 'KIND',MO001 AS '活動代號',MO003 AS '特價代號',MO004 AS '活動名稱',MO005 AS '活動起始日期',MO006 AS '活動截止日期'
                            FROM [TK].dbo.POSMO
                            WHERE MO001 LIKE '{0}%'
                            ) AS TEMP



                            ", YEARS);


            return SB;

        }


        public void SETFASTREPORT2(string TB036)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2(TB036);


            Report report1 = new Report();

            report1.Load(@"REPORT\營銷-特價商品銷售.frx");

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

        public StringBuilder SETSQL2(string TB036)
        {


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                            SELECT (ISNULL(POSMB.MB003,'')+ISNULL(MI003,'')+ISNULL(MM003,'')+ISNULL(MO003,'') ) AS '特價代號',(ISNULL(POSMB.MB004,'')+ISNULL(MI004,'')+ISNULL(MM004,'')+ISNULL(MO004,''))  AS '特價名稱',TB002  AS '店代',MA002  AS '店名',TB010  AS '品號',INVMB.MB002  AS '品名',SUM(TB019) AS '數量',SUM(TB031) AS '未稅金額'
                            FROM [TK].dbo.INVMB,[TK].dbo.WSCMA,[TK].dbo.POSTB
                            LEFT JOIN [TK].dbo.POSMB ON MB003=TB036
                            LEFT JOIN [TK].dbo.POSMI ON MI003=TB036
                            LEFT JOIN [TK].dbo.POSMM ON MM003=TB036
                            LEFT JOIN [TK].dbo.POSMO ON MO003=TB036
                            WHERE TB010=INVMB.MB001
                            AND MA001=TB002
                            AND TB036 LIKE '%{0}%'
                            GROUP BY (ISNULL(POSMB.MB003,'')+ISNULL(MI003,'')+ISNULL(MM003,'')+ISNULL(MO003,'') ) ,(ISNULL(POSMB.MB004,'')+ISNULL(MI004,'')+ISNULL(MM004,'')+ISNULL(MO004,'')),TB002,MA002,TB010,INVMB.MB002

 

                            ", TB036);


            return SB;

        }

        public void SETFASTREPORT3(string SDATE,string EDATE,string TB036)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3(SDATE, EDATE, TB036);


            Report report1 = new Report();

            report1.Load(@"REPORT\營銷活動報表.frx");

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

            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL3(string SDATE, string EDATE, string TB036)
        {


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                            SELECT TB004 AS '營業日期',TB002 AS '店代',MA002 AS '店名',TB003 AS '機號',TB010 AS '品號',MB002 AS '品名',TB019 AS '數量',TB031 AS '未稅金額',TB044 AS '備註'
                            ,商品+組合+贈品加價購+配對搭贈 AS '活動代號'
                            ,商品名稱+組合名稱+贈品加價購名稱+配對搭贈名稱  AS '活動名稱'
                            FROM 
                            (
                            SELECT POSTB.*
                            ,ISNULL(RTRIM(LTRIM(MB003)),'') AS '商品',ISNULL(MB004,'') AS '商品名稱'
                            ,ISNULL(RTRIM(LTRIM(MI003)),'') AS '組合',ISNULL(MI004,'') AS '組合名稱'
                            ,ISNULL(RTRIM(LTRIM(MM003)),'') AS '贈品加價購',ISNULL(MM004,'') AS '贈品加價購名稱'
                            ,ISNULL(RTRIM(LTRIM(MO003)),'') AS '配對搭贈',ISNULL(MO004,'') AS '配對搭贈名稱'
                            FROM [TK].dbo.POSTB WITH (NOLOCK)
                            LEFT JOIN [TK].dbo.POSMB ON MB003=TB036
                            LEFT JOIN [TK].dbo.POSMI ON MI003=TB036
                            LEFT JOIN [TK].dbo.POSMM ON MM003=TB036
                            LEFT JOIN [TK].dbo.POSMO ON MO003=TB036

                            WHERE 1=1
                            AND ISNULL(TB044,'')<>''
                            ) AS TEMMP
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            LEFT JOIN [TK].dbo.WSCMA ON MA001=TB002
                            WHERE 1=1
                            AND ISNULL(商品+組合+贈品加價購+配對搭贈,'')<>''
                            AND TB001>='{0}' AND TB001<='{1}'
                            AND (ISNULL(商品+組合+贈品加價購+配對搭贈,'') LIKE '%{2}%' OR 商品名稱+組合名稱+贈品加價購名稱+配對搭贈名稱 LIKE '%{2}%')
                            ORDER BY TB002,TB004 
                            ", SDATE, EDATE, TB036);


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
            SETFASTREPORT2(textBox1.Text.ToString().Trim());
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3(dateTimePicker2.Value.ToString("yyyyMMdd"),dateTimePicker3.Value.ToString("yyyyMMdd"), textBox2.Text.ToString().Trim());
        }
        #endregion


    }
}
