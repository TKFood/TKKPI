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
                            SELECT MB001 AS '活動代號',MB003 AS '特價代號',MB004 AS '特價名稱',MB012 AS '特價起始日期',MB013 AS '特價截止日期'
                            FROM [TK].dbo.POSMB
                            WHERE MB001 LIKE '{0}%'
                            ORDER BY MB005,MB004

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
                            SELECT POSMB.MB003  AS '特價代號',POSMB.MB004  AS '特價名稱',TB002  AS '店代',MA002  AS '店名',TB013  AS '品號',INVMB.MB002  AS '品名',SUM(TB019) AS '數量',SUM(TB031) AS '未稅金額'
                            FROM [TK].dbo.POSTB,[TK].dbo.INVMB,[TK].dbo.WSCMA,[TK].dbo.POSMB
                            WHERE TB013=INVMB.MB001
                            AND MA001=TB002
                            AND TB036=POSMB.MB003
                            AND TB036 LIKE '%{0}%'
                            GROUP BY POSMB.MB003,POSMB.MB004,TB002,MA002,TB013,INVMB.MB002

                            ", TB036);


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
        #endregion
    }
}
