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
    public partial class frmREPORTSOTREMATRIX : Form
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
      

        public frmREPORTSOTREMATRIX()
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
            Sequel.AppendFormat(@"SELECT [MA001],[MA002] FROM [TKKPI].[dbo].[KIINDTBDEPSTORES] ORDER BY [MA001]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MA001", typeof(string));
            dt.Columns.Add("MA002", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MA001";
            comboBox1.DisplayMember = "MA002";
            sqlConn.Close();

        }
        public void SETFASTREPORT(string LA015, string LA007)
        {
        

            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(LA015, LA007);
            Report report1 = new Report();
            report1.Load(@"REPORT\門市-各月銷售商品資料.frx");

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

            //report1.SetParameterValue("P1", P1);
          

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string LA015, string LA007)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT SUBSTRING (CONVERT(NVARCHAR,LA015,112),1,6) LA015,LA006,MA002,LA005,MB002,LA007,SUM(LA016) LA016,SUM(LA017) LA017
                            FROM [TK].dbo.SASLA,[TK].dbo.COPMA,[TK].dbo.INVMB
                            WHERE  LA006=MA001
                            AND LA005=MB001
                            AND (LA005 LIKE '4%' OR LA005 LIKE '5%' )
                            AND MB002 NOT LIKE '%試吃%'
                            AND CONVERT(NVARCHAR,LA015,112) LIKE '{0}%'
                            AND MA001='{1}'
                            GROUP BY SUBSTRING (CONVERT(NVARCHAR,LA015,112),1,6) ,LA006,MA002,LA005,MB002,LA007
                            ORDER BY SUBSTRING (CONVERT(NVARCHAR,LA015,112),1,6) ,LA006,MA002,LA005,MB002,LA007

 

                            ", LA015, LA007);

            return SB;

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyy"),comboBox1.SelectedValue.ToString());
        }
        #endregion
    }
}
