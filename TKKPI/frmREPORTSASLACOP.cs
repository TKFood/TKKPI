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
    public partial class frmREPORTSASLACOP : Form
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


        public frmREPORTSASLACOP()
        {
            InitializeComponent();

            SETDATES();
        }

        #region FUNCTION
        public void SETDATES()
        {
            dateTimePicker1.Value = DateTime.Now.AddMonths(-1);
        }
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();

            report1.Load(@"REPORT\上月的實收客戶淨額+業務員.frx");

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
            DateTime dt = dateTimePicker1.Value;
            DateTime FirstDay = new DateTime(dt.Year, dt.Month, 1);
            DateTime LastDay = new DateTime(dt.AddMonths(1).Year, dt.AddMonths(1).Month, 1).AddDays(-1);
                        
            string DEPNO = textBox1.Text.Trim();

            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                            SELECT LA008,MV002,LA006,MA002
                            ,(SUM(LA017)-SUM(LA020)-SUM(LA022)) 銷貨淨額
                            ,SUM(LA024) 銷貨成本
                            ,(SUM(LA017)-SUM(LA020)-SUM(LA022)-SUM(LA023)-SUM(LA024)) 銷貨毛利

                            ,SUM(LA017) 銷貨金額,SUM(LA020) 銷退金額,SUM(LA022) 折讓金額,SUM(LA023) 壞帳金額
                            FROM [TK].dbo.SASLA,[TK].dbo.COPMA,[TK].dbo.CMSMV
                            WHERE LA006=MA001
                            AND LA008=MV001
                            AND LA015>='{0}' AND LA015<='{1}'
                            AND LA007 LIKE '{2}%'
                            GROUP BY LA008,MV002,LA006,MA002
                            ORDER BY SUM(LA017) DESC
                            ", FirstDay.ToString("yyyyMMdd"), LastDay.ToString("yyyyMMdd"), DEPNO);


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
