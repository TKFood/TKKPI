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
    public partial class frmREPORTSALSA : Form
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

        public frmREPORTSALSA()
        {
            InitializeComponent();

            comboBox4load();
        }


        #region FUNCTION
        public void comboBox4load()
        {
            ComboBox CBX = comboBox4;
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[KINDS],[NAMES],[VALUE] FROM [TKKPI].[dbo].[TBPARA] WHERE [KINDS]='frmREPORTSALSA-REPORTS' ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);

            CBX.DataSource = dt.DefaultView;
            CBX.ValueMember = "NAMES";
            CBX.DisplayMember = "NAMES";
            sqlConn.Close();

            CBX.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }
        public void SETFASTREPORT(string REPORTS,string SDAYS,string EDAYS)
        {
            StringBuilder SQL1 = new StringBuilder();


            Report report1 = new Report();
            //report1.Load(@"REPORT\商品銷售-業務門市觀光.frx");
            //SQL1 = SETSQL1();

            if (REPORTS.Equals("查業務門市觀光-數量"))
            {
                report1.Load(@"REPORT\商品銷售-業務門市觀光-數量.frx");
                SQL1 = SETSQL1(SDAYS, EDAYS);
            }
            else if (REPORTS.Equals("查業務門市觀光-金額"))
            {
                report1.Load(@"REPORT\商品銷售-業務門市觀光-金額.frx");
                SQL1 = SETSQL2(SDAYS, EDAYS);
            }


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

        public StringBuilder SETSQL1(string SDAYS, string EDAYS)
        {
            StringBuilder SB = new StringBuilder();
            SDAYS = SDAYS + "01";
            EDAYS = EDAYS + "31";

            SB.AppendFormat(@"                             
                            SELECT KINDS AS '銷售別',LA005 AS '品號',YEARS AS '年度',MONTHS AS '月份',LA016 AS '銷售數量',MB002 AS '品名',MB003 AS '規格'
                            FROM
                            (
                            SELECT '門市' KINDS,LA005,SUBSTRING(CONVERT(NVARCHAR,LA015,112),1,4) YEARS,SUBSTRING(CONVERT(NVARCHAR,LA015,112),5,2) MONTHS,SUM(LA016-LA019+LA025) LA016
                            FROM [TK].dbo.SASLA
                            WHERE 1=1
                            AND CONVERT(NVARCHAR,LA015,112)>='{0}'
                            AND CONVERT(NVARCHAR,LA015,112)<='{1}'
                            AND LA007 IN ('106501','106502','106503','106504')
                            GROUP BY LA005,SUBSTRING(CONVERT(NVARCHAR,LA015,112),1,4) ,SUBSTRING(CONVERT(NVARCHAR,LA015,112),5,2)
                            UNION
                            SELECT '觀光',LA005,SUBSTRING(CONVERT(NVARCHAR,LA015,112),1,4) YEARS,SUBSTRING(CONVERT(NVARCHAR,LA015,112),5,2) MONTHS,SUM(LA016-LA019+LA025) LA016
                            FROM [TK].dbo.SASLA
                            WHERE 1=1
                            AND CONVERT(NVARCHAR,LA015,112)>='{0}'
                            AND CONVERT(NVARCHAR,LA015,112)<='{1}'
                            AND LA007 IN ('106701')
                            GROUP BY LA005,SUBSTRING(CONVERT(NVARCHAR,LA015,112),1,4) ,SUBSTRING(CONVERT(NVARCHAR,LA015,112),5,2)
                            UNION
                            SELECT '業務',LA005,SUBSTRING(CONVERT(NVARCHAR,LA015,112),1,4) YEARS,SUBSTRING(CONVERT(NVARCHAR,LA015,112),5,2) MONTHS,SUM(LA016-LA019+LA025) LA016
                            FROM [TK].dbo.SASLA
                            WHERE 1=1
                            AND CONVERT(NVARCHAR,LA015,112)>='{0}'
                            AND CONVERT(NVARCHAR,LA015,112)<='{1}'
                            AND LA007 LIKE '117%'
                            GROUP BY LA005,SUBSTRING(CONVERT(NVARCHAR,LA015,112),1,4) ,SUBSTRING(CONVERT(NVARCHAR,LA015,112),5,2)
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.INVMB ON MB001=LA005
                            WHERE (LA005 LIKE '4%' OR LA005 LIKE '5%')

                             ",SDAYS,EDAYS);

            talbename = "TEMPds1";

            return SB;

        }

        public StringBuilder SETSQL2(string SDAYS, string EDAYS)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                         
                          ");

            talbename = "TEMPds1";

            return SB;

        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(comboBox4.Text.ToString(),dateTimePicker1.Value.ToString("yyyyMM"), dateTimePicker2.Value.ToString("yyyyMM"));
        }

        #endregion
    }
}
