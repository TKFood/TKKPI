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
    public partial class frmREPORTINVLA : Form
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


        public frmREPORTINVLA()
        {
            InitializeComponent();
        }


        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(textBox1.Text.Trim());
            Report report1 = new Report();

            report1.Load(@"REPORT\呆滯-久未使用物料.frx");

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

        public StringBuilder SETSQL(string DAYS)
        {


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                            SELECT LA001 AS '品號',MB002 AS '品名',LA011 AS '庫存數',LA009 AS '庫別',LA016 AS '批號',CONVERT(INT,(MB065/MB064)*LA011) AS '庫存金額'
                            ,SUBSTRING(最近生產日商品,1,8) AS '最近生產日'
                            ,SUBSTRING(最近生產日商品,10,100) AS '最近生產商品'
                            ,MB064,MB065
                            FROM (
                            SELECT LA001,MB002,SUM(LA005*LA011) AS 'LA011',LA009,LA016,MB065,MB064
                            ,(SELECT TOP 1 TC003+'-'+TA034 FROM [TK].dbo.MOCTC,[TK].dbo.MOCTE,[TK].dbo.MOCTA WHERE TC001=TE001 AND TC002=TE002 AND TE011=TA001 AND TE012=TA002 AND TE004=LA001 ORDER BY TC003 DESC) AS '最近生產日商品'
                            FROM [TK].dbo.INVLA,[TK].dbo.INVMB
                            WHERE LA001=MB001
                            AND MB064>0
                            AND ISNULL(LA016,'')<>''
                            AND LA016 LIKE '2%'
                            AND (LA001 LIKE '201%' OR LA001 LIKE '202%' OR LA001 LIKE '203%' OR LA001 LIKE '204%'  OR LA001 LIKE '205%'  OR LA001 LIKE '206%')
                            --AND LA001='202003185'
                            GROUP BY LA001,MB002,LA009,LA016,MB065,MB064
                            HAVING SUM(LA005*LA011)>0
                            ) AS TEMP 
                            WHERE DATEDIFF(DAY,SUBSTRING(最近生產日商品,1,8),GETDATE() )>={0}
                            ORDER BY SUBSTRING(最近生產日商品,10,100)

                            ", DAYS);


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
