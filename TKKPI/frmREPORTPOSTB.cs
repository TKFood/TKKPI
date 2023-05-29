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
    public partial class frmREPORTPOSTB : Form
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

        public frmREPORTPOSTB()
        {
            InitializeComponent();

            SETDATE();
        }

        #region FUNCTION
        public void SETDATE()
        {
            DateTime today = DateTime.Today.AddDays(-1); // 當天日期-1

            // 指定為星期一
            DateTime monday = today;
            while (monday.DayOfWeek != DayOfWeek.Monday)
            {
                monday = monday.AddDays(-1);
            }

            // 指定為星期日
            DateTime sunday = today;
            while (sunday.DayOfWeek != DayOfWeek.Sunday)
            {
                sunday = sunday.AddDays(-1);
            }

            dateTimePicker1.Value = monday;
            dateTimePicker2.Value = sunday;
        }

        public void SETFASTREPORT(string SDAYS, string EDAYS)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(SDAYS, EDAYS);
            Report report1 = new Report();

            report1.Load(@"REPORT\各門市銷售排名及毛利.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

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

        //      --AND TB_EIP_SCH_WORK.SUBJECT  LIKE '%校稿%'
        public StringBuilder SETSQL(string SDAYS, string EDAYS)
        {

            StringBuilder SB = new StringBuilder();
         
            SB.AppendFormat(@"                               
                            SELECT *
                            ,(CASE WHEN 未稅金額>0 AND 成本>0 THEN (未稅金額-成本)/未稅金額 ELSE 0 END) AS '毛利率'
                            FROM
                            (
                            SELECT TB002 AS '門市代' ,MA002 AS '門市',TB010 AS '品號',MB002 AS '品名',SUM(TB019)  AS '銷售數量' ,SUM(TB031)  AS '未稅金額'
                            ,(SELECT SUM(LA013) FROM [TK].dbo.INVLA WHERE LA004>='{0}' AND LA004<='{1}' AND TB002=LA006 AND TB010=LA001) AS  '成本'
                            FROM [TK].dbo.POSTB,[TK].dbo.WSCMA,[TK].dbo.INVMB
                            WHERE 1=1
                            AND MA001=TB002
                            AND TB010=MB001
                            AND TB002 IN (SELECT  [TT002] FROM [TKKPI].[dbo].[SALESTORES])
                            AND (TB010 LIKE '4%' OR TB010 LIKE '5%')
                            AND TB010 NOT LIKE '499%'
                            AND TB010 NOT LIKE '599%'
                            AND TB010 NOT LIKE '506%'
                            AND TB001>='{0}' AND TB001<='{1}'
                            GROUP BY TB002,MA002,TB010,MB002
                            HAVING SUM(TB031)<>0
                            ) AS TEMP
                            ORDER BY 門市代,未稅金額 DESC
                            

                            ", SDAYS, EDAYS);


            return SB;

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyyMMdd"),dateTimePicker2.Value.ToString("yyyyMMdd"));
        }
        #endregion
    }
}
