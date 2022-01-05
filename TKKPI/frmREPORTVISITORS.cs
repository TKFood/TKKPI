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
    public partial class frmREPORTVISITORS : Form
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

        public frmREPORTVISITORS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();
            StringBuilder SQL3 = new StringBuilder();

            SQL1 = SETSQL();
            SQL2 = SETSQL2();
            SQL3 = SETSQL3();
            Report report1 = new Report();
            report1.Load(@"REPORT\營銷來客報表.frx");

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
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL2.ToString();
            TableDataSource table2= report1.GetDataSource("Table2") as TableDataSource;
            table2.SelectCommand = SQL3.ToString();



            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT TT002,STORESNAME,YEARS,MONTHS,NUMS,TT008,AVGTT011,SUMTT011,ROUND(TT008/NUMS,4) AS PCTS
                            FROM (
                            SELECT TT002,STORESNAME,YEARS,MONTHS,SUM(Fin_data+Fout_data)/2 AS NUMS
                            ,(SELECT ISNULL(SUM(TT008),0) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND TT001 LIKE YEARS+RIGHT('00'+CAST(MONTHS AS nvarchar(10)),2) +'%') AS 'TT008'
                            ,(SELECT ISNULL(SUM(TT011)/SUM(TT008),0) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND TT001 LIKE YEARS+RIGHT('00'+CAST(MONTHS AS nvarchar(10)),2) +'%') 'AVGTT011'
                            ,(SELECT ISNULL(SUM(TT011),0) FROM [TK].dbo.POSTT WHERE View_t_visitors.TT002=POSTT.TT002 AND TT001 LIKE YEARS+RIGHT('00'+CAST(MONTHS AS nvarchar(10)),2) +'%') 'SUMTT011'
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704') 
                            AND YEARS='{0}'
                            GROUP BY  TT002,STORESNAME,YEARS,MONTHS
                            )AS TEMP 
                            ORDER BY  TT002,STORESNAME,YEARS,MONTHS
                            ", dateTimePicker1.Value.ToString("yyyy"));

            return SB;

        }
        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT TT002,STORESNAME,YEARS,WEEKS,SUM(Fin_data+Fout_data)/2 AS NUMS
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704') 
                            AND YEARS='{0}'
                            GROUP BY  TT002,STORESNAME,YEARS,WEEKS
                            ORDER BY  TT002,STORESNAME,YEARS,WEEKS
                            ", dateTimePicker1.Value.ToString("yyyy"));

            return SB;

        }

        public StringBuilder SETSQL3()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT TT002,STORESNAME,YEARS,MONTHS,HOURS,SUM(Fin_data+Fout_data)/2 AS NUMS
                            FROM [TKMK].[dbo].[View_t_visitors]
                            WHERE  TT002 IN ('106501','106502','106503','106504','106513','106702','106703','106704') 
                            AND YEARS='{0}'
                            GROUP BY  TT002,STORESNAME,YEARS,MONTHS,HOURS
                            ORDER BY  TT002,STORESNAME,YEARS,MONTHS,CONVERT(INT,HOURS)
                            ", dateTimePicker1.Value.ToString("yyyy"));

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
