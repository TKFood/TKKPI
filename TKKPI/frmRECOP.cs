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

namespace TKKPI
{
    public partial class frmRECOP : Form
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


        public frmRECOP()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\業務商品排名表.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT MV002 AS '業務',TH005 AS '品名',SUM(TH013) AS '金額',SUM(LA011) AS '數量',MB004 AS '單位',SUM(TH013)/SUM(SUM(TH013)) OVER ()  AS '金額百分比'");
            SB.AppendFormat(" FROM [TK].dbo.COPTG, [TK].dbo.COPTH,[TK].dbo.INVLA,[TK].dbo.INVMB,[TK].dbo.CMSMV");
            SB.AppendFormat(" WHERE TG001=TH001 AND TG002=TH002");
            SB.AppendFormat(" AND LA006=TH001 AND LA007=TH002 AND LA008=TH003");
            SB.AppendFormat(" AND MB001=TH004");
            SB.AppendFormat(" AND MV001=TG006");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TG005='{0}'",comboBox1.Text.ToString().Substring(0,6));
            SB.AppendFormat(" AND TG006='{0}'",comboBox2.Text.ToString().Substring(0,6));
            SB.AppendFormat(" GROUP BY MV002,TH005,MB004");
            SB.AppendFormat(" ORDER BY SUM(TH013) DESC");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

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
