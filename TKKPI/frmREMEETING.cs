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
    public partial class frmREMEETING : Form
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

        public frmREMEETING()
        {
            InitializeComponent();

            SETDT();
        }

        #region FUNCTION

        public void SETDT()
        {
            DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);
            dateTimePicker2.Value = LastDay;


        }
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\未出訂單業績明細.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl5;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT '國內' AS '國別','劉莉琴' AS '業務員',TC008 AS '交易幣別',  SUM(TD012) AS '金額' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TC001 NOT IN ('A223') AND TD016='N' AND TC006='140049' AND TC005='106000'");
            SB.AppendFormat(" GROUP BY TC008");
            SB.AppendFormat(" UNION ALL");
            SB.AppendFormat(" SELECT '國內' AS '國別','蔡顏鴻' AS '業務員',TC008 AS '交易幣別',  SUM(TD012) AS '金額' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TC001 NOT IN ('A223') AND TD016='N' AND TC006='140078' AND TC005='106200'");
            SB.AppendFormat(" GROUP BY TC008");
            SB.AppendFormat(" UNION ALL");
            SB.AppendFormat(" SELECT '大陸' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  SUM(TD012) AS '金額'");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TC001 NOT IN ('A223') AND TD016='N' AND TC006='160155' AND TC005='106800'");
            SB.AppendFormat(" GROUP BY TC008");
            SB.AppendFormat(" UNION ALL");
            SB.AppendFormat(" SELECT '國外' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  SUM(TD012) AS '金額'");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TC001 NOT IN ('A223') AND TD016='N' AND TC006='160155' AND TC005='106300'");
            SB.AppendFormat("GROUP BY TC008 ");
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
