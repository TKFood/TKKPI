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
    public partial class frmREANALYTICS : Form
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

        DateTime LASTYEARSTART;
        DateTime LASTYEAREND;
        DateTime NOWYEARSTART;
        DateTime NOWYEAREND;

        public frmREANALYTICS()
        {
            InitializeComponent();

            SETDATETIME();
        }

        #region FUNCTION
        public void SETDATETIME()
        {
            DateTime dt =Convert.ToDateTime(DateTime.Now.Year.ToString() + "/1/1");

            dateTimePicker1.Value = dt;
        }

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();
            StringBuilder SQL3 = new StringBuilder();

            SQL1 = SETSQL1();
            SQL2 = SETSQL2();
            SQL3 = SETSQL3();
            Report report1 = new Report();
            report1.Load(@"REPORT\銷售分析-全公司.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();
            TableDataSource table1 = report1.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL2.ToString();
            TableDataSource table2 = report1.GetDataSource("Table2") as TableDataSource;
            table2.SelectCommand = SQL3.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL1()
        {
            LASTYEARSTART = dateTimePicker1.Value;
            LASTYEARSTART = LASTYEARSTART.AddYears(-1);

            LASTYEAREND = dateTimePicker2.Value;
            LASTYEAREND = LASTYEAREND.AddYears(-1);

            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT 部門,銷售未稅金額,銷售未稅金額/SUM(銷售未稅金額) OVER () AS 百分比,去年同期銷售未稅金額");
            SB.AppendFormat(" FROM (");
            SB.AppendFormat(" SELECT '業務' AS '部門',SUM(TH037)  AS '銷售未稅金額'");
            SB.AppendFormat(" ,(SELECT SUM(TH037)");
            SB.AppendFormat(" FROM [TK].dbo.COPTH WITH(NOLOCK),[TK].dbo.COPTG WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.INVMA WITH(NOLOCK),[TK].dbo.INVLA WITH(NOLOCK)");
            SB.AppendFormat(" WHERE TG001=TH001 AND TG002=TH002");
            SB.AppendFormat(" AND MB001=TH004");
            SB.AppendFormat(" AND MB007=MA002 AND MA001='3'");
            SB.AppendFormat(" AND LA006=TH001 AND LA007=TH002 AND LA008=TH003");
            SB.AppendFormat(" AND (TH004 LIKE '4%' OR TH004 LIKE '5%')");
            SB.AppendFormat(" AND TH020='Y'");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", LASTYEARSTART.ToString("yyyyMMdd"), LASTYEAREND.ToString("yyyyMMdd"));
            SB.AppendFormat(" ) AS '去年同期銷售未稅金額' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTH WITH(NOLOCK),[TK].dbo.COPTG WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.INVMA WITH(NOLOCK),[TK].dbo.INVLA WITH(NOLOCK)");
            SB.AppendFormat(" WHERE TG001=TH001 AND TG002=TH002");
            SB.AppendFormat(" AND MB001=TH004");
            SB.AppendFormat(" AND MB007=MA002 AND MA001='3'");
            SB.AppendFormat(" AND LA006=TH001 AND LA007=TH002 AND LA008=TH003");
            SB.AppendFormat(" AND (TH004 LIKE '4%' OR TH004 LIKE '5%')");
            SB.AppendFormat(" AND TH020='Y'");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" UNION ALL ");
            SB.AppendFormat(" SELECT '營銷' AS '部門', SUM(TB031)  AS '銷售未稅金額'");
            SB.AppendFormat(" ,(");
            SB.AppendFormat(" SELECT SUM(TB031)");
            SB.AppendFormat(" FROM [TK].dbo.POSTB WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.INVMA WITH(NOLOCK)");
            SB.AppendFormat(" WHERE TB010=MB001");
            SB.AppendFormat(" AND MB007=MA002 AND MA001='3'");
            SB.AppendFormat(" AND (TB010 LIKE '4%' OR TB010 LIKE '5%')");
            SB.AppendFormat(" AND TB001>='{0}' AND TB001<='{1}'",LASTYEARSTART.ToString("yyyyMMdd"),LASTYEAREND.ToString("yyyyMMdd"));
            SB.AppendFormat(" )");
            SB.AppendFormat(" FROM [TK].dbo.POSTB WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.INVMA WITH(NOLOCK)");
            SB.AppendFormat(" WHERE TB010=MB001");
            SB.AppendFormat(" AND MB007=MA002 AND MA001='3'");
            SB.AppendFormat(" AND (TB010 LIKE '4%' OR TB010 LIKE '5%')");
            SB.AppendFormat(" AND TB001>='{0}' AND TB001<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" ) AS TEMP ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");


            return SB;

        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT 業務,銷售未稅金額,去年同期銷售未稅金額,銷售未稅金額/SUM(銷售未稅金額) OVER () AS 百分比");
            SB.AppendFormat(" FROM (");
            SB.AppendFormat(" SELECT MV002 AS '業務',SUM(TH037)  AS '銷售未稅金額'");
            SB.AppendFormat(" ,(SELECT ISNULL(SUM(TH037),0)");
            SB.AppendFormat(" FROM [TK].dbo.COPTH TH WITH(NOLOCK),[TK].dbo.COPTG TG WITH(NOLOCK),[TK].dbo.INVMB MB WITH(NOLOCK),[TK].dbo.INVLA LA WITH(NOLOCK),[TK].dbo.COPMA CMA WITH(NOLOCK),[TK].dbo.CMSMR  MR WITH(NOLOCK),[TK].dbo.CMSMV MV WITH(NOLOCK)");
            SB.AppendFormat(" WHERE TG.TG001=TH.TH001 AND TG.TG002=TH.TH002");
            SB.AppendFormat(" AND MB.MB001=TH.TH004");
            SB.AppendFormat(" AND LA.LA006=TH.TH001 AND LA.LA007=TH.TH002 AND LA.LA008=TH.TH003");
            SB.AppendFormat(" AND (TH.TH004 LIKE '4%' OR TH.TH004 LIKE '5%')");
            SB.AppendFormat(" AND TH.TH020='Y'");
            SB.AppendFormat(" AND MR.MR001='4' AND MR.MR002=CMA.MA019");
            SB.AppendFormat(" AND TG.TG004=CMA.MA001");
            SB.AppendFormat(" AND MV.MV001=TG006");
            SB.AppendFormat(" AND TG.TG003>='{0}' AND TG.TG003<='{1}'", LASTYEARSTART.ToString("yyyyMMdd"), LASTYEAREND.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND MV.MV002=CMSMV.MV002");
            SB.AppendFormat(" ) AS  '去年同期銷售未稅金額'");
            SB.AppendFormat(" FROM [TK].dbo.COPTH WITH(NOLOCK),[TK].dbo.COPTG WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.INVLA WITH(NOLOCK),[TK].dbo.COPMA WITH(NOLOCK),[TK].dbo.CMSMR WITH(NOLOCK),[TK].dbo.CMSMV WITH(NOLOCK)");
            SB.AppendFormat(" WHERE TG001=TH001 AND TG002=TH002");
            SB.AppendFormat(" AND MB001=TH004");
            SB.AppendFormat(" AND LA006=TH001 AND LA007=TH002 AND LA008=TH003");
            SB.AppendFormat(" AND (TH004 LIKE '4%' OR TH004 LIKE '5%')");
            SB.AppendFormat(" AND TH020='Y'");
            SB.AppendFormat(" AND MR001='4' AND MR002=COPMA.MA019");
            SB.AppendFormat(" AND TG004=COPMA.MA001");
            SB.AppendFormat(" AND MV001=TG006");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" GROUP BY MV002");
            SB.AppendFormat(" HAVING  SUM(TH037) >0");
            SB.AppendFormat(" ) AS TEMP");
            SB.AppendFormat(" ORDER BY 銷售未稅金額 DESC");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        public StringBuilder SETSQL3()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT 門市,銷售未稅金額,銷售未稅金額/SUM(銷售未稅金額) OVER () AS 百分比,去年同期銷售未稅金額");
            SB.AppendFormat(" FROM (");
            SB.AppendFormat(" SELECT WSCMA.MA002 AS '門市', SUM(TB031)  AS '銷售未稅金額'");
            SB.AppendFormat(" ,(");
            SB.AppendFormat(" SELECT SUM(TB031)");
            SB.AppendFormat(" FROM [TK].dbo.POSTB  TB WITH(NOLOCK),[TK].dbo.INVMB MB WITH(NOLOCK),[TK].dbo.WSCMA  WMA WITH(NOLOCK),[TK].dbo.INVMA IMA WITH(NOLOCK)");
            SB.AppendFormat(" WHERE TB.TB010=MB.MB001");
            SB.AppendFormat(" AND WMA.MA001=TB.TB002");
            SB.AppendFormat(" AND MB007=IMA.MA002 AND IMA.MA001='3'");
            SB.AppendFormat(" AND (TB.TB010 LIKE '4%' OR TB.TB010 LIKE '5%')");
            SB.AppendFormat(" AND TB.TB001>='{0}' AND TB.TB001<='{1}'", LASTYEARSTART.ToString("yyyyMMdd"), LASTYEAREND.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND WMA.MA002=WSCMA.MA002");
            SB.AppendFormat(" ) AS  '去年同期銷售未稅金額'");
            SB.AppendFormat(" FROM [TK].dbo.POSTB WITH(NOLOCK),[TK].dbo.INVMB WITH(NOLOCK),[TK].dbo.WSCMA WITH(NOLOCK),[TK].dbo.INVMA WITH(NOLOCK)");
            SB.AppendFormat(" WHERE TB010=MB001");
            SB.AppendFormat(" AND WSCMA.MA001=TB002");
            SB.AppendFormat(" AND MB007=INVMA.MA002 AND INVMA.MA001='3'");
            SB.AppendFormat(" AND (TB010 LIKE '4%' OR TB010 LIKE '5%')");
            SB.AppendFormat(" AND TB001>='{0}' AND TB001<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" GROUP BY  WSCMA.MA002");
            SB.AppendFormat(" HAVING SUM(TB031)>0");
            SB.AppendFormat(" ) AS TEMP ");
            SB.AppendFormat(" ORDER BY 銷售未稅金額 DESC");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
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
