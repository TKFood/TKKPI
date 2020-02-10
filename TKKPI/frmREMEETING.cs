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
            //本月最後1天
            DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);
            dateTimePicker2.Value = LastDay;

            //下月第一天
            DateTime NEXTMONTH = new DateTime(DateTime.Now.Year, DateTime.Now.AddMonths(1).Month, 1);
            dateTimePicker3.Value = NEXTMONTH;

            //本年年末
            DateTime endYear = new DateTime(DateTime.Now.Year, 12, 31);
            dateTimePicker4.Value = endYear;

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
            report1.Preview = previewControl1;
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
            SB.AppendFormat(" SELECT '國內' AS '國別','何姍怡' AS '業務員',TC008 AS '交易幣別',  SUM(TD012) AS '金額' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002");
            SB.AppendFormat(" AND TD013>='{0}' AND TD013<='{1}' ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TC001 NOT IN ('A223') AND TD016='N' AND TC006='100005' AND TC005='106200'");
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


        public void SETFASTREPORT2()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2();
            Report report2 = new Report();
            report2.Load(@"REPORT\未出訂單業績統計.frx");

            report2.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report2.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

           
            report2.Preview = previewControl2;
            report2.Show();
        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

           
            SB.AppendFormat(" DECLARE @DAY1 NVARCHAR(8)");
            SB.AppendFormat(" DECLARE @DAY2 NVARCHAR(8)");
            SB.AppendFormat(" SET @DAY1 = '{0}'",dateTimePicker3.Value.ToString("yyyyMM")+"01");
            SB.AppendFormat(" SET @DAY2 = '{0}'",dateTimePicker4.Value.ToString("yyyyMM") + "31");
            SB.AppendFormat("    ");
            SB.AppendFormat(" SELECT ");
            SB.AppendFormat(" 類別,年月,國別,業務員,ISNULL(SUM(Tmoney),0) AS 'Tmoney'   FROM");
            SB.AppendFormat(" ( ");
            SB.AppendFormat(" SELECT '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','劉莉琴' AS '業務員',TC008 AS '交易幣別',  ");
            SB.AppendFormat(" SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD  ");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002 AND TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN ('A223') AND TD016='N' AND TC006='140049' AND TC005='106000'  ");
            SB.AppendFormat(" GROUP BY SUBSTRING(TD013,1,6),TC008  ");
            SB.AppendFormat(" UNION ALL");
            SB.AppendFormat(" SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','劉莉琴' AS '業務員',TC008 AS '交易幣別',  ");
            SB.AppendFormat(" SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD  ");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002 AND TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN ('A223') AND TD016='N' AND TC006='140049' AND TC005='106000'  ");
            SB.AppendFormat(" GROUP BY SUBSTRING(TD013,1,6),TC008 ");
            SB.AppendFormat(" UNION ALL  ");
            SB.AppendFormat(" SELECT  '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','蔡顏鴻' AS '業務員',TC008 AS '交易幣別',");
            SB.AppendFormat(" SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD  ");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN ('A223') AND TD016='N' AND TC006='140078' AND TC005='106200'  ");
            SB.AppendFormat(" GROUP BY SUBSTRING(TD013,1,6),TC008 ");
            SB.AppendFormat(" UNION ALL ");
            SB.AppendFormat(" SELECT  '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','蔡顏鴻' AS '業務員',TC008 AS '交易幣別',");
            SB.AppendFormat(" SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD  ");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN ('A223') AND TD016='N' AND TC006='140078' AND TC005='106200'  ");
            SB.AppendFormat(" GROUP BY SUBSTRING(TD013,1,6),TC008 ");
            SB.AppendFormat(" UNION ALL  ");
            SB.AppendFormat(" SELECT  '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','何姍怡' AS '業務員',TC008 AS '交易幣別',");
            SB.AppendFormat(" SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD  ");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN ('A223') AND TD016='N' AND TC006='100005' AND TC005='106200'  ");
            SB.AppendFormat(" GROUP BY SUBSTRING(TD013,1,6),TC008 ");
            SB.AppendFormat(" UNION ALL ");
            SB.AppendFormat(" SELECT  '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','何姍怡' AS '業務員',TC008 AS '交易幣別',");
            SB.AppendFormat(" SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD  ");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN ('A223') AND TD016='N' AND TC006='100005' AND TC005='106200'  ");
            SB.AppendFormat(" GROUP BY SUBSTRING(TD013,1,6),TC008 ");
            SB.AppendFormat(" UNION ALL    ");
            SB.AppendFormat(" SELECT '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','大陸' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  ");
            SB.AppendFormat(" SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD ");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN ('A223') AND TD016='N' AND TC006='160155' AND TC005='106800'  ");
            SB.AppendFormat(" GROUP BY SUBSTRING(TD013,1,6),TC008 ");
            SB.AppendFormat(" UNION ALL   ");
            SB.AppendFormat(" SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','大陸' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  ");
            SB.AppendFormat(" SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD  ");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 IN ('A223') AND TD016='N' AND TC006='160155' AND TC005='106800'  ");
            SB.AppendFormat(" GROUP BY SUBSTRING(TD013,1,6),TC008  ");
            SB.AppendFormat(" UNION ALL ");
            SB.AppendFormat(" SELECT '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國外' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  ");
            SB.AppendFormat(" SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney'");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD ");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN ('A223') AND TD016='N' AND TC006='160155' AND TC005='106300' ");
            SB.AppendFormat(" GROUP BY SUBSTRING(TD013,1,6),TC008  ");
            SB.AppendFormat(" UNION ALL ");
            SB.AppendFormat(" SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國外' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  ");
            SB.AppendFormat(" SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney'");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD ");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 IN ('A223') AND TD016='N' AND TC006='160155' AND TC005='106300' ");
            SB.AppendFormat(" GROUP BY SUBSTRING(TD013,1,6),TC008  ");
            SB.AppendFormat(" UNION ALL ");
            SB.AppendFormat(" SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','營銷部' AS '業務員',TC008 AS '交易幣別',  ");
            SB.AppendFormat(" SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' ");
            SB.AppendFormat(" FROM [TK].dbo.COPTC,[TK].dbo.COPTD  ");
            SB.AppendFormat(" WHERE TC001=TD001 AND TC002=TD002 AND TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN ('A228') AND TD016='N'   ");
            SB.AppendFormat(" GROUP BY SUBSTRING(TD013,1,6),TC008  ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ) AS TEMP");
            SB.AppendFormat(" GROUP BY 年月,國別,業務員,類別");
            SB.AppendFormat(" ORDER BY 年月,國別,業務員,類別");
            SB.AppendFormat("    ");
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
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2();
        }
        #endregion


    }
}
