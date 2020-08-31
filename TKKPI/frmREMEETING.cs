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

            SB.AppendFormat(@" 
                            SELECT '國內' AS '國別','劉莉琴' AS '業務員',TC008 AS '交易幣別',  SUM(TD012) AS '金額' 
                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD
                            WHERE TC001 = TD001 AND TC002 = TD002
                            AND TD013 >= '{0}' AND TD013 <= '{1}'
                            AND TC001 IN('A221', 'A222', 'A225', 'A226') AND TD016 = 'N' AND TC006 = '140049'
                            GROUP BY TC008
                            UNION ALL
                            SELECT '國內' AS '國別', '蔡顏鴻' AS '業務員', TC008 AS '交易幣別', SUM(TD012) AS '金額'
                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD
                            WHERE TC001 = TD001 AND TC002 = TD002
                            AND TD013 >= '{0}' AND TD013 <= '{1}'
                            AND TC001 IN('A221', 'A222', 'A225', 'A226') AND TD016 = 'N' AND TC006 = '140078'
                            GROUP BY TC008
                            UNION ALL
                            SELECT '國內' AS '國別', '陳帟靜' AS '業務員', TC008 AS '交易幣別', SUM(TD012) AS '金額'
                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD
                            WHERE TC001 = TD001 AND TC002 = TD002
                            AND TD013 >= '{0}' AND TD013 <= '{1}'
                            AND TC001 IN('A221', 'A222', 'A225', 'A226') AND TD016 = 'N' AND TC006 = '160123'
                            GROUP BY TC008
                            UNION ALL
                            SELECT '國內' AS '國別', '黃鈺涵' AS '業務員', TC008 AS '交易幣別', SUM(TD012) AS '金額'
                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD
                            WHERE TC001 = TD001 AND TC002 = TD002
                             AND TD013 >= '{0}' AND TD013 <= '{1}'
                            AND TC001 IN('A221', 'A222', 'A225', 'A226') AND TD016 = 'N' AND TC006 = '190003'
                            GROUP BY TC008
                            UNION ALL
                            SELECT '國內' AS '國別', '何姍怡' AS '業務員', TC008 AS '交易幣別', SUM(TD012) AS '金額'
                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD
                            WHERE TC001 = TD001 AND TC002 = TD002
                            AND TD013 >= '{0}' AND TD013 <= '{1}'
                            AND TC001 IN('A221', 'A222', 'A225', 'A226') AND TD016 = 'N' AND TC006 = '100005'
                            GROUP BY TC008
                            UNION ALL
                            SELECT '大陸' AS '國別', '洪櫻芬' AS '業務員', TC008 AS '交易幣別', SUM(TD012) AS '金額'
                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD
                            WHERE TC001 = TD001 AND TC002 = TD002
                            AND TD013 >= '{0}' AND TD013 <= '{1}'
                            AND TC001 IN('A221', 'A222', 'A225', 'A226') AND TD016 = 'N' AND TC006 = '160155'
                            AND TC004 IN(SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('010'))
                            GROUP BY TC008
                            UNION ALL
                            SELECT '國外' AS '國別', '洪櫻芬' AS '業務員', TC008 AS '交易幣別', SUM(TD012) AS '金額'
                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD
                            WHERE TC001 = TD001 AND TC002 = TD002
                            AND TD013 >= '{0}' AND TD013 <= '{1}'
                            AND TC001 IN('A221', 'A222', 'A225', 'A226') AND TD016 = 'N' AND TC006 = '160155'
                            AND TC004  NOT IN(SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('010'))
                            GROUP BY TC008
                            UNION ALL
                            SELECT '國外' AS '國別', '王琇平' AS '業務員', TC008 AS '交易幣別', SUM(TD012) AS '金額'
                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD
                            WHERE TC001 = TD001 AND TC002 = TD002
                            AND TD013 >= '{0}' AND TD013 <= '{1}'
                            AND TC001 IN('A221', 'A222', 'A225', 'A226') AND TD016 = 'N' AND TC006 = '190024'
                            AND TC004 NOT IN(SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('001'))
                            GROUP BY TC008
                            UNION ALL
                            SELECT '國內' AS '國別', '王琇平' AS '業務員', TC008 AS '交易幣別', SUM(TD012) AS '金額'
                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD
                            WHERE TC001 = TD001 AND TC002 = TD002
                            AND TD013 >= '{0}' AND TD013 <= '{1}'
                            AND TC001 IN('A221', 'A222', 'A225', 'A226') AND TD016 = 'N' AND TC006 = '190024'
                            AND TC004 IN(SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('001'))
                            GROUP BY TC008
                            ",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

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

           
            SB.AppendFormat(@" DECLARE @DAY1 NVARCHAR(8)
                               DECLARE @DAY2 NVARCHAR(8)
                               SET @DAY1 = '{0}'
                               SET @DAY2 = '{1}'
    
                             SELECT 
                             類別,年月,國別,業務員,ISNULL(SUM(Tmoney),0) AS 'Tmoney'   FROM
                             ( 
                             SELECT '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','劉莉琴' AS '業務員',TC008 AS '交易幣別',  
                             SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='140049'
                             GROUP BY SUBSTRING(TD013,1,6),TC008  
                             UNION ALL
                             SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','劉莉琴' AS '業務員',TC008 AS '交易幣別',  
                             SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='140049' 
                             GROUP BY SUBSTRING(TD013,1,6),TC008 
                             UNION ALL  
                             SELECT  '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','蔡顏鴻' AS '業務員',TC008 AS '交易幣別',
                             SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='140078' 
                             GROUP BY SUBSTRING(TD013,1,6),TC008 
                             UNION ALL 
                             SELECT  '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','蔡顏鴻' AS '業務員',TC008 AS '交易幣別',
                             SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='140078' 
                             GROUP BY SUBSTRING(TD013,1,6),TC008 
                             UNION ALL  
                             SELECT  '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','陳帟靜' AS '業務員',TC008 AS '交易幣別',
                             SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='160123' 
                             GROUP BY SUBSTRING(TD013,1,6),TC008 
                             UNION ALL 
                             SELECT  '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','陳帟靜' AS '業務員',TC008 AS '交易幣別',
                             SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='160123'  
                             GROUP BY SUBSTRING(TD013,1,6),TC008 
                             UNION ALL  
                             SELECT  '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','黃鈺涵' AS '業務員',TC008 AS '交易幣別',
                             SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='190003' 
                             GROUP BY SUBSTRING(TD013,1,6),TC008 
                             UNION ALL 
                             SELECT  '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','黃鈺涵' AS '業務員',TC008 AS '交易幣別',
                             SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='190003' 
                             GROUP BY SUBSTRING(TD013,1,6),TC008 
                             UNION ALL  
                             SELECT  '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','何姍怡' AS '業務員',TC008 AS '交易幣別',
                             SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='100005' 
                             GROUP BY SUBSTRING(TD013,1,6),TC008 
                             UNION ALL 
                             SELECT  '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','何姍怡' AS '業務員',TC008 AS '交易幣別',
                             SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='100005'  
                             GROUP BY SUBSTRING(TD013,1,6),TC008 
                             UNION ALL    
                             SELECT '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','大陸' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  
                             SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD 
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='160155' 
                             AND TC004 IN (SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('010'))  
                             GROUP BY SUBSTRING(TD013,1,6),TC008 
                             UNION ALL   
                             SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','大陸' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  
                             SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='160155' 
                             AND TC004 IN (SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('010'))    
                             GROUP BY SUBSTRING(TD013,1,6),TC008  
                             UNION ALL 
                             SELECT '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國外' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  
                             SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney'
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD 
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='160155'
                             AND TC004 NOT IN (SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('010'))   
                             GROUP BY SUBSTRING(TD013,1,6),TC008  
                             UNION ALL 
                             SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國外' AS '國別','洪櫻芬' AS '業務員',TC008 AS '交易幣別',  
                             SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney'
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD 
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='160155' 
                              AND TC004 NOT IN (SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('010'))  
                             GROUP BY SUBSTRING(TD013,1,6),TC008  
                             UNION ALL 
                             SELECT '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國外' AS '國別','王琇平' AS '業務員',TC008 AS '交易幣別',  
                             SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney'
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD 
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='190024' 
                             AND TC004 NOT IN(SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('001'))
                             GROUP BY SUBSTRING(TD013,1,6),TC008  
                             UNION ALL 
                             SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國外' AS '國別','王琇平' AS '業務員',TC008 AS '交易幣別',  
                             SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney'
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD 
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='190024' 
                             AND TC004 NOT IN(SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('001'))
                             GROUP BY SUBSTRING(TD013,1,6),TC008  
                             UNION ALL 
                             SELECT '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','王琇平' AS '業務員',TC008 AS '交易幣別',  
                             SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney'
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD 
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='190024' 
                             AND TC004 IN(SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('001'))
                             GROUP BY SUBSTRING(TD013,1,6),TC008  
                             UNION ALL 
                             SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','王琇平' AS '業務員',TC008 AS '交易幣別',  
                             SUM(TD012) AS '金額'  ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney'
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD 
                             WHERE TC001=TD001 AND TC002=TD002 AND  TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN('A221', 'A222', 'A225', 'A226') AND TD016='N' AND TC006='190024' 
                             AND TC004 IN(SELECT MA001 FROM[TK].dbo.COPMA WHERE MA019 IN('001'))
                             GROUP BY SUBSTRING(TD013,1,6),TC008 
                             UNION ALL 
                             SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月','國內' AS '國別','營銷部' AS '業務員',TC008 AS '交易幣別',  
                             SUM(TD012) AS '金額'   ,CASE WHEN TC008='NTD'  THEN SUM(TD012)*1 ELSE CASE WHEN TC008='RMB'  THEN SUM(TD012)*4 ELSE CASE WHEN TC008='USD'  THEN SUM(TD012)*30  END END END AS 'Tmoney' 
                             FROM [TK].dbo.COPTC,[TK].dbo.COPTD  
                             WHERE TC001=TD001 AND TC002=TD002 AND TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN ('A228') AND TD016='N'   
                             GROUP BY SUBSTRING(TD013,1,6),TC008  
 
                             ) AS TEMP
                             GROUP BY 年月,國別,業務員,類別
                             ORDER BY 年月,國別,業務員,類別 
                            ", dateTimePicker3.Value.ToString("yyyyMM") + "01", dateTimePicker4.Value.ToString("yyyyMM") + "31");



            return SB;

        }
        public void SETFASTREPORT3()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3();
            Report report3 = new Report();
            report3.Load(@"REPORT\工作交辨.frx");

            report3.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
            TableDataSource table = report3.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();


            report3.Preview = previewControl3;
            report3.Show();
        }

        public StringBuilder SETSQL3()
        {
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(" SELECT USERA.NAME AS '交辨人',USERB.NAME AS '被交辨人',CONVERT(nvarchar,TEMP.CREATE_TIME,112) AS '交辨時間',TEMP.DESCRIPTION AS '交辨內容',CONVERT(nvarchar,TEMP.END_TIME,112) AS '希望交辨完成時間',CASE WHEN TEMP.WORK_STATE='Proceeding' THEN '處理中' WHEN TEMP.WORK_STATE='Completed' THEN '完成' WHEN TEMP.WORK_STATE='Audit' THEN '完成但未確認' WHEN TEMP.WORK_STATE='NotYetBegin' THEN '未回覆' ELSE TEMP.WORK_STATE END AS '交辨狀況',CONVERT(nvarchar,TEMP.COMPLETE_TIME,112) AS '交辨完成時間',TEMP.COMPLETE_DESC AS '交辨回覆'");
            SB.AppendFormat(" FROM (");
            SB.AppendFormat(" SELECT [TB_EIP_SCH_DEVOLVE].[CREATE_TIME],[TB_EIP_SCH_DEVOLVE].[CREATE_USER],[TB_EIP_SCH_DEVOLVE].[DESCRIPTION],[TB_EIP_SCH_DEVOLVE].[END_TIME]");
            SB.AppendFormat(" ,[TB_EIP_SCH_WORK].[EXECUTE_USER],[TB_EIP_SCH_WORK].[COMPLETE_TIME],[TB_EIP_SCH_WORK].[DEVOLVE_GUID],[TB_EIP_SCH_WORK].[WORK_STATE],[TB_EIP_SCH_WORK].[COMPLETE_DESC]");
            SB.AppendFormat(" FROM [UOF].dbo.TB_EIP_SCH_DEVOLVE, [UOF].dbo.TB_EIP_SCH_WORK");
            SB.AppendFormat(" WHERE TB_EIP_SCH_DEVOLVE.DEVOLVE_GUID=TB_EIP_SCH_WORK.DEVOLVE_GUID");
            SB.AppendFormat(" AND TB_EIP_SCH_DEVOLVE.[CREATE_USER]<>TB_EIP_SCH_WORK.[EXECUTE_USER]");
            SB.AppendFormat(" ) AS TEMP");
            SB.AppendFormat(" LEFT JOIN [UOF].dbo.TB_EB_USER USERA ON USERA.[USER_GUID]=TEMP.[CREATE_USER]");
            SB.AppendFormat(" LEFT JOIN [UOF].dbo.TB_EB_USER USERB ON USERB.[USER_GUID]=TEMP.[EXECUTE_USER]");
            SB.AppendFormat(" WHERE TEMP.WORK_STATE NOT IN ('Completed')");
            SB.AppendFormat(" ORDER BY [CREATE_TIME]");
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
        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3();
        }
        #endregion


    }
}
