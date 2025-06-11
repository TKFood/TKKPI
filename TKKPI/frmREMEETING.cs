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
    public partial class frmREMEETING : Form
    {
        int SQLTIMEOUT = 120;

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

           
        }
        private void frmREMEETING_Load(object sender, EventArgs e)
        {
            SETDT();
        }

        #region FUNCTION

        public void SETDT()
        {
            //本月最後1天
            DateTime LastDay = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1).AddDays(-1);
            dateTimePicker2.Value = LastDay;

            //下月第一天
            DateTime NEXTMONTH = new DateTime(DateTime.Now.AddMonths(1).Year, DateTime.Now.AddMonths(1).Month, 1);
            dateTimePicker3.Value = NEXTMONTH;

            DateTime endYear = new DateTime(NEXTMONTH.AddMonths(3).Year, NEXTMONTH.AddMonths(3).Month,1);
            dateTimePicker4.Value = endYear;

            dateTimePicker5.Value = DateTime.Now;
            dateTimePicker6.Value = LastDay;
            dateTimePicker7.Value = NEXTMONTH;
            dateTimePicker8.Value = endYear;


            DateTime today = DateTime.Today; // 當天日期
            // 先退回到上週的某一天（今天 - 7 天），再從那天往回推到週一
            DateTime lastMonday = today.AddDays(-7);
            while (lastMonday.DayOfWeek != DayOfWeek.Monday)
            {
                lastMonday = lastMonday.AddDays(-1);
            }

            DateTime oneWeekAgo = DateTime.Today.AddDays(-7);
            while (oneWeekAgo.DayOfWeek != DayOfWeek.Sunday)
            {
                oneWeekAgo = oneWeekAgo.AddDays(1);
            }           

            dateTimePicker9.Value = lastMonday;
            dateTimePicker10.Value = oneWeekAgo;
            dateTimePicker11.Value = lastMonday;
            dateTimePicker12.Value = oneWeekAgo;


        }
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\未出訂單業績明細V2.frx");

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

            report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                           SELECT 部門,業務員,單別,單名,交易幣別,SUM(金額) 金額,CONVERT(INT,SUM(未出金額)) 未出金額
                            ,CASE WHEN 交易幣別 IN ('NTD') THEN CONVERT(INT,SUM(未出金額))  WHEN 交易幣別 IN ('RMB') THEN CONVERT(INT,SUM(未出金額))*4  WHEN 交易幣別 IN ('USD') THEN CONVERT(INT,SUM(未出金額))*30  WHEN 交易幣別 IN ('HKD') THEN CONVERT(INT,SUM(未出金額))*4 END AS '本幣金額'
                            FROM (
                            SELECT MV004 AS '部門',MV002 AS '業務員',TC001 AS '單別',MQ002  AS '單名',TC008 AS '交易幣別',  (TD012) AS '金額' ,TC016 AS '稅別',(CASE WHEN TC016 IN ('1') THEN ((TD008-TD009)*TD011*TD026)/1.05 ELSE ((TD008-TD009)*TD011*TD026) END) AS '未出金額'
                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.CMSMV,[TK].dbo.CMSMQ
                            WHERE TC001 = TD001 AND TC002 = TD002
                            AND TC006=MV001
                            AND TC001=MQ001
                            AND TC027='Y'
                            AND TD013 >= '{0}' AND TD013 <= '{1}'
                            AND TC001 IN('A221', 'A222', 'A225', 'A226') AND TD016 = 'N'
                            ) AS TEMP
                            GROUP BY 部門,業務員,交易幣別,單別,單名
                            ORDER BY 單別,單名,業務員

                            ", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));

            return SB;

        }


        public void SETFASTREPORT2()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2();
            Report report2 = new Report();
            report2.Load(@"REPORT\未出訂單業績統計.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report2.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report2.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

           
            report2.Preview = previewControl2;
            report2.Show();
        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

            //AND TC027='Y'
            SB.AppendFormat(@" DECLARE @DAY1 NVARCHAR(8)
                               DECLARE @DAY2 NVARCHAR(8)
                               SET @DAY1 = '{0}'
                               SET @DAY2 = '{1}'
    
                             SELECT 
                                類別,年月,部門,業務員,ISNULL(SUM(未出金額),0) AS '未出金額'   
                                ,CASE WHEN 交易幣別 IN ('NTD') THEN CONVERT(INT,SUM(未出金額)) ELSE (CASE WHEN 交易幣別 IN ('RMB') THEN CONVERT(INT,SUM(未出金額))*4  ELSE ( CASE WHEN 交易幣別 IN ('USD') THEN CONVERT(INT,SUM(未出金額))*30 END ) END ) END AS '本幣金額'
                                FROM
                                ( 
                                SELECT '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月',MV004 AS '部門',MV002 AS '業務員', (TD012) AS '金額',TC008 AS '交易幣別' ,TC016 AS '稅別',(CASE WHEN TC016 IN ('1') THEN ((TD008-TD009)*TD011*TD026)/1.05 ELSE ((TD008-TD009)*TD011*TD026) END) AS '未出金額'
                                FROM[TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.CMSMV
                                WHERE TC001 = TD001 AND TC002 = TD002
                                AND TC006=MV001
                                
                                AND TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN ('A221', 'A222', 'A225', 'A226') AND TD016='N' 
                                UNION ALL
                                SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月',MV004 AS '部門',MV002 AS '業務員', (TD012) AS '金額' ,TC008 AS '交易幣別',TC016 AS '稅別',(CASE WHEN TC016 IN ('1') THEN ((TD008-TD009)*TD011*TD026)/1.05 ELSE ((TD008-TD009)*TD011*TD026) END) AS '未出金額'
                                FROM[TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.CMSMV
                                WHERE TC001 = TD001 AND TC002 = TD002
                                AND TC006=MV001
                                
                                AND TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN ('A221', 'A222', 'A225', 'A226') AND TD016='N' 
 
                                ) AS TEMP
                                GROUP BY  年月,部門,業務員,類別,交易幣別
                                ORDER BY   年月,部門,業務員,類別
                            ", dateTimePicker3.Value.ToString("yyyyMM") + "01", dateTimePicker4.Value.ToString("yyyyMM") + "31");



            return SB;

        }
        public void SETFASTREPORT3()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3();
            Report report3 = new Report();
            report3.Load(@"REPORT\工作交辨.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report3.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

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

        public void SETFASTREPORT4(string DATES_TODAY,string DATES_LASTMONTHDAY,string DATES_START,string DATES_END,string DATES_LASTMONDAY, string DATES_LASTSUNDAY,string DATES_CARS_START,string DATES_CARS_END)
        {
            StringBuilder SQL4 = new StringBuilder();
            StringBuilder SQL5 = new StringBuilder();
            StringBuilder SQL6 = new StringBuilder();
            StringBuilder SQL7 = new StringBuilder();

            //訂單未出貨金額
            SQL4 = SETSQL4(DATES_TODAY, DATES_LASTMONTHDAY);
            //未出訂單業績統計
            SQL5 = SETSQL5(DATES_START, DATES_END);
            //各門市上週銷售
            SQL6 = SETSQL6(DATES_LASTMONDAY, DATES_LASTSUNDAY);
            //觀光業績及車次明細表
            SQL7 = SETSQL7(DATES_CARS_START, DATES_CARS_END);

            Report report4 = new Report(); 
            report4.Load(@"REPORT\每週週報表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report4.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            report4.Dictionary.Connections[0].CommandTimeout = SQLTIMEOUT;

            //訂單未出貨金額
            TableDataSource table = report4.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL4.ToString();
            //未出訂單業績統計
            TableDataSource table1 = report4.GetDataSource("Table1") as TableDataSource;
            table1.SelectCommand = SQL5.ToString();
            //各門市上週銷售
            TableDataSource table2 = report4.GetDataSource("Table2") as TableDataSource;
            table2.SelectCommand = SQL6.ToString();
            //觀光業績及車次明
            TableDataSource table3 = report4.GetDataSource("Table3") as TableDataSource;
            table3.SelectCommand = SQL7.ToString();

            report4.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            report4.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));

            report4.Preview = previewControl4;
            report4.Show();
        }

        public StringBuilder SETSQL4(string DATES_TODAY, string DATES_LASTMONTHDAY)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"                
                            SELECT 
                            部門,業務員,單別,單名,交易幣別,SUM(金額) 金額,CONVERT(INT,SUM(未出金額)) 未出金額
                            ,CASE WHEN 交易幣別 IN ('NTD') THEN CONVERT(INT,SUM(未出金額))  WHEN 交易幣別 IN ('RMB') THEN CONVERT(INT,SUM(未出金額))*4  WHEN 交易幣別 IN ('USD') THEN CONVERT(INT,SUM(未出金額))*30  WHEN 交易幣別 IN ('HKD') THEN CONVERT(INT,SUM(未出金額))*4 END AS '本幣金額'
                            FROM (
	                            SELECT MV004 AS '部門',MV002 AS '業務員',TC001 AS '單別',MQ002  AS '單名',TC008 AS '交易幣別',  (TD012) AS '金額' ,TC016 AS '稅別',(CASE WHEN TC016 IN ('1') THEN ((TD008-TD009)*TD011*TD026)/1.05 ELSE ((TD008-TD009)*TD011*TD026) END) AS '未出金額'
	                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.CMSMV,[TK].dbo.CMSMQ
	                            WHERE TC001 = TD001 AND TC002 = TD002
	                            AND TC006=MV001
	                            AND TC001=MQ001
	                            AND TC027='Y'
	                            AND TD013 >= '{0}' AND TD013 <= '{1}'
	                            AND TC001 IN('A221', 'A222', 'A225', 'A226') AND TD016 = 'N'
                            ) AS TEMP
                            GROUP BY 部門,業務員,交易幣別,單別,單名
                            ORDER BY 單別,單名,業務員

                            ", DATES_TODAY, DATES_LASTMONTHDAY);


            return SB;

        }

        public StringBuilder SETSQL5(string DATES_START, string DATES_END)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            DECLARE @DAY1 NVARCHAR(8)
                            DECLARE @DAY2 NVARCHAR(8)
                            SET @DAY1 = '{0}'
                            SET @DAY2 = '{1}'
    
                            SELECT 
                            類別,年月,部門,業務員,ISNULL(SUM(未出金額),0) AS '未出金額'   
                            ,CASE WHEN 交易幣別 IN ('NTD') THEN CONVERT(INT,SUM(未出金額)) ELSE (CASE WHEN 交易幣別 IN ('RMB') THEN CONVERT(INT,SUM(未出金額))*4  ELSE ( CASE WHEN 交易幣別 IN ('USD') THEN CONVERT(INT,SUM(未出金額))*30 END ) END ) END AS '本幣金額'
                            FROM
                            ( 
	                            SELECT '實際訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月',MV004 AS '部門',MV002 AS '業務員', (TD012) AS '金額',TC008 AS '交易幣別' ,TC016 AS '稅別',(CASE WHEN TC016 IN ('1') THEN ((TD008-TD009)*TD011*TD026)/1.05 ELSE ((TD008-TD009)*TD011*TD026) END) AS '未出金額'
	                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.CMSMV
	                            WHERE TC001 = TD001 AND TC002 = TD002
	                            AND TC006=MV001
                                
	                            AND TD013>=@DAY1 AND TD013<=@DAY2 AND TC001  IN ('A221', 'A222', 'A225', 'A226') AND TD016='N' 
	                            UNION ALL
	                            SELECT '預計訂單' AS 類別,SUBSTRING(TD013,1,6) AS '年月',MV004 AS '部門',MV002 AS '業務員', (TD012) AS '金額' ,TC008 AS '交易幣別',TC016 AS '稅別',(CASE WHEN TC016 IN ('1') THEN ((TD008-TD009)*TD011*TD026)/1.05 ELSE ((TD008-TD009)*TD011*TD026) END) AS '未出金額'
	                            FROM[TK].dbo.COPTC,[TK].dbo.COPTD,[TK].dbo.CMSMV
	                            WHERE TC001 = TD001 AND TC002 = TD002
	                            AND TC006=MV001
                                
	                            AND TD013>=@DAY1 AND TD013<=@DAY2 AND TC001 NOT IN ('A221', 'A222', 'A225', 'A226') AND TD016='N' 
 
                            ) AS TEMP
                            GROUP BY  年月,部門,業務員,類別,交易幣別
                            ORDER BY   年月,部門,業務員,類別

                            ", DATES_START, DATES_END);


            return SB;

        }
        public StringBuilder SETSQL6(string DATES_LASTMONDAY, string DATES_LASTSUNDAY)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@" 
                            SELECT *
                            ,(CASE WHEN 未稅金額>0 AND 成本>0 THEN (未稅金額-成本)/未稅金額 ELSE 0 END) AS '毛利率'
                            ,CONVERT(INT,(CASE WHEN 含稅金額>0 AND 銷售數量>0 THEN 含稅金額/銷售數量 ELSE 0 END) ) AS '含稅單價'
                            FROM
                            (
                            SELECT TB002 AS '門市代' ,MA002 AS '門市',TB010 AS '品號',MB002 AS '品名',SUM(TB019)  AS '銷售數量' ,SUM(TB031)  AS '未稅金額',SUM(TB031+TB032) AS '含稅金額'
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
 
                            ", DATES_LASTMONDAY, DATES_LASTSUNDAY);


            return SB;

        }

        public StringBuilder SETSQL7(string DATES_LDATES_CARS_START, string DATES_CARS_END)
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(@"                             
                            SELECT 
                            [INDATES] AS '日期',[YEARS] AS '年',[WEEKS] AS '週',[TOTALMONEYS] AS 銷售組當日業績,[GROUPMONEYS] AS '團客業績',([TOTALMONEYS]-[GROUPMONEYS]) AS '散客業績',[CARNUM] AS '遊覽車次',[CARAVGMONEYS] AS '每車平均業績'
                            FROM [TKMK].[dbo].[TBFACTORYINCOME]
                            WHERE [INDATES]>='{0}' AND [INDATES]<='{1}'
                            UNION ALL
                            -- 總計
                            SELECT 
                              '總計',
                              '',
                              '',
                              SUM([TOTALMONEYS]),
                              SUM([GROUPMONEYS]),
                              SUM([TOTALMONEYS] - [GROUPMONEYS]),
                              SUM([CARNUM]),
                              CASE 
                                WHEN SUM([CARNUM]) = 0 THEN 0 
                                ELSE SUM([GROUPMONEYS]) / SUM([CARNUM]) 
                              END
                            FROM [TKMK].[dbo].[TBFACTORYINCOME]
                            WHERE [INDATES] >= '{0}' AND [INDATES] <= '{1}'
 
                            ", DATES_LDATES_CARS_START, DATES_CARS_END);


            return SB;

        }
        public void ADDTKMK_TBFACTORYINCOME(string SDATES, string EDATES)
        {
            SqlCommand cmd = new SqlCommand();
            SqlTransaction tran;
            int result;
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();


                sbSql.AppendFormat(@" 
                                    DELETE [TKMK].[dbo].[TBFACTORYINCOME]
                                    WHERE INDATES>='{0}' AND INDATES<='{1}'

                                    INSERT INTO [TKMK].[dbo].[TBFACTORYINCOME]
                                    ([INDATES],[YEARS],[WEEKS],[TOTALMONEYS],[GROUPMONEYS],[VISITORMONEYS],[CARNUM],[CARAVGMONEYS])

                                    SELECT INDATES,YEARS,WEEKS,CONVERT(INT,TOTALMONEYS) TOTALMONEYS,CONVERT(INT,GROUPMONEYS)  GROUPMONEYS,CONVERT(INT,VISITORMONEYS)  VISITORMONEYS,CARNUM
                                    ,CASE WHEN CARNUM>0 THEN CONVERT(INT,ROUND(GROUPMONEYS/CARNUM,0))  ELSE 0 END AS 'CARAVGMONEYS'
                                    FROM (
                                    SELECT 
                                    TA001 AS 'INDATES'
                                    ,DATEPART(YEAR, [TA001]) AS YEARS
                                    ,DATEPART(Week, [TA001]) AS WEEKS
                                    ,SUM(TA026) AS 'TOTALMONEYS'
                                    ,(SELECT ROUND(ISNULL(SUM([SALESMMONEYS]),0),0) FROM [TKMK].[dbo].[GROUPSALES] WHERE  [STATUS]='完成接團' AND CONVERT(nvarchar,[CREATEDATES],112)=TA001) AS 'GROUPMONEYS'
                                    ,(SUM(TA026)-(SELECT ROUND(ISNULL(SUM([SALESMMONEYS]),0),0) FROM [TKMK].[dbo].[GROUPSALES] WHERE  [STATUS]='完成接團' AND CONVERT(nvarchar,[CREATEDATES],112)=TA001)) AS 'VISITORMONEYS'
                                    ,(SELECT ISNULL(SUM(CARNUM),0) FROM [TKMK].[dbo].[GROUPSALES] WHERE  [STATUS]='完成接團' AND CONVERT(nvarchar,[CREATEDATES],112)=TA001) AS 'CARNUM'
                                    FROM [TK].dbo.POSTA
                                    WHERE TA002 IN (SELECT  [TA002]  FROM [TKMK].[dbo].[TB_POS_TA002])
                                    AND TA001>='{0}' AND TA001<='{1}'
                                    GROUP BY TA001
                                    ) AS TEMP
                                    ORDER BY INDATES
                                    ", SDATES, EDATES);

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                    MessageBox.Show("完成");

                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
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
        private void button3_Click(object sender, EventArgs e)
        {
            string DATES_TODAY = dateTimePicker5.Value.ToString("yyyyMMdd");
            string DATES_LASTMONTHDAY = dateTimePicker6.Value.ToString("yyyyMMdd");
            string DATES_START = dateTimePicker7.Value.ToString("yyyyMM") + "01";
            string DATES_END = dateTimePicker8.Value.ToString("yyyyMM") + "31";
            string DATES_LASTMONDAY = dateTimePicker9.Value.ToString("yyyyMMdd");
            string DATES_LASTSUNDAY = dateTimePicker10.Value.ToString("yyyyMMdd");
            string DATES_CARS_START = dateTimePicker11.Value.ToString("yyyyMMdd");
            string DATES_CARS_END = dateTimePicker12.Value.ToString("yyyyMMdd");

            ADDTKMK_TBFACTORYINCOME(DATES_CARS_START, DATES_CARS_END);

            SETFASTREPORT4(DATES_TODAY, DATES_LASTMONTHDAY, DATES_START, DATES_END, DATES_LASTMONDAY, DATES_LASTSUNDAY, DATES_CARS_START, DATES_CARS_END);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            string DATES_CARS_START = dateTimePicker11.Value.ToString("yyyyMMdd");
            string DATES_CARS_END = dateTimePicker12.Value.ToString("yyyyMMdd");

            ADDTKMK_TBFACTORYINCOME(DATES_CARS_START, DATES_CARS_END);
        }
        #endregion


    }
}
