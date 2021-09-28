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
    public partial class frmREPORTSWEEKS : Form
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

        public frmREPORTSWEEKS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();
            StringBuilder SQL2 = new StringBuilder();
            StringBuilder SQL3 = new StringBuilder();
            StringBuilder SQL4 = new StringBuilder();

            SQL1 = SETSQL();
            SQL2 = SETSQL2();
            SQL3 = SETSQL3();
            SQL4 = SETSQL4();

            Report report1 = new Report();

            report1.Load(@"REPORT\業務-週報表.frx");

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
            TableDataSource table2 = report1.GetDataSource("Table2") as TableDataSource;
            table2.SelectCommand = SQL3.ToString();
            TableDataSource table3 = report1.GetDataSource("Table3") as TableDataSource;
            table3.SelectCommand = SQL4.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            //string FirstDay = dateTimePicker1.Value.ToString("yyyyMM") + "01";
            //string LastDay = dateTimePicker1.Value.ToString("yyyyMM") + "31";

            string FirstDay = DateTime.Now.AddDays(-7).ToString("yyyyMMdd");
            string LastDay = DateTime.Now.ToString("yyyyMMdd");

            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                             SELECT  [CLIENTNAME]  AS '客戶',[TBSALESEVENTS].[SALES]   AS '業務員',[TBSALESEVENTS].[KINDS]   AS '類別',[TBSALESEVENTS].[CLIENTS]   AS '客戶名',[TBSALESEVENTS].[PROJECTS]   AS '專案',[TBSALESEVENTS].[EVENTS]  AS '待辨',[TBSALESEVENTS].[SDAYS]  AS '開始日',[TBSALESEVENTS].[EDAYS]  AS '結案日',ISNULL([TBSALESEVENTS].[COMMENTS],'本週無記錄') AS '進度',CONVERT(NVARCHAR,[TBSALESEVENTS].[UPDATEDATES],112) AS '更新日期'
                            ,[TB_CLINETS].[ID],[TBSALESEVENTS].[ID]
                            FROM [TKBUSINESS].[dbo].[TB_CLINETS]
                            LEFT JOIN [TKBUSINESS].[dbo].[TBSALESEVENTS] ON [TB_CLINETS].CLIENTNAME=[TBSALESEVENTS].CLIENTS AND CONVERT(NVARCHAR,[TBSALESEVENTS].UPDATEDATES,112)>='{0}'  AND CONVERT(NVARCHAR,[TBSALESEVENTS].UPDATEDATES,112)<='{1}'
                            WHERE 1=1
                            ORDER BY  [TB_CLINETS].[ID],[TBSALESEVENTS].UPDATEDATES

                            ", FirstDay, LastDay);


            return SB;

        }

        public StringBuilder SETSQL2()
        {
            //string FirstDay = dateTimePicker1.Value.ToString("yyyyMM") + "01";
            //string LastDay = dateTimePicker1.Value.ToString("yyyyMM") + "31";

            string FirstDay = DateTime.Now.AddDays(-7).ToString("yyyyMMdd");


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                            SELECT [TBSALESEVENTS].[CLIENTS] AS '客戶',[TBSALESEVENTS].[SALES] AS '業務員',[TBSALESEVENTS].[KINDS]  AS '類別',[TBSALESEVENTS].[PROJECTS]  AS '專案',[TBSALESEVENTS].[EVENTS]  AS '待辨',[TBSALESEVENTS].[SDAYS] AS '開始日',[TBSALESEVENTS].[EDAYS]  AS '結案日',[TBSALESEVENTS].[COMMENTS] AS '進度',CONVERT(NVARCHAR,[TBSALESEVENTS].[UPDATEDATES],112) AS '更新日期'
                            ,[TBSALESEVENTS].[ID]
                            FROM [TKBUSINESS].[dbo].[TBSALESEVENTS]
                            WHERE CONVERT(NVARCHAR,[TBSALESEVENTS].UPDATEDATES,112)>='{0}' 
                            AND [TBSALESEVENTS].[CLIENTS] NOT IN (SELECT  [CLIENTNAME] FROM [TKBUSINESS].[dbo].[TB_CLINETS])
                            ORDER BY [TBSALESEVENTS].[SALES],[TBSALESEVENTS].[CLIENTS],[TBSALESEVENTS].[UPDATEDATES]

                            ", FirstDay);


            return SB;

        }

        public StringBuilder SETSQL3()
        {
            //string FirstDay = dateTimePicker1.Value.ToString("yyyyMM") + "01";
            //string LastDay = dateTimePicker1.Value.ToString("yyyyMM") + "31";

            string FirstDay = DateTime.Now.AddDays(-7).ToString("yyyyMMdd");


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                           SELECT [TBSALESEVENTS].[CLIENTS] AS '客戶',[TBSALESEVENTS].[SALES] AS '業務員',[TBSALESEVENTS].[KINDS]  AS '類別',[TBSALESEVENTS].[PROJECTS]  AS '專案',[TBSALESEVENTS].[EVENTS]  AS '待辨',[TBSALESEVENTS].[SDAYS] AS '開始日',[TBSALESEVENTS].[EDAYS]  AS '結案日',[TBSALESEVENTS].[COMMENTS] AS '進度',CONVERT(NVARCHAR,[TBSALESEVENTS].[UPDATEDATES],112) AS '更新日期'
                            ,[TBSALESEVENTS].[ID]
                            FROM [TKBUSINESS].[dbo].[TBSALESEVENTS]
                            WHERE [TBSALESEVENTS].ISCLOSE='N'
                            AND CONVERT(NVARCHAR,[TBSALESEVENTS].UPDATEDATES,112)<'{0}'  
                            ORDER BY [TBSALESEVENTS].[SALES],[TBSALESEVENTS].[CLIENTS],[TBSALESEVENTS].[UPDATEDATES]

                            ", FirstDay);


            return SB;

        }

        public StringBuilder SETSQL4()
        {
            string FirstDay = dateTimePicker1.Value.ToString("yyyyMM") + "01";
            string LastDay = dateTimePicker1.Value.ToString("yyyyMM") + "31";


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                            SELECT USER_NAME AS '業務員',SALES AS '負責客戶數',COMS AS '拜訪客戶數',NOTES AS '拜訪次數',CONVERT(decimal(16,2),(CONVERT(decimal(16,2),COMS)/CONVERT(decimal(16,2),SALES))*100 ) AS '拜訪客戶完成率%'
                            ,ORDERS,USER_ID,USER_ACCOUNT
                            FROM (
                            SELECT [ORDERS],[USER_ID],[USER_NAME],[USER_ACCOUNT]
                            ,(SELECT COUNT(DISTINCT  [COMPANY_ID]) FROM  [192.168.1.223].[HJ_BM_DB].[dbo].[tb_COMPANY] WHERE [STATUS]='1' AND [COMPANY_NAME] NOT LIKE '%停用%' AND [OWNER_ID]=[USER_ID]) AS 'SALES'
                            ,(SELECT COUNT(DISTINCT  [tb_NOTE].[COMPANY_ID]) FROM  [192.168.1.223].[HJ_BM_DB].[dbo].[tb_NOTE], [192.168.1.223].[HJ_BM_DB].[dbo].[tb_COMPANY] WHERE [STATUS]='1' AND [tb_NOTE].COMPANY_ID=[tb_COMPANY].COMPANY_ID AND [OWNER_ID]=[USER_ID]  AND CONVERT(nvarchar,[tb_NOTE].[CREATE_DATETIME],112)>='{0}'  AND CONVERT(nvarchar,[tb_NOTE].[CREATE_DATETIME],112)<='{1}') AS 'COMS'
                            ,(SELECT COUNT([tb_NOTE].[COMPANY_ID]) FROM  [192.168.1.223].[HJ_BM_DB].[dbo].[tb_NOTE], [192.168.1.223].[HJ_BM_DB].[dbo].[tb_COMPANY] WHERE [STATUS]='1' AND [tb_NOTE].COMPANY_ID=[tb_COMPANY].COMPANY_ID AND [OWNER_ID]=[USER_ID]  AND CONVERT(nvarchar,[tb_NOTE].[CREATE_DATETIME],112)>='{0}'  AND CONVERT(nvarchar,[tb_NOTE].[CREATE_DATETIME],112)<='{1}') AS 'NOTES'
                            FROM  [192.168.1.223].[HJ_BM_DB].[dbo].[COPSALES]
                            ) AS TEMP
                            ORDER BY [ORDERS] 

                            ", FirstDay, LastDay);


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
