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
    public partial class frmREPROTPOS : Form
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


        public frmREPROTPOS()
        {
            InitializeComponent();

            dateTimePicker1.Value = DateTime.Now;
        }

        #region FUNCTION
        public void SETFASTREPORT(string YEARS)
        {
            StringBuilder SQL1 = new StringBuilder();    

            SQL1 = SETSQL(YEARS);
    

            Report report1 = new Report();

            report1.Load(@"REPORT\營銷-特價報表.frx");

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

        public StringBuilder SETSQL(string YEARS)
        {
    
          
            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                           SELECT KIND,活動代號,特價代號,活動名稱,活動起始日期,活動截止日期
                            ,(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTB WHERE TB036=特價代號) AS '未稅金額'
                            ,(SELECT ISNULL(SUM(TB019),0) FROM [TK].dbo.POSTB WHERE TB036=特價代號) AS '數量'
                            FROM (
                            SELECT '特價' AS 'KIND', MB001 AS '活動代號',MB003 AS '特價代號',MB004 AS '活動名稱',MB012 AS '活動起始日期',MB013 AS '活動截止日期'
                            FROM [TK].dbo.POSMB
                            WHERE MB001 LIKE '{0}%'
                            UNION ALL
                            SELECT '組合品搭贈' AS 'KIND',MI001 AS '活動代號',MI003 AS '特價代號',MI004 AS '活動名稱',MI005 AS '活動起始日期',MI006 AS '活動截止日期'
                            FROM [TK].dbo.POSMI
                            WHERE MI001 LIKE '{0}%'
                            UNION ALL
                            SELECT '滿額折價' AS 'KIND',MM001 AS '活動代號',MM003 AS '特價代號',MM004 AS '活動名稱',MM005 AS '活動起始日期',MM006 AS '活動截止日期'
                            FROM [TK].dbo.POSMM
                            WHERE MM001 LIKE '{0}%'
                            UNION ALL
                            SELECT '配對搭贈' AS 'KIND',MO001 AS '活動代號',MO003 AS '特價代號',MO004 AS '活動名稱',MO005 AS '活動起始日期',MO006 AS '活動截止日期'
                            FROM [TK].dbo.POSMO
                            WHERE MO001 LIKE '{0}%'
                            ) AS TEMP



                            ", YEARS);


            return SB;

        }


        public void SETFASTREPORT2(string TB036)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2(TB036);


            Report report1 = new Report();

            report1.Load(@"REPORT\營銷-特價商品銷售.frx");

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

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL2(string TB036)
        {


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                            SELECT (ISNULL(POSMB.MB003,'')+ISNULL(MI003,'')+ISNULL(MM003,'')+ISNULL(MO003,'') ) AS '特價代號',(ISNULL(POSMB.MB004,'')+ISNULL(MI004,'')+ISNULL(MM004,'')+ISNULL(MO004,''))  AS '特價名稱',TB002  AS '店代',MA002  AS '店名',TB010  AS '品號',INVMB.MB002  AS '品名',SUM(TB019) AS '數量',SUM(TB031) AS '未稅金額'
                            FROM [TK].dbo.INVMB,[TK].dbo.WSCMA,[TK].dbo.POSTB
                            LEFT JOIN [TK].dbo.POSMB ON MB003=TB036
                            LEFT JOIN [TK].dbo.POSMI ON MI003=TB036
                            LEFT JOIN [TK].dbo.POSMM ON MM003=TB036
                            LEFT JOIN [TK].dbo.POSMO ON MO003=TB036
                            WHERE TB010=INVMB.MB001
                            AND MA001=TB002
                            AND TB036 LIKE '%{0}%'
                            GROUP BY (ISNULL(POSMB.MB003,'')+ISNULL(MI003,'')+ISNULL(MM003,'')+ISNULL(MO003,'') ) ,(ISNULL(POSMB.MB004,'')+ISNULL(MI004,'')+ISNULL(MM004,'')+ISNULL(MO004,'')),TB002,MA002,TB010,INVMB.MB002

 

                            ", TB036);


            return SB;

        }

        public void SETFASTREPORT3(string SDATE,string EDATE,string TB036)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3(SDATE, EDATE, TB036);


            Report report1 = new Report();

            report1.Load(@"REPORT\營銷活動報表.frx");

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

            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL3(string SDATE, string EDATE, string TB036)
        {


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                            SELECT TB004 AS '營業日期',TB002 AS '店代',MA002 AS '店名',TB003 AS '機號',TB010 AS '品號',MB002 AS '品名',TB019 AS '數量',TB031 AS '未稅金額',TB044 AS '備註'
                            ,商品+組合+贈品加價購+配對搭贈 AS '活動代號'
                            ,商品名稱+組合名稱+贈品加價購名稱+配對搭贈名稱  AS '活動名稱'
                            FROM 
                            (
                            SELECT POSTB.*
                            ,ISNULL(RTRIM(LTRIM(MB003)),'') AS '商品',ISNULL(MB004,'') AS '商品名稱'
                            ,ISNULL(RTRIM(LTRIM(MI003)),'') AS '組合',ISNULL(MI004,'') AS '組合名稱'
                            ,ISNULL(RTRIM(LTRIM(MM003)),'') AS '贈品加價購',ISNULL(MM004,'') AS '贈品加價購名稱'
                            ,ISNULL(RTRIM(LTRIM(MO003)),'') AS '配對搭贈',ISNULL(MO004,'') AS '配對搭贈名稱'
                            FROM [TK].dbo.POSTB WITH (NOLOCK)
                            LEFT JOIN [TK].dbo.POSMB ON MB003=TB036
                            LEFT JOIN [TK].dbo.POSMI ON MI003=TB036
                            LEFT JOIN [TK].dbo.POSMM ON MM003=TB036
                            LEFT JOIN [TK].dbo.POSMO ON MO003=TB036

                            WHERE 1=1
                            AND ISNULL(TB044,'')<>''
                            ) AS TEMMP
                            LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                            LEFT JOIN [TK].dbo.WSCMA ON MA001=TB002
                            WHERE 1=1
                            AND ISNULL(商品+組合+贈品加價購+配對搭贈,'')<>''
                            AND TB001>='{0}' AND TB001<='{1}'
                            AND (ISNULL(商品+組合+贈品加價購+配對搭贈,'') LIKE '%{2}%' OR 商品名稱+組合名稱+贈品加價購名稱+配對搭贈名稱 LIKE '%{2}%')
                            ORDER BY TB002,TB004 
                            ", SDATE, EDATE, TB036);


            return SB;

        }

        public void SETFASTREPORT4(string SDATE, string EDATE)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL4(SDATE, EDATE);


            Report report1 = new Report();

            report1.Load(@"REPORT\營銷-觀光賣場總表.frx");

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


            //report1.SetParameterValue("P1", SDATE);
            //report1.SetParameterValue("P2", EDATE);


            report1.Preview = previewControl4;
            report1.Show();
        }

        public StringBuilder SETSQL4(string SDATE, string EDATE)
        {


            StringBuilder SB = new StringBuilder();


            SB.AppendFormat(@"   
                            SELECT  CASE WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=1 THEN '星期一' WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=2 THEN '星期二'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=3 THEN '星期三'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=4 THEN '星期四'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=5 THEN '星期五'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=6 THEN '星期六'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=7 THEN '星期日' END AS '星期'
                            ,TA001 AS '日期',MA002 AS '賣場',TA002 AS '賣場代號',SUM(未稅金額) 總未稅金額,SUM(團客未稅金額) 團客未稅金額,(SUM(未稅金額)-SUM(團客未稅金額)) 散客未稅金額
                            ,(SELECT ISNULL(SUM([CARNUM]),0) FROM [TKMK].[dbo].[GROUPSALES] WITH (NOLOCK) WHERE [STATUS]='完成接團' AND CONVERT(NVARCHAR,[CREATEDATES],112)=TA001) AS '來車數'
                            ,(SELECT ISNULL([NAMES],'')+CHAR(10) FROM [TKKPI].[dbo].[SALESPROJECTS] WITH (NOLOCK) WHERE SDATES<=TA001 AND EDATES>=TA001 FOR XML PATH('')) AS '調整事項'
                            ,(SELECT ISNULL([MB004],'')+CHAR(10) FROM [TK].dbo.POSMB  WITH (NOLOCK) WHERE MB012<=TA001 AND MB013>=TA001  FOR XML PATH('')) AS 'POS活動'
                            ,(SELECT ISNULL([MI004],'')+CHAR(10) FROM [TK].dbo.POSMI  WITH (NOLOCK) WHERE MI005<=TA001 AND MI006>=TA001  FOR XML PATH('')) AS '組合活動'
                            ,(SELECT ISNULL([MM004],'')+CHAR(10) FROM [TK].dbo.POSMM  WITH (NOLOCK) WHERE MM005<=TA001 AND MM006>=TA001  FOR XML PATH('')) AS '贈品加價購活動'
                            ,(SELECT ISNULL([MO003],'')+CHAR(10) FROM [TK].dbo.POSMO  WITH (NOLOCK) WHERE MO005<=TA001 AND MO006>=TA001  FOR XML PATH('')) AS '配對搭贈活動'

                            FROM 
                            (
                            SELECT TA001,TA002
                            ,(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTB TB WITH (NOLOCK) WHERE POSTA.TA001=TB.TB001 AND POSTA.TA002=TB.TB002) AS '未稅金額'
                            ,(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTA TA WITH (NOLOCK),[TK].dbo.POSTB TB WITH (NOLOCK) WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TA.TA003=TB.TB003 AND TA.TA006=TB.TB006 AND POSTA.TA001=TB.TB001 AND POSTA.TA002=TB.TB002 AND TA009 LIKE '68%') AS '團客未稅金額'
                            FROM [TK].dbo.POSTA WITH (NOLOCK)
                            WHERE 1=1
                            AND TA002 IN ('106701')
                            AND TA001>='{0}' AND TA001<='{1}'
                            GROUP BY TA001,TA002

                            ) AS TEMP
                            LEFT JOIN [TK].dbo.WSCMA  WITH (NOLOCK) ON MA001=TA002
                            GROUP BY TA001,TA002,MA002
                            ORDER BY TA001,TA002,MA002
 

                            ", SDATE, EDATE);


            return SB;

        }

        public void Search(string SDATE,string EDATE)
        {
            //try
            //{
            //    //20210902密
            //    Class1 TKID = new Class1();//用new 建立類別實體
            //    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //    //資料庫使用者密碼解密
            //    sqlsb.Password = TKID.Decryption(sqlsb.Password);
            //    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            //    String connectionString;
            //    sqlConn = new SqlConnection(sqlsb.ConnectionString);
               
            //    StringBuilder sbSql = new StringBuilder();


            //    sbSql.AppendFormat(@"   
                            
            //                        SELECT  CASE WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=1 THEN '星期一' WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=2 THEN '星期二'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=3 THEN '星期三'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=4 THEN '星期四'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=5 THEN '星期五'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=6 THEN '星期六'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=7 THEN '星期日' END AS '星期'
            //                        ,TA001 AS '日期',MA002 AS '賣場',TA002 AS '賣場代號',SUM(未稅金額) 總未稅金額,SUM(團客未稅金額) 團客未稅金額,(SUM(未稅金額)-SUM(團客未稅金額)) 散客未稅金額
            //                        ,(SELECT ISNULL(SUM([CARNUM]),0) FROM [TKMK].[dbo].[GROUPSALES] WITH (NOLOCK) WHERE [STATUS]='完成接團' AND CONVERT(NVARCHAR,[CREATEDATES],112)=TA001) AS '來車數'
            //                        ,(SELECT ISNULL([NAMES],'')+CHAR(10) FROM [TKKPI].[dbo].[SALESPROJECTS] WITH (NOLOCK) WHERE SDATES<=TA001 AND EDATES>=TA001 FOR XML PATH('')) AS '調整事項'
            //                        ,(SELECT ISNULL([MB004],'')+CHAR(10) FROM [TK].dbo.POSMB  WITH (NOLOCK) WHERE MB012<=TA001 AND MB013>=TA001  FOR XML PATH('')) AS 'POS活動'
            //                        FROM 
            //                        (
            //                        SELECT TA001,TA002,TA003,TA006
            //                        ,(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTB TB WITH (NOLOCK) WHERE POSTA.TA001=TB.TB001 AND POSTA.TA002=TB.TB002 AND POSTA.TA003=TB.TB003 AND POSTA.TA006=TB.TB006) AS '未稅金額'
            //                        ,(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTA TA WITH (NOLOCK),[TK].dbo.POSTB TB WITH (NOLOCK) WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TA.TA003=TB.TB003 AND TA.TA006=TB.TB006 AND POSTA.TA001=TB.TB001 AND POSTA.TA002=TB.TB002 AND POSTA.TA003=TB.TB003 AND POSTA.TA006=TB.TB006 AND TA009 LIKE '68%') AS '團客未稅金額'
            //                        FROM [TK].dbo.POSTA WITH (NOLOCK)
            //                        WHERE 1=1
            //                        AND TA002 IN ('106701')
            //                        AND TA001>='{0}' AND TA001<='{1}'

            //                        ) AS TEMP
            //                        LEFT JOIN [TK].dbo.WSCMA  WITH (NOLOCK) ON MA001=TA002
            //                        GROUP BY TA001,TA002,MA002
            //                        ORDER BY TA001,TA002,MA002

            //                        ", SDATE, EDATE);

            //    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
            //    sqlCmdBuilder = new SqlCommandBuilder(adapter);

            //    sqlConn.Open();
            //    ds.Clear();
            //    adapter.Fill(ds, "ds");
            //    sqlConn.Close();


            //    if (ds.Tables["ds"].Rows.Count == 0)
            //    {
            //        dataGridView1.DataSource = null;
            //    }
            //    else
            //    {
            //        dataGridView1.DataSource = ds.Tables["ds"];
                 
            //        //rownum = ds.Tables[talbename].Rows.Count - 1;

            //        //依內容自動換行
            //        dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            //        //自動調欄寬
            //        dataGridView1.AutoResizeColumns();
            //        //自動調欄高
            //        dataGridView1.AutoResizeRows();
            //        //設定顯示格式-數字
            //        dataGridView1.Columns["總未稅金額"].DefaultCellStyle.Format = "N0";
            //        dataGridView1.Columns["團客未稅金額"].DefaultCellStyle.Format = "N0";
            //        dataGridView1.Columns["散客未稅金額"].DefaultCellStyle.Format = "N0";

            //        dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

            //        //dataGridView1.CurrentCell = dataGridView1[0, 2];

            //    }



            //}
            //catch
            //{

            //}
            //finally
            //{

            //}

        }

        #endregion

        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(dateTimePicker1.Value.ToString("yyyy"));
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETFASTREPORT2(textBox1.Text.ToString().Trim());
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT3(dateTimePicker2.Value.ToString("yyyyMMdd"),dateTimePicker3.Value.ToString("yyyyMMdd"), textBox2.Text.ToString().Trim());
        }
        private void button4_Click(object sender, EventArgs e)
        {
            //Search(dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));

            SETFASTREPORT4(dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
        }

        #endregion


    }
}
