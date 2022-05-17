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
    public partial class frmREPROTPOSSTORES : Form
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


        public frmREPROTPOSSTORES()
        {
            InitializeComponent();
        }


        #region FUNCTION

        public void SETFASTREPORT4(string SDATE, string EDATE)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL4(SDATE, EDATE);


            Report report1 = new Report();

            report1.Load(@"REPORT\門市-各門市銷售總表.frx");

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

            //組合活動，過濾會員等級折扣 ISNULL(MI009,'')=''

            SB.AppendFormat(@"   
                            --20220506 門市銷售資料

                            SELECT  CASE WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=1 THEN '星期一' WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=2 THEN '星期二'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=3 THEN '星期三'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=4 THEN '星期四'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=5 THEN '星期五'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=6 THEN '星期六'WHEN DATEPART(WEEKDAY, CONVERT(DATETIME,TA001)-1)=7 THEN '星期日' END AS '星期'
                            ,TA001 AS '日期',MA002 AS '賣場',TA002 AS '賣場代號',SUM(未稅金額) 總未稅金額
                            ,(SELECT ISNULL([NAMES],'')+CHAR(10) FROM [TKKPI].[dbo].[SALESPROJECTSSTORES] WITH (NOLOCK) WHERE SDATES<=TA001 AND EDATES>=TA001 FOR XML PATH('')) AS '調整事項'
                            ,(SELECT ISNULL([MB004],'')+CHAR(10) FROM [TK].dbo.POSMB  WITH (NOLOCK) LEFT JOIN [TK].dbo.POSMF WITH (NOLOCK) ON MF003=MB003 WHERE MB012<=TA001 AND MB013>=TA001  AND (ISNULL(MF004,'')='' OR MF004 IN (TA002))  FOR XML PATH('')) AS 'POS活動'
                            ,(SELECT ISNULL([MI004],'')+CHAR(10) FROM [TK].dbo.POSMI  WITH (NOLOCK) LEFT JOIN [TK].dbo.POSMF WITH (NOLOCK) ON MF003=MI003  WHERE MI005<=TA001 AND MI006>=TA001 AND ISNULL(MI009,'')='' AND (ISNULL(MF004,'')='' OR MF004 IN (TA002))  FOR XML PATH('')) AS '組合活動'
                            ,(SELECT ISNULL([MM004],'')+CHAR(10) FROM [TK].dbo.POSMM  WITH (NOLOCK) LEFT JOIN [TK].dbo.POSMF WITH (NOLOCK) ON MF003=MM003  WHERE MM005<=TA001 AND MM006>=TA001 AND (ISNULL(MF004,'')='' OR MF004 IN (TA002))  FOR XML PATH('')) AS '贈品加價購活動'
                            ,(SELECT ISNULL([MO004],'')+CHAR(10) FROM [TK].dbo.POSMO  WITH (NOLOCK) LEFT JOIN [TK].dbo.POSMF WITH (NOLOCK) ON MF003=MO003  WHERE MO005<=TA001 AND MO006>=TA001 AND (ISNULL(MF004,'')='' OR MF004 IN (TA002))  FOR XML PATH('')) AS '配對搭贈活動'

                            FROM 
                            (
                            SELECT TA001,TA002
                            ,(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTB TB WITH (NOLOCK) WHERE POSTA.TA001=TB.TB001 AND POSTA.TA002=TB.TB002 AND TB.TB042 NOT IN ('4') ) AS '未稅金額'
                            ,(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTA TA WITH (NOLOCK),[TK].dbo.POSTB TB WITH (NOLOCK) WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TA.TA003=TB.TB003 AND TA.TA006=TB.TB006 AND POSTA.TA001=TB.TB001 AND POSTA.TA002=TB.TB002 AND TB.TB042 NOT IN ('4') AND TA009 LIKE '68%') AS '團客未稅金額'
                            FROM [TK].dbo.POSTA WITH (NOLOCK)
                            WHERE 1=1
                            AND TA002 IN ('106501','106502','106503','106504')
                            AND TA001>='{0}' AND TA001<='{1}'
                            GROUP BY TA001,TA002

                            ) AS TEMP
                            LEFT JOIN [TK].dbo.WSCMA  WITH (NOLOCK) ON MA001=TA002
                            GROUP BY MA002,TA001,TA002
                            ORDER BY MA002,TA001,TA002
 
 

                            ", SDATE, EDATE);


            return SB;

        }

        public void Search(string SYEARS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            DataSet ds = new DataSet();

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

                talbename = "TEMPds1";
                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                        SELECT 
                                        [NAMES] AS '調整事項'
                                        ,[SDATES] AS ' 開始日'
                                        ,[EDATES] AS '結束日'

                                        FROM [TKKPI].[dbo].[SALESPROJECTSSTORES]
                                        WHERE 1=1
                                        AND SDATES LIKE '{0}%'
                                        ORDER BY SDATES

                                         ", SYEARS);



                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, talbename);
                sqlConn.Close();


                if (ds.Tables[talbename].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    dataGridView1.DataSource = ds.Tables[talbename];
                    dataGridView1.AutoResizeColumns();
                    //rownum = ds.Tables[talbename].Rows.Count - 1;
                    dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

                    //dataGridView1.CurrentCell = dataGridView1[0, 2];

                }


            }
            catch
            {

            }
            finally
            {

            }

        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            string MNAME = null;
            textBox3.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    MNAME = row.Cells["調整事項"].Value.ToString();
                    textBox3.Text = row.Cells["調整事項"].Value.ToString();

                    SETFASTREPORT5(MNAME);


                }
                else
                {


                }
            }
        }

        public void SETFASTREPORT5(string MNAMES)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL5(MNAMES);


            Report report1 = new Report();

            report1.Load(@"REPORT\門市-調整事項品號.frx");

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


            report1.Preview = previewControl5;
            report1.Show();
        }

        public StringBuilder SETSQL5(string MNAMES)
        {


            StringBuilder SB = new StringBuilder();

            //組合活動，過濾會員等級折扣 ISNULL(MI009,'')=''

            SB.AppendFormat(@"   
                            SELECT 
                            MB001 AS '品號'
                            ,MB002 AS '品名'
                            ,[NAMES] AS '調整事項'
                            ,[SDATES] AS '開始日'
                            ,[EDATES] AS '結束日'

                            FROM [TKKPI].dbo.[SALESPROJECTSSTORES],[TKKPI].dbo.[SALESPROJECTSINVMBSTORES]
                            WHERE 1=1
                            AND SALESPROJECTSSTORES.NAMES=SALESPROJECTSINVMBSTORES.MNAMES
                            AND MNAMES ='{0}'
                            ORDER BY MB001

                            ", MNAMES);


            return SB;

        }

        public void SETFASTREPORT6(string SDATES, string EDATES, string MNAME)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL6(SDATES, EDATES, MNAME);


            Report report1 = new Report();

            report1.Load(@"REPORT\門市-調整事項每日總金額.frx");

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


            report1.Preview = previewControl6;
            report1.Show();
        }

        public StringBuilder SETSQL6(string SDATES, string EDATES, string MNAME)
        {


            StringBuilder SB = new StringBuilder();

            //組合活動，過濾會員等級折扣 ISNULL(MI009,'')=''

            SB.AppendFormat(@"   
                          
                           SELECT ISNULL(KINDS,'') AS '調整事項',TA001 AS '銷售日',TA002 AS '賣場代',MA002 AS '賣場',總未稅金額
                            FROM (
                            SELECT NAMES AS 'KINDS',TA001,TA002
                            ,(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTB TB WITH(NOLOCK) WHERE POSTA.TA001=TB.TB001 AND POSTA.TA002=TB.TB002 AND TB.TB010 IN (SELECT MB001 FROM [TKKPI].dbo.SALESPROJECTSINVMBSTORES WHERE MNAMES='{2}' ) AND TB.TB042 NOT IN ('4') ) AS '總未稅金額'
                            FROM [TK].dbo.POSTA WITH(NOLOCK)
                            LEFT JOIN [TKKPI].dbo.SALESPROJECTSSTORES ON NAMES='{2}' AND  SDATES<=TA001 AND EDATES>=TA001
                            WHERE 1=1
                            AND TA002 IN ('106501','106502','106503','106504')
                            AND TA001>='{0}' AND TA001<='{1}'
                            GROUP BY TA001,TA002,NAMES
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.WSCMA ON MA001=TA002
                            ORDER BY TA001,TA002

                            ", SDATES, EDATES, MNAME);


            return SB;

        }

        public void SearchPOS(string SYEARS)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            DataSet ds = new DataSet();

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

                talbename = "TEMPds1";
                sbSql.Clear();

                sbSql.AppendFormat(@"  
                                       
                                    SELECT '活動特價' AS '類型',MB004 AS '活動名稱',MB012 AS '開始日',MB013 AS '結束日',MF004 AS '適用門市',MB003 AS '活動代號'
                                    FROM [TK].dbo.POSMB
                                    LEFT JOIN [TK].dbo.POSMF ON MF003=MB003
                                    WHERE 1=1
                                    AND MB008='Y'
                                    AND MB013 LIKE '{0}%'
                                    AND MF004 IN ('106501','106502','106503','106504')
                                    UNION ALL
                                    SELECT  '組合品搭贈' AS KIND,MI004,MI005,MI006,MF004,MI003
                                    FROM [TK].dbo.POSMI
                                    LEFT JOIN [TK].dbo.POSMF ON MF003=MI003
                                    WHERE 1=1
                                    AND MI015='Y'
                                    AND MI005 LIKE  '{0}%'
                                    AND MF004 IN ('106501','106502','106503','106504')
                                    UNION ALL
                                    SELECT  '滿額折價' AS KIND,MM004,MM005,MM006,MM004,MM003
                                    FROM [TK].dbo.POSMM
                                    LEFT JOIN [TK].dbo.POSMF ON MF003=MM003
                                    WHERE 1=1
                                    AND MM015='Y'
                                    AND MM005 LIKE  '{0}%'
                                    AND MF004 IN ('106501','106502','106503','106504')
                                    UNION ALL
                                    SELECT  '配對搭贈' AS KIND,MO004,MO005,MO006,MF004,MO003
                                    FROM [TK].dbo.POSMO
                                    LEFT JOIN [TK].dbo.POSMF ON MF003=MO003
                                    WHERE 1=1
                                    AND MO008='Y'
                                    AND MO005 LIKE  '{0}%'
                                    AND MF004 IN ('106501','106502','106503','106504')


                                    ", SYEARS);



                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, talbename);
                sqlConn.Close();


                if (ds.Tables[talbename].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    dataGridView2.DataSource = ds.Tables[talbename];
                    dataGridView2.AutoResizeColumns();
                    //rownum = ds.Tables[talbename].Rows.Count - 1;
                    dataGridView2.CurrentCell = dataGridView2.Rows[rownum].Cells[0];

                    //dataGridView1.CurrentCell = dataGridView1[0, 2];

                }


            }
            catch
            {

            }
            finally
            {

            }

        }
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            string MNAME = null;
            string ID = null;
            textBox4.Text = null;

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    MNAME = row.Cells["活動名稱"].Value.ToString();
                    ID = row.Cells["活動代號"].Value.ToString().Trim();

                    textBox4.Text = row.Cells["活動名稱"].Value.ToString();
                    textBox5.Text = row.Cells["活動代號"].Value.ToString();

                    SETFASTREPORT8(ID);


                }
                else
                {


                }
            }
        }

        public void SETFASTREPORT8(string ID)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL8(ID);


            Report report1 = new Report();

            report1.Load(@"REPORT\門市-POS活動品號.frx");

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


            report1.Preview = previewControl8;
            report1.Show();
        }

        public StringBuilder SETSQL8(string ID)
        {


            StringBuilder SB = new StringBuilder();

            //組合活動，過濾會員等級折扣 ISNULL(MI009,'')=''

            SB.AppendFormat(@"   
                            SELECT MC004,INVMB.MB002,POSMB.MB012,POSMB.MB013,POSMB.MB004
                            FROM [TK].dbo.POSMC,[TK].dbo.INVMB,[TK].dbo.POSMB
                            WHERE 1=1
                            AND MC004=INVMB.MB001
                            AND POSMB.MB003=MC003
                            AND MC011='Y'
                            AND MC003='{0}'
                            UNION ALL
                            SELECT MJ004,MB002,MI005,MI006,MI004
                            FROM [TK].dbo.POSMJ,[TK].dbo.INVMB,[TK].dbo.POSMI
                            WHERE 1=1
                            AND MJ004=MB001
                            AND MI003=MJ003
                            AND MJ006='Y'
                            AND MJ003='{0}'
                            UNION ALL
                            SELECT CONVERT(NVARCHAR,MN005),'金額以上',MM005,MM006,MM004
                            FROM [TK].dbo.POSMN,[TK].dbo.POSMM
                            WHERE 1=1
                            AND MN003=MM003
                            AND MN010='Y'
                            AND MN003='{0}'
                            UNION ALL
                            SELECT MP005,MB002,MO011,MO012,MO004
                            FROM [TK].dbo.POSMP,[TK].dbo.INVMB,[TK].dbo.POSMO
                            WHERE 1=1
                            AND MP005=MB001
                            AND MP003=MO003
                            AND MP008='Y'
                            AND MP003='{0}'


                            ", ID);


            return SB;

        }
        public void SETFASTREPORT9(string SDATES, string EDATES, string ID)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL9(SDATES, EDATES, ID);


            Report report1 = new Report();

            report1.Load(@"REPORT\門市-POS活動每日總金額.frx");

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


            report1.Preview = previewControl9;
            report1.Show();
        }

        public StringBuilder SETSQL9(string SDATES, string EDATES, string ID)
        {


            StringBuilder SB = new StringBuilder();

            //組合活動，過濾會員等級折扣 ISNULL(MI009,'')=''

            SB.AppendFormat(@"   

                            SELECT ISNULL(KINDS,'') AS 'POS活動',TA001 AS '銷售日',TA002 AS '賣場代',MA002 AS '賣場',總未稅金額
                            FROM (
                            SELECT MB004 AS 'KINDS',TA001,TA002
                            ,(SELECT ISNULL(SUM(TB031),0) FROM [TK].dbo.POSTB TB WITH(NOLOCK) WHERE POSTA.TA001=TB.TB001 AND POSTA.TA002=TB.TB002 AND TB.TB010 IN (SELECT MC004 FROM [TK].dbo.POSMC WHERE MC003='{2}') AND TB.TB042 NOT IN ('4') ) AS '總未稅金額'
                            
                            FROM [TK].dbo.POSTA WITH(NOLOCK)
                            LEFT JOIN [TK].dbo.POSMB ON MB003='{2}' AND  MB012<=TA001 AND MB013>=TA001
                            LEFT JOIN [TK].dbo.POSMI ON MI003='{2}' AND  MI005<=TA001 AND MI006>=TA001
                            LEFT JOIN [TK].dbo.POSMM ON MM003='{2}' AND  MM005<=TA001 AND MM006>=TA001
                            LEFT JOIN [TK].dbo.POSMO ON MO003='{2}' AND  MO005<=TA001 AND MO006>=TA001
 
                            WHERE 1=1
                            AND TA002 IN ('106501','106502','106503','106504')
                            AND TA001>='{0}' AND TA001<='{1}'
                            GROUP BY TA001,TA002,MB004
                            ) AS TEMP
                            LEFT JOIN [TK].dbo.WSCMA ON MA001=TA002
                            ORDER BY TA001,TA002
                           

                            ", SDATES, EDATES, ID);


            return SB;

        }
        public void SETFASTREPORT10(string SDATES, string EDATES, string ID)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL10(SDATES, EDATES, ID);


            Report report1 = new Report();

            report1.Load(@"REPORT\門市-POS活動每日明細.frx");

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


            report1.Preview = previewControl10;
            report1.Show();
        }

        public StringBuilder SETSQL10(string SDATES, string EDATES, string ID)
        {


            StringBuilder SB = new StringBuilder();

            //組合活動，過濾會員等級折扣 ISNULL(MI009,'')=''

            SB.AppendFormat(@"    
                            SELECT TEMP.MB003,TEMP.MB004  AS 'POS活動',TEMP.MB012 AS '開始日',TEMP.MB013 AS '結束日',TA001 AS '銷售日',TA002 AS '賣場代',MA002  AS '賣場',MC004 AS '品號',INVMB.MB002 AS '品名'
                            ,總未稅金額
                            ,團客金額
                            ,(總未稅金額-團客金額) AS 散客金額
                            FROM
                            (
                            SELECT 
                            MB003,MB004,TA001,TA002,MC004,POSMB.MB012,POSMB.MB013
                            ,(SELECT ISNULL(SUM(TB031),0) FROM  [TK].dbo.POSTA TA WITH(NOLOCK),[TK].dbo.POSTB TB WITH(NOLOCK) WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TA.TA003=TB.TB003 AND TA.TA006=TB.TB006 AND TA.TA001=POSTA.TA001 AND TA.TA002=POSTA.TA002 AND TB.TB010=MC004 AND TB.TB042 NOT IN ('4') ) AS '總未稅金額'
                            ,(SELECT ISNULL(SUM(TB031),0) FROM  [TK].dbo.POSTA TA WITH(NOLOCK),[TK].dbo.POSTB TB WITH(NOLOCK) WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TA.TA003=TB.TB003 AND TA.TA006=TB.TB006 AND TA.TA009 LIKE '68%' AND TA.TA001=POSTA.TA001 AND TA.TA002=POSTA.TA002 AND TB.TB010=MC004 AND TB.TB042 NOT IN ('4') ) AS '團客金額'

                            FROM [TK].dbo.POSMB,[TK].dbo.POSMC,[TK].dbo.POSTA WITH(NOLOCK) 
                            WHERE 1=1
                            AND MB003=MC003
                            AND MC003='{2}'
                            AND TA002 IN ('106501','106502','106503','106504')
                            AND TA001>='{0}' AND TA001<='{1}' 
                            GROUP BY 
                            MB003,MB004,TA001,TA002,MC004,POSMB.MB012,POSMB.MB013

                            UNION ALL
                            SELECT 
                            MI003,MI004,TA001,TA002,MJ004,MI005,MI006
                            ,(SELECT ISNULL(SUM(TB031),0) FROM  [TK].dbo.POSTA TA WITH(NOLOCK),[TK].dbo.POSTB TB WITH(NOLOCK) WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TA.TA003=TB.TB003 AND TA.TA006=TB.TB006 AND TA.TA001=POSTA.TA001 AND TA.TA002=POSTA.TA002 AND TB.TB010=MJ004) AS '總未稅金額'
                            ,(SELECT ISNULL(SUM(TB031),0) FROM  [TK].dbo.POSTA TA WITH(NOLOCK),[TK].dbo.POSTB TB WITH(NOLOCK) WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TA.TA003=TB.TB003 AND TA.TA006=TB.TB006 AND TA.TA009 LIKE '68%' AND TA.TA001=POSTA.TA001 AND TA.TA002=POSTA.TA002 AND TB.TB010=MJ004) AS '團客金額'

                            FROM [TK].dbo.POSMI,[TK].dbo.POSMJ,[TK].dbo.POSTA WITH(NOLOCK) 
                            WHERE 1=1
                            AND MI003=MJ003
                            AND MI003='{2}'
                            AND TA002 IN ('106501','106502','106503','106504')
                            AND TA001>='{0}' AND TA001<='{1}' 
                            GROUP BY 
                            MI003,MI004,TA001,TA002,MJ004,MI005,MI006
                            UNION ALL

                            SELECT 
                            MO003,MO004,TA001,TA002,MP005,MO011,MO012
                            ,(SELECT ISNULL(SUM(TB031),0) FROM  [TK].dbo.POSTA TA WITH(NOLOCK),[TK].dbo.POSTB TB WITH(NOLOCK) WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TA.TA003=TB.TB003 AND TA.TA006=TB.TB006 AND TA.TA001=POSTA.TA001 AND TA.TA002=POSTA.TA002 AND TB.TB010=MP005) AS '總未稅金額'
                            ,(SELECT ISNULL(SUM(TB031),0) FROM  [TK].dbo.POSTA TA WITH(NOLOCK),[TK].dbo.POSTB TB WITH(NOLOCK) WHERE TA.TA001=TB.TB001 AND TA.TA002=TB.TB002 AND TA.TA003=TB.TB003 AND TA.TA006=TB.TB006 AND TA.TA009 LIKE '68%' AND TA.TA001=POSTA.TA001 AND TA.TA002=POSTA.TA002 AND TB.TB010=MP005) AS '團客金額'

                            FROM [TK].dbo.POSMO,[TK].dbo.POSMP,[TK].dbo.POSTA WITH(NOLOCK) 
                            WHERE 1=1
                            AND MO003=MP003
                            AND MO003='{2}'
                            AND TA002 IN ('106501','106502','106503','106504')
                            AND TA001>='{0}' AND TA001<='{1}' 
                            GROUP BY 
                            MO003,MO004,TA001,TA002,MP005,MO011,MO012


                            ) AS TEMP
                            LEFT JOIN [TK].dbo.WSCMA ON MA001=TA002
                            LEFT JOIN [TK].dbo.INVMB ON MB001=MC004
                            WHERE 總未稅金額>0





                            ", SDATES, EDATES, ID);


            return SB;

        }

        public void SETFASTREPORT(string YEARS)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL(YEARS);


            Report report1 = new Report();

            report1.Load(@"REPORT\門市-特價報表.frx");

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

            report1.Load(@"REPORT\門市-特價商品銷售.frx");

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

        public void SETFASTREPORT3(string SDATE, string EDATE, string TB036)
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3(SDATE, EDATE, TB036);


            Report report1 = new Report();

            report1.Load(@"REPORT\門市-活動報表.frx");

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

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            SETFASTREPORT4(dateTimePicker4.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMMdd"));
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker6.Value.ToString("yyyy"));
        }
        private void button6_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox3.Text.ToString().Trim()))
            {
                SETFASTREPORT6(dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"), textBox3.Text.ToString().Trim());
                //SETFASTREPORT7(dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"), textBox3.Text.ToString().Trim());
            }
        }
        private void button7_Click(object sender, EventArgs e)
        {
            SearchPOS(dateTimePicker6.Value.ToString("yyyy"));
        }
        private void button8_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox5.Text.ToString().Trim()))
            {
                SETFASTREPORT9(dateTimePicker10.Value.ToString("yyyyMMdd"), dateTimePicker11.Value.ToString("yyyyMMdd"), textBox5.Text.ToString().Trim());
                SETFASTREPORT10(dateTimePicker10.Value.ToString("yyyyMMdd"), dateTimePicker11.Value.ToString("yyyyMMdd"), textBox5.Text.ToString().Trim());
            }
        }

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
            SETFASTREPORT3(dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"), textBox2.Text.ToString().Trim());
        }

        #endregion


    }
}
