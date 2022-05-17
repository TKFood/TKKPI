﻿using System;
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

        #endregion


    }
}
