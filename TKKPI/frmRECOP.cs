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
using TKITDLL;

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

            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox4load();
        }

        #region FUNCTION

        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [ID],[ID]+[NAME] AS NAME FROM [TKKPI].[dbo].[COPDEP] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void comboBox2load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[ID]+[NAME]  AS NAME FROM [TKKPI].[dbo].[COPSALES] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "ID";
            comboBox2.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void comboBox3load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT [ID],[ID]+[NAME] AS NAME FROM [TKKPI].[dbo].[COPDEP] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "ID";
            comboBox3.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void comboBox4load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[ID]+[NAME]  AS NAME FROM [TKKPI].[dbo].[COPSALES] ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "ID";
            comboBox4.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\業務商品排名表.frx");

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
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT MV002 AS '業務',TH005 AS '品名',SUM(TH037) AS '金額',SUM(LA011) AS '數量',MB004 AS '單位',SUM(TH037)/SUM(SUM(TH037)) OVER ()  AS '金額百分比'");
            SB.AppendFormat(" FROM [TK].dbo.COPTG, [TK].dbo.COPTH,[TK].dbo.INVLA,[TK].dbo.INVMB,[TK].dbo.CMSMV");
            SB.AppendFormat(" WHERE TG001=TH001 AND TG002=TH002");
            SB.AppendFormat(" AND LA006=TH001 AND LA007=TH002 AND LA008=TH003");
            SB.AppendFormat(" AND MB001=TH004");
            SB.AppendFormat(" AND MV001=TG006");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TG005='{0}'",comboBox1.SelectedValue.ToString());
            SB.AppendFormat(" AND TG006='{0}'",comboBox2.SelectedValue.ToString());
            SB.AppendFormat(" GROUP BY MV002,TH005,MB004");
            SB.AppendFormat(" ORDER BY SUM(TH037) DESC");
            SB.AppendFormat("   ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT2()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL2();
            Report report2 = new Report();
            report2.Load(@"REPORT\業務客戶排名表.frx");

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

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report2.Preview = previewControl2;
            report2.Show();
        }

        public StringBuilder SETSQL2()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT MV002 AS '業務',TG007 AS '客戶',SUM(TH037) AS '金額',SUM(TH037)/SUM(SUM(TH037)) OVER ()  AS '金額百分比'");
            SB.AppendFormat(" FROM [TK].dbo.COPTG, [TK].dbo.COPTH,[TK].dbo.INVLA,[TK].dbo.INVMB,[TK].dbo.CMSMV");
            SB.AppendFormat(" WHERE TG001=TH001 AND TG002=TH002");
            SB.AppendFormat(" AND LA006=TH001 AND LA007=TH002 AND LA008=TH003");
            SB.AppendFormat(" AND MB001=TH004");
            SB.AppendFormat(" AND MV001=TG006");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TG005='{0}'", comboBox3.SelectedValue.ToString());
            SB.AppendFormat(" AND TG006='{0}'", comboBox4.SelectedValue.ToString());
            SB.AppendFormat(" GROUP BY MV002,TG007");
            SB.AppendFormat(" ORDER BY SUM(TH037) DESC");
            SB.AppendFormat("  ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT3()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL3();
            Report report3 = new Report();
            report3.Load(@"REPORT\門市商品排名表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report3.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report3.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report3.Preview = previewControl3;
            report3.Show();
        }

        public StringBuilder SETSQL3()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat("  SELECT MA002 AS '門市',MB002 AS '品名',SUM(TB031)  AS '銷售金額',SUM(TB019)  AS '銷售數量',MB004 AS '單位',SUM(TB031)/SUM(SUM(TB031)) OVER (partition by MA002)  AS '金額百分比'");
            SB.AppendFormat(" ,((SELECT SUM(TB031) FROM [TK].dbo.POSTB TB WITH(NOLOCK) WHERE TB.TB002=POSTB.TB002 AND TB.TB010=POSTB.TB010 AND TB.TB001>='{0}' AND TB.TB002<='{1}' )/(SELECT SUM(TB031) FROM [TK].dbo.POSTB TB WITH(NOLOCK) WHERE TB.TB002=POSTB.TB002  AND TB.TB001>='{2}' AND TB.TB002<='{3}' )) AS '月百分比'", dateTimePicker5.Value.ToString("yyyyMM")+"01", dateTimePicker6.Value.ToString("yyyyMMdd"), dateTimePicker5.Value.ToString("yyyyMM") + "01", dateTimePicker6.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" FROM [TK].dbo.POSTB,[TK].dbo.INVMB,[TK].dbo.WSCMA");
            SB.AppendFormat(" WHERE TB010=MB001");
            SB.AppendFormat(" AND TB002=MA001");
            SB.AppendFormat(" AND TB001>='{0}' AND TB002<='{1}'", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND ( TB010 LIKE '4%' OR TB010 LIKE '5%' )");
            SB.AppendFormat(" AND TB002 IN ('106701','106502','106503','106504','106513','106514','106501')");
            SB.AppendFormat(" GROUP BY MA002,MB002,MB004,TB002,TB010  ");
            SB.AppendFormat(" HAVING SUM(TB031)>0");
            SB.AppendFormat(" ORDER BY MA002,SUM(TB031) DESC,TB002,TB010");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        public void SETFASTREPORT4()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL4();
            Report report4 = new Report();
            report4.Load(@"REPORT\業績報表-營銷.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report4.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report4.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report4.Preview = previewControl4;
            report4.Show();
        }

        public StringBuilder SETSQL4()
        {
            StringBuilder SB = new StringBuilder();

            
            SB.AppendFormat(" SELECT TA002 AS '代號',MA002 AS '名稱',SUM(TA026)  AS '未稅金額'");
            SB.AppendFormat(" FROM [TK].dbo.POSTA,[TK].dbo.WSCMA");
            SB.AppendFormat(" WHERE TA002=MA001");
            SB.AppendFormat(" AND TA001>='{0}' AND TA001<='{1}'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TA002 NOT LIKE '1067%'");
            SB.AppendFormat(" GROUP BY TA002,MA002");
            SB.AppendFormat(" HAVING SUM(TA026)>0");
            SB.AppendFormat(" UNION ALL");
            SB.AppendFormat(" SELECT TA002 AS '代號',MA002 AS '名稱',SUM(TA026)  AS '未稅金額'");
            SB.AppendFormat(" FROM [TK].dbo.POSTA,[TK].dbo.WSCMA");
            SB.AppendFormat(" WHERE TA002=MA001");
            SB.AppendFormat(" AND TA001>='{0}' AND TA001<='{1}'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TA002 LIKE '1067%'");
            SB.AppendFormat(" GROUP BY TA002,MA002");
            SB.AppendFormat(" HAVING SUM(TA026)>0");
            SB.AppendFormat(" UNION ALL");
            SB.AppendFormat(" SELECT TG005,'官網',SUM(TG045)");
            SB.AppendFormat(" FROM [TK].dbo.COPTG,[TK].dbo.CMSME");
            SB.AppendFormat(" WHERE TG005=ME001");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TG005 IN ('116300')");
            SB.AppendFormat(" AND TG001 NOT IN ('A230')");
            SB.AppendFormat(" GROUP BY TG005,ME002");
            SB.AppendFormat(" HAVING SUM(TG045)>0");
            SB.AppendFormat(" UNION ALL");
            SB.AppendFormat(" SELECT TG005,'現銷',SUM(TG045)");
            SB.AppendFormat(" FROM [TK].dbo.COPTG,[TK].dbo.CMSME");
            SB.AppendFormat(" WHERE TG005=ME001");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND TG005 IN ('116300')");
            SB.AppendFormat(" AND TG001 IN ('A230')");
            SB.AppendFormat(" GROUP BY TG005,ME002");
            SB.AppendFormat(" HAVING SUM(TG045)>0");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");


            return SB;

        }

        public void SETFASTREPORT5()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL5();
            Report report5 = new Report();
            report5.Load(@"REPORT\業績報表-事拓.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report5.Load(@"REPORT\業績報表-事拓.frx");
            report5.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            TableDataSource table = report5.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report5.Preview = previewControl5;
            report5.Show();
        }

        public StringBuilder SETSQL5()
        {
            StringBuilder SB = new StringBuilder();
                       
            SB.AppendFormat(" SELECT TG006 AS '代號',MV002 AS '名稱',(SUM(TG045)-(SELECT ISNULL(SUM(TI010),0) FROM [TK].dbo.COPTI WHERE TI006=TG006 AND TI001 IN ('A241','A242') AND TI003>='{0}' AND TI003<='{1}')) AS '未稅金額'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" FROM [TK].dbo.COPTG,[TK].dbo.CMSMV");
            SB.AppendFormat(" WHERE TG006=MV001");
            SB.AppendFormat(" AND TG001 IN ('A231','A232')");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" GROUP BY TG006,MV002");
            SB.AppendFormat(" HAVING SUM(TG045)>0");
            SB.AppendFormat(" UNION ALL");
            SB.AppendFormat(" SELECT TG006 AS '代號','全聯' AS '名稱',(SUM(TG045)-(SELECT ISNULL(SUM(TI010),0) FROM [TK].dbo.COPTI WHERE TI006=TG006 AND TI001 IN ('A244') AND TI003>='{0}' AND TI003<='{1}')) AS '未稅金額'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" FROM [TK].dbo.COPTG,[TK].dbo.CMSMV");
            SB.AppendFormat(" WHERE TG006=MV001");
            SB.AppendFormat(" AND TG001 IN ('A237')");
            SB.AppendFormat(" AND TG003>='{0}' AND TG003<='{1}'", dateTimePicker7.Value.ToString("yyyyMMdd"), dateTimePicker8.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" GROUP BY TG006,MV002");
            SB.AppendFormat(" HAVING SUM(TG045)>0");
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
            SETFASTREPORT4();
            SETFASTREPORT5();
        }

        #endregion


    }
}
