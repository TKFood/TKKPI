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
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
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
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
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
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
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
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
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

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
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

            report2.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
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

            report3.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
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
            SB.AppendFormat(" FROM [TK].dbo.POSTB,[TK].dbo.INVMB,[TK].dbo.WSCMA");
            SB.AppendFormat(" WHERE TB010=MB001");
            SB.AppendFormat(" AND TB002=MA001");
            SB.AppendFormat(" AND TB001>='{0}' AND TB002<='{1}'", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" AND ( TB010 LIKE '4%' OR TB010 LIKE '5%' )");
            SB.AppendFormat(" AND TB002 IN ('106701','106502','106503','106504','106513','106514')");
            SB.AppendFormat(" GROUP BY MA002,MB002,MB004 ");
            SB.AppendFormat(" HAVING SUM(TB031)>0");
            SB.AppendFormat(" ORDER BY MA002,SUM(TB031) DESC");
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
        #endregion


    }
}
