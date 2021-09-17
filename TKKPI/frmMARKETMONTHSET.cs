using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Threading;
using TKITDLL;

namespace TKKPI
{
    public partial class frmMARKETMONTHSET : Form
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
        DataTable dt = new DataTable();
        string talbename = null;
        string ID;
        int rownum = 0;

        public frmMARKETMONTHSET()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            try
            {
                talbename = "TEMP1";
                DataSet ds = new DataSet();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSql.AppendFormat(" SELECT  [YEARMONTH] AS '活動年月',[MB001] AS '品號',[MB002] AS '品名',[MONTHSET] AS '活動內容',[ID] FROM [TKKPI].[dbo].[MARKETMONTHSET]  WHERE [YEARMONTH]='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, talbename);
                sqlConn.Close();

                label1.Text = "資料筆數:" + ds.Tables[talbename].Rows.Count.ToString();

                if (ds.Tables[talbename].Rows.Count == 0)
                {

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
        public void SearchLastyear()
        {
            try
            {
                talbename = "TEMP2";
                DataSet ds = new DataSet();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSql.AppendFormat(" SELECT KIND AS '市場',TB010 AS '品號',MB002 AS '品名',NN AS '數量',MM AS '金額'");
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT '門市' AS KIND,TB010,MB002,SUM(TB019) AS NN,SUM(TB033) AS MM ");
                sbSql.AppendFormat(" FROM [TK].dbo.POSTB WITH (NOLOCK)");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.INVMB  WITH (NOLOCK) ON TB010=MB001");
                sbSql.AppendFormat(" WHERE TB001 LIKE '{0}%'",dateTimePicker1.Value.AddYears(-1).ToString("yyyyMM"));
                sbSql.AppendFormat(" AND TB010 IN ( SELECT [MB001] FROM [TKKPI].[dbo].[MARKETMONTHSET] WHERE [YEARMONTH]='{0}')", dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" GROUP BY TB010,MB002 ");
                sbSql.AppendFormat(" UNION ALL ");
                sbSql.AppendFormat(" SELECT '銷貨' AS KIND,TH004,TH005,SUM(LA011) AS NUM,SUM(TH037+TH038) AS MM ");
                sbSql.AppendFormat(" FROM [TK].dbo.COPTH   WITH (NOLOCK)");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.INVLA  WITH (NOLOCK) ON LA006=TH001 AND LA007=TH002 AND LA008=TH003");
                sbSql.AppendFormat(" WHERE TH020='Y'");
                sbSql.AppendFormat(" AND TH002 LIKE '{0}%'", dateTimePicker1.Value.AddYears(-1).ToString("yyyyMM"));
                sbSql.AppendFormat(" AND TH004 IN ( SELECT [MB001] FROM [TKKPI].[dbo].[MARKETMONTHSET] WHERE [YEARMONTH]='{0}')", dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" AND TH001 NOT IN ('A233')");
                sbSql.AppendFormat(" GROUP BY TH004,TH005");
                sbSql.AppendFormat(" UNION ALL ");
                sbSql.AppendFormat(" SELECT '電商' AS KIND,TH004,TH005,SUM(LA011) AS NUM,SUM(TH037+TH038) AS MM");
                sbSql.AppendFormat(" FROM [TK].dbo.COPTH   WITH (NOLOCK)");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.INVLA  WITH (NOLOCK) ON LA006=TH001 AND LA007=TH002 AND LA008=TH003");
                sbSql.AppendFormat(" WHERE TH020='Y'");
                sbSql.AppendFormat(" AND TH002 LIKE '{0}%'", dateTimePicker1.Value.AddYears(-1).ToString("yyyyMM"));
                sbSql.AppendFormat(" AND TH004 IN ( SELECT [MB001] FROM [TKKPI].[dbo].[MARKETMONTHSET] WHERE [YEARMONTH]='{0}')", dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" AND TH001  IN ('A233')");
                sbSql.AppendFormat(" GROUP BY TH004,TH005 ");
                sbSql.AppendFormat(" ) AS TEMP");
                //sbSql.AppendFormat(" ", dateTimePicker1.Value.ToString("yyyyMM"));

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, talbename);
                sqlConn.Close();

                //label1.Text = "資料筆數:" + ds.Tables[talbename].Rows.Count.ToString();

                if (ds.Tables[talbename].Rows.Count == 0)
                {

                }
                else
                {
                    dataGridView2.DataSource = ds.Tables[talbename];
                    dataGridView2.AutoResizeColumns();
                    //rownum = ds.Tables[talbename].Rows.Count - 1;
                    dataGridView2.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

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
        public void SearchNOWyear()
        {
            try
            {
                talbename = "TEMP3";
                DataSet ds = new DataSet();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSql.AppendFormat(" SELECT KIND AS '市場',TB010 AS '品號',MB002 AS '品名',NN AS '數量',MM AS '金額'");
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT '門市' AS KIND,TB010,MB002,SUM(TB019) AS NN,SUM(TB033) AS MM ");
                sbSql.AppendFormat(" FROM [TK].dbo.POSTB WITH (NOLOCK)");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.INVMB  WITH (NOLOCK) ON TB010=MB001");
                sbSql.AppendFormat(" WHERE TB001 LIKE '{0}%'", dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" AND TB010 IN ( SELECT [MB001] FROM [TKKPI].[dbo].[MARKETMONTHSET] WHERE [YEARMONTH]='{0}')", dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" GROUP BY TB010,MB002 ");
                sbSql.AppendFormat(" UNION ALL ");
                sbSql.AppendFormat(" SELECT '銷貨' AS KIND,TH004,TH005,SUM(LA011) AS NUM,SUM(TH037+TH038) AS MM ");
                sbSql.AppendFormat(" FROM [TK].dbo.COPTH   WITH (NOLOCK)");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.INVLA  WITH (NOLOCK) ON LA006=TH001 AND LA007=TH002 AND LA008=TH003");
                sbSql.AppendFormat(" WHERE TH020='Y'");
                sbSql.AppendFormat(" AND TH002 LIKE '{0}%'", dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" AND TH004 IN ( SELECT [MB001] FROM [TKKPI].[dbo].[MARKETMONTHSET] WHERE [YEARMONTH]='{0}')", dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" AND TH001 NOT IN ('A233')");
                sbSql.AppendFormat(" GROUP BY TH004,TH005");
                sbSql.AppendFormat(" UNION ALL ");
                sbSql.AppendFormat(" SELECT '電商' AS KIND,TH004,TH005,SUM(LA011) AS NUM,SUM(TH037+TH038) AS MM");
                sbSql.AppendFormat(" FROM [TK].dbo.COPTH   WITH (NOLOCK)");
                sbSql.AppendFormat(" LEFT JOIN [TK].dbo.INVLA  WITH (NOLOCK) ON LA006=TH001 AND LA007=TH002 AND LA008=TH003");
                sbSql.AppendFormat(" WHERE TH020='Y'");
                sbSql.AppendFormat(" AND TH002 LIKE '{0}%'", dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" AND TH004 IN ( SELECT [MB001] FROM [TKKPI].[dbo].[MARKETMONTHSET] WHERE [YEARMONTH]='{0}')", dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" AND TH001  IN ('A233')");
                sbSql.AppendFormat(" GROUP BY TH004,TH005 ");
                sbSql.AppendFormat(" ) AS TEMP");
                //sbSql.AppendFormat(" ", dateTimePicker1.Value.ToString("yyyyMM"));

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, talbename);
                sqlConn.Close();

                label1.Text = "資料筆數:" + ds.Tables[talbename].Rows.Count.ToString();

                if (ds.Tables[talbename].Rows.Count == 0)
                {

                }
                else
                {
                    dataGridView3.DataSource = ds.Tables[talbename];
                    dataGridView3.AutoResizeColumns();
                    //rownum = ds.Tables[talbename].Rows.Count - 1;
                    dataGridView3.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

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
        public void Searchyear()
        {
            try
            {
                talbename = "TEMP3";
                DataSet ds = new DataSet();

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSql.AppendFormat(" SELECT ID,NN+NN1 AS '數量',MM+MM1 AS'金額'");
                sbSql.AppendFormat(" FROM (");
                sbSql.AppendFormat(" SELECT [BASEMONTH].ID");
                sbSql.AppendFormat(" ,ISNULL((SELECT SUM(LA011) FROM [TK].dbo.COPTH  WITH (NOLOCK) LEFT JOIN [TK].dbo.INVLA  WITH (NOLOCK) ON LA006=TH001 AND LA007=TH002 AND LA008=TH003 WHERE TH002 LIKE '{0}'+[BASEMONTH].ID+'%'AND TH004 IN ( SELECT [MB001] FROM [TKKPI].[dbo].[MARKETMONTHSET] WHERE [YEARMONTH]='{1}')  AND TH020='Y' ),0) AS NN",dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" ,ISNULL((SELECT SUM(TH013) FROM [TK].dbo.COPTH  WITH (NOLOCK) LEFT JOIN [TK].dbo.INVLA  WITH (NOLOCK) ON LA006=TH001 AND LA007=TH002 AND LA008=TH003 WHERE TH002 LIKE '{0}'+[BASEMONTH].ID+'%'AND TH004 IN ( SELECT [MB001] FROM [TKKPI].[dbo].[MARKETMONTHSET] WHERE [YEARMONTH]='{1}')  AND TH020='Y' ),0) AS MM",dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" ,ISNULL((SELECT SUM(TB019) FROM [TK].dbo.POSTB WITH (NOLOCK) WHERE TB001 LIKE '{0}'+[BASEMONTH].ID+'%' AND TB010 IN ( SELECT [MB001] FROM [TKKPI].[dbo].[MARKETMONTHSET] WHERE [YEARMONTH]='{1}')),0 ) AS NN1 ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" ,ISNULL((SELECT SUM(TB033) FROM [TK].dbo.POSTB WITH (NOLOCK) WHERE TB001 LIKE '{0}'+[BASEMONTH].ID+'%' AND TB010 IN ( SELECT [MB001] FROM [TKKPI].[dbo].[MARKETMONTHSET] WHERE [YEARMONTH]='{1}')),0 ) AS MM1", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat(" FROM [TKKPI].dbo.[BASEMONTH] ");
                sbSql.AppendFormat(" ) AS TEMP ");
                sbSql.AppendFormat(" ");
                //sbSql.AppendFormat(" ", dateTimePicker1.Value.ToString("yyyyMM"));

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, talbename);
                sqlConn.Close();

                label1.Text = "資料筆數:" + ds.Tables[talbename].Rows.Count.ToString();

                if (ds.Tables[talbename].Rows.Count == 0)
                {

                }
                else
                {
                    dataGridView4.DataSource = ds.Tables[talbename];
                    dataGridView4.AutoResizeColumns();
                    //rownum = ds.Tables[talbename].Rows.Count - 1;
                    dataGridView4.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

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
        private void showwaitfrm()
        {
            try
            {
                PleaseWait objPleaseWait = new PleaseWait();
                objPleaseWait.ShowDialog();
            }
            catch
            {

            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count >= 1)
            {
                ID = dataGridView1.CurrentRow.Cells["活動年月"].Value.ToString();
            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Thread TD;

            TD = new Thread(showwaitfrm);
            TD.Start();
            Thread.Sleep(2000);   //此行可以不需要，主要用於等待主窗體填充數據
            Search();
            SearchLastyear();
            SearchNOWyear();
            Searchyear();
            TD.Abort(); //主窗體加載完成數據後，線程結束，關閉等待窗體。
        }
        private void button2_Click(object sender, EventArgs e)
        {
            frmMARKETMONTHSETDETAIL objfrmMARKETMONTHSETDETAIL = new frmMARKETMONTHSETDETAIL(ID);
            objfrmMARKETMONTHSETDETAIL.ShowDialog();
            button1.PerformClick();
        }

        #endregion

       
    }
}
