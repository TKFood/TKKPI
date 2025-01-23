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
    public partial class frmRPOETPOSMA : Form
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

        public frmRPOETPOSMA()
        {
            InitializeComponent();
            SETDATE();

        }

        #region FUNCTION
        public void SETDATE()
        {
            dateTimePicker1.Value = DateTime.Now;
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
                                    SELECT *
                                    ,ISNULL((SELECT SUM(TB019) FROM [TK].dbo.POSTB WHERE TB036=活動代號),0) AS '總銷售數量'
                                    ,ISNULL((SELECT SUM(TB031) FROM [TK].dbo.POSTB WHERE TB036=活動代號),0) AS '總未稅金額'
                                    FROM 
	                                    (
	                                    SELECT '活動特價' AS '類型',MB004 AS '活動名稱',MB012 AS '開始日',MB013 AS '結束日',MB003 AS '活動代號'
	                                    FROM [TK].dbo.POSMB
	                                    WHERE 1=1
	                                    AND MB008='Y'
	                                    AND MB013 LIKE '{0}%'
	                                    UNION ALL
	                                    SELECT  '組合品搭贈' AS KIND,MI004,MI005,MI006,MI003
	                                    FROM [TK].dbo.POSMI
	                                    WHERE 1=1
	                                    AND MI015='Y'
	                                    AND MI005 LIKE  '{0}%'
	                                    UNION ALL
	                                    SELECT  '滿額折價' AS KIND,MM004,MM005,MM006,MM003
	                                    FROM [TK].dbo.POSMM
	                                    WHERE 1=1
	                                    AND MM015='Y'
	                                    AND MM005 LIKE  '{0}%'
	                                    UNION ALL
	                                    SELECT  '配對搭贈' AS KIND,MO004,MO005,MO006,MO003
	                                    FROM [TK].dbo.POSMO
	                                    WHERE 1=1
	                                    AND MO008='Y'
	                                    AND MO005 LIKE  '{0}%'
                                    ) AS TEMP 
                                    WHERE 1=1
                                    ORDER BY 類型,活動代號

--

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

                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    dataGridView1.Columns["類型"].Width = 100;
                    dataGridView1.Columns["活動名稱"].Width = 240;
                    dataGridView1.Columns["開始日"].Width = 100;
                    dataGridView1.Columns["結束日"].Width = 100;
                    dataGridView1.Columns["活動代號"].Width = 200;
                    dataGridView1.Columns["總銷售數量"].Width = 100;
                    dataGridView1.Columns["總未稅金額"].Width = 100;
                    dataGridView1.Columns["總銷售數量"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView1.Columns["總銷售數量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊
                    dataGridView1.Columns["總未稅金額"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView1.Columns["總未稅金額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊 
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
            string TB036 = null;
            dataGridView2.DataSource = null;
            dataGridView3.DataSource = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    TB036 = row.Cells["活動代號"].Value.ToString();
                   

                    SEARCH_POS_SET(TB036);
                    SEARCH_POS_POSTB(TB036);
                    SEARCH_POSTB_ME001(TB036);
                    SEARCH_POSTB_ME001_DAILY(TB036);

                }
                else
                {


                }
            }
        }

        public void SEARCH_POS_SET(string TB036)
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
                                    SELECT MC004 AS '品號',INVMB.MB002 AS '品名'
                                    FROM [TK].dbo.POSMC,[TK].dbo.INVMB,[TK].dbo.POSMB
                                    WHERE 1=1
                                    AND MC004=INVMB.MB001
                                    AND POSMB.MB003=MC003
                                    AND MC011='Y'
                                    AND MC003='{0}'
                                    UNION ALL
                                    SELECT MJ004,MB002
                                    FROM [TK].dbo.POSMJ,[TK].dbo.INVMB,[TK].dbo.POSMI
                                    WHERE 1=1
                                    AND MJ004=MB001
                                    AND MI003=MJ003
                                    AND MJ006='Y'
                                    AND MJ003='{0}'
                                    UNION ALL
                                    SELECT CONVERT(NVARCHAR,MN005),'金額以上'
                                    FROM [TK].dbo.POSMN,[TK].dbo.POSMM
                                    WHERE 1=1
                                    AND MN003=MM003
                                    AND MN010='Y'
                                    AND MN003='{0}'
                                    UNION ALL
                                    SELECT MP005,MB002
                                    FROM [TK].dbo.POSMP,[TK].dbo.INVMB,[TK].dbo.POSMO
                                    WHERE 1=1
                                    AND MP005=MB001
                                    AND MP003=MO003
                                    AND MP008='Y'
                                    AND MP003='{0}'

                                    ", TB036);



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
                    dataGridView2.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

                    dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView2.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    
                }


            }
            catch
            {

            }
            finally
            {

            }
        }
        public void SEARCH_POS_POSTB(string TB036)
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
                                    SELECT TB010  AS '品號',MB002 AS '品名',CONVERT(INT,SUM(TB019)) AS '銷售數量',CONVERT(INT,SUM(TB031)) 未稅金額
                                    FROM [TK].dbo.POSTB,[TK].dbo.INVMB
                                    WHERE TB010=MB001
                                    AND ISNULL(TB036,'')<>''
                                    AND TB036='{0}'
                                    GROUP BY TB010,MB002
                                    ORDER BY  TB010,MB002


                                    ", TB036);



                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, talbename);
                sqlConn.Close();


                if (ds.Tables[talbename].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else
                {
                    dataGridView3.DataSource = ds.Tables[talbename];
                    dataGridView3.AutoResizeColumns();
                    //rownum = ds.Tables[talbename].Rows.Count - 1;
                    dataGridView3.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

                    dataGridView3.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView3.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    dataGridView3.Columns["銷售數量"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView3.Columns["銷售數量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊
                    dataGridView3.Columns["未稅金額"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView3.Columns["未稅金額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊 
                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        public void SEARCH_POSTB_ME001(string TB036)
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
                                    ME001 AS '門市ID',
                                    ME002 AS '門市',
                                    TB010  AS '品號',
                                    MB002 AS '品名', 
                                    SUM(TB019) AS '銷售數量', 
                                    SUM(TB031)  AS '銷售未稅金額'
                                    FROM [TK].dbo.POSTB,[TK].dbo.INVMB,[TK].dbo.CMSME
                                    WHERE TB010=MB001
                                    AND ME001=TB002
                                    AND TB036='{0}'
                                    GROUP BY ME001,ME002,TB010,MB002
                                    ORDER BY ME001,ME002,TB010,MB002

                                    ", TB036);



                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, talbename);
                sqlConn.Close();


                if (ds.Tables[talbename].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;
                }
                else
                {
                    dataGridView4.DataSource = ds.Tables[talbename];
                    dataGridView4.AutoResizeColumns();
                    //rownum = ds.Tables[talbename].Rows.Count - 1;
                    dataGridView4.CurrentCell = dataGridView4.Rows[rownum].Cells[0];

                    dataGridView4.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView4.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    dataGridView4.Columns["銷售數量"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView4.Columns["銷售數量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊
                    dataGridView4.Columns["銷售未稅金額"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView4.Columns["銷售未稅金額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊 
                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        public void SEARCH_POSTB_ME001_DAILY(string TB036)
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
                                    ME001 AS '門市ID',
                                    ME002 AS '門市',
                                    TB001 AS '日期',
                                    TB010  AS '品號',
                                    MB002 AS '品名', 
                                    SUM(TB019) AS '銷售數量', 
                                    SUM(TB031)  AS '銷售未稅金額'
                                    FROM [TK].dbo.POSTB,[TK].dbo.INVMB,[TK].dbo.CMSME
                                    WHERE TB010=MB001
                                    AND ME001=TB002
                                    AND TB036='{0}'
                                    GROUP BY ME001,ME002,TB001,TB010,MB002
                                    ORDER BY ME001,ME002,TB001,TB010,MB002

                                    ", TB036);



                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, talbename);
                sqlConn.Close();


                if (ds.Tables[talbename].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;
                }
                else
                {
                    dataGridView5.DataSource = ds.Tables[talbename];
                    dataGridView5.AutoResizeColumns();
                    //rownum = ds.Tables[talbename].Rows.Count - 1;
                    dataGridView5.CurrentCell = dataGridView5.Rows[rownum].Cells[0];

                    dataGridView5.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView5.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    dataGridView5.Columns["銷售數量"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView5.Columns["銷售數量"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊
                    dataGridView5.Columns["銷售未稅金額"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView5.Columns["銷售未稅金額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊 
                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        #endregion

        #region BUTTON
        private void button7_Click(object sender, EventArgs e)
        {
            SearchPOS(dateTimePicker1.Value.ToString("yyyy"));
        }


        #endregion

       
    }
}
