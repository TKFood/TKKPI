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
            dateTimePicker2.Value = DateTime.Now;
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

        public void SearchPOS_DG6(string sYears)
        {
            // 檢查輸入參數
            if (string.IsNullOrWhiteSpace(sYears)) return;

            try
            {
                // 取得連線字串並解密
                var connectionString = ConfigurationManager.ConnectionStrings["dbconn"]?.ConnectionString;
                if (string.IsNullOrEmpty(connectionString)) return;

                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(connectionString);
                Class1 tkid = new Class1();
                sqlsb.Password = tkid.Decryption(sqlsb.Password);
                sqlsb.UserID = tkid.Decryption(sqlsb.UserID);

                // 使用參數化查詢，避免 SQL 注入
                string sql = @"
                                SELECT 
                                    '滿額折價' AS [類型],
                                    MM004 AS [活動名稱],
                                    MM005 AS [開始日],
                                    MM006 AS [結束日],
                                    MM003 AS [活動代號]
                                FROM [TK].dbo.POSMM
                                WHERE MM015 = 'Y'
                                  AND MM005 LIKE @SYEARS + '%'
                                ORDER BY [類型], [活動代號]";

                DataTable dataTable = new DataTable();

                // 使用 using 確保 Connection 與 Adapter 離開區塊後自動釋放資源
                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@SYEARS", sYears);

                    // Fill 會自動 Handle Open/Close，但明確指定較清晰
                    adapter.Fill(dataTable);
                }

                // UI 資料繫結與樣式調整
                if (dataTable.Rows.Count == 0)
                {
                    dataGridView6.DataSource = null;
                }
                else
                {
                    dataGridView6.DataSource = dataTable;
                    dataGridView6.AutoResizeColumns();

                    // 安全性設置 CurrentCell (確保 rownum 位於有效範圍內)
                    if (rownum >= 0 && rownum < dataGridView6.Rows.Count)
                    {
                        dataGridView6.CurrentCell = dataGridView6.Rows[rownum].Cells[0];
                    }

                    // 設定 DataGridView 欄位樣式
                    dataGridView6.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView6.DefaultCellStyle.Font = new Font("Tahoma", 10);

                    if (dataGridView6.Columns["類型"] != null) dataGridView6.Columns["類型"].Width = 100;
                    if (dataGridView6.Columns["活動名稱"] != null) dataGridView6.Columns["活動名稱"].Width = 240;
                    if (dataGridView6.Columns["開始日"] != null) dataGridView6.Columns["開始日"].Width = 100;
                    if (dataGridView6.Columns["結束日"] != null) dataGridView6.Columns["結束日"].Width = 100;
                    if (dataGridView6.Columns["活動代號"] != null) dataGridView6.Columns["活動代號"].Width = 200;
                }
            }
            catch (Exception ex)
            {
                // 記錄 Exception 或跳出提示通知使用者
                //MessageBox.Show($"查詢失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            string TA034 = null;
            dataGridView7.DataSource = null;
            dataGridView8.DataSource = null;

            if (dataGridView6.CurrentRow != null)
            {
                int rowindex = dataGridView6.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView6.Rows[rowindex];
                    TA034 = row.Cells["活動代號"].Value.ToString();

                    SEARCH_DG7(TA034);
                    SEARCH_DG8(TA034);
                }
                else
                {


                }
            }
        }

        public void SEARCH_DG7(string TA034)
        {
            // 檢查輸入參數
            if (string.IsNullOrWhiteSpace(TA034)) return;

            try
            {
                // 取得連線字串並解密
                var connectionString = ConfigurationManager.ConnectionStrings["dbconn"]?.ConnectionString;
                if (string.IsNullOrEmpty(connectionString)) return;

                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(connectionString);
                Class1 tkid = new Class1();
                sqlsb.Password = tkid.Decryption(sqlsb.Password);
                sqlsb.UserID = tkid.Decryption(sqlsb.UserID);

                // 使用參數化查詢，避免 SQL 注入
                string sql = @"
                               SELECT 
                                TA002 AS '門市代號'
                                ,ME002 AS '門市'
                                ,COUNT(TA014) AS '交易筆數'
                                ,SUM(TA033) AS '交易金額'
                                ,SUM(TA025) AS '折扣金額'

                                FROM [TK].dbo.POSTA WITH(NOLOCK)
                                INNER JOIN [TK].dbo.CMSME ON ME001=TA002
                                WHERE TA034=@TA034
                                GROUP BY TA002,ME002
                                ORDER BY  TA002
                                ";

                DataTable dataTable = new DataTable();

                // 使用 using 確保 Connection 與 Adapter 離開區塊後自動釋放資源
                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@TA034", TA034);

                    // Fill 會自動 Handle Open/Close，但明確指定較清晰
                    adapter.Fill(dataTable);
                }

                // UI 資料繫結與樣式調整
                if (dataTable.Rows.Count == 0)
                {
                    dataGridView7.DataSource = null;
                }
                else
                {
                    dataGridView7.DataSource = dataTable;
                    dataGridView7.AutoResizeColumns();

                    // 安全性設置 CurrentCell (確保 rownum 位於有效範圍內)
                    if (rownum >= 0 && rownum < dataGridView7.Rows.Count)
                    {
                        dataGridView7.CurrentCell = dataGridView7.Rows[rownum].Cells[0];
                    }

                    // 設定 DataGridView 欄位樣式
                    dataGridView7.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView7.DefaultCellStyle.Font = new Font("Tahoma", 10);

                    dataGridView7.Columns["交易筆數"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView7.Columns["交易筆數"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊 
                    dataGridView7.Columns["交易金額"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView7.Columns["交易金額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊 
                    dataGridView7.Columns["折扣金額"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView7.Columns["折扣金額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊 
                }
            }
            catch (Exception ex)
            {
                // 記錄 Exception 或跳出提示通知使用者
                //MessageBox.Show($"查詢失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        public void SEARCH_DG8(string TA034)
        {
            // 檢查輸入參數
            if (string.IsNullOrWhiteSpace(TA034)) return;

            try
            {
                // 取得連線字串並解密
                var connectionString = ConfigurationManager.ConnectionStrings["dbconn"]?.ConnectionString;
                if (string.IsNullOrEmpty(connectionString)) return;

                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(connectionString);
                Class1 tkid = new Class1();
                sqlsb.Password = tkid.Decryption(sqlsb.Password);
                sqlsb.UserID = tkid.Decryption(sqlsb.UserID);

                // 使用參數化查詢，避免 SQL 注入
                string sql = @"
                               SELECT 
                            TA002 AS '門市代號'
                            ,ME002 AS '門市'
                            ,TA001 AS '交易日期'
                            ,TA014 AS '發票'
                            ,TA033 AS '交易金額'
                            ,TA025 AS '折扣金額'

                            FROM [TK].dbo.POSTA WITH(NOLOCK)
                            INNER JOIN [TK].dbo.CMSME ON ME001=TA002
                            WHERE TA034=@TA034
                            ORDER BY  TA002,TA001
                                ";

                DataTable dataTable = new DataTable();

                // 使用 using 確保 Connection 與 Adapter 離開區塊後自動釋放資源
                using (SqlConnection conn = new SqlConnection(sqlsb.ConnectionString))
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    cmd.Parameters.AddWithValue("@TA034", TA034);

                    // Fill 會自動 Handle Open/Close，但明確指定較清晰
                    adapter.Fill(dataTable);
                }

                // UI 資料繫結與樣式調整
                if (dataTable.Rows.Count == 0)
                {
                    dataGridView8.DataSource = null;
                }
                else
                {
                    dataGridView8.DataSource = dataTable;
                    dataGridView8.AutoResizeColumns();

                    // 安全性設置 CurrentCell (確保 rownum 位於有效範圍內)
                    if (rownum >= 0 && rownum < dataGridView8.Rows.Count)
                    {
                        dataGridView8.CurrentCell = dataGridView8.Rows[rownum].Cells[0];
                    }

                    // 設定 DataGridView 欄位樣式
                    dataGridView8.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView8.DefaultCellStyle.Font = new Font("Tahoma", 10);

                    dataGridView8.Columns["交易金額"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView8.Columns["交易金額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊 
                    dataGridView8.Columns["折扣金額"].DefaultCellStyle.Format = "N0"; // 格式化為千分位，無小數位
                    dataGridView8.Columns["折扣金額"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight; // 右對齊 
                    
                }
            }
            catch (Exception ex)
            {
                // 記錄 Exception 或跳出提示通知使用者
                //MessageBox.Show($"查詢失敗：{ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region BUTTON
        private void button7_Click(object sender, EventArgs e)
        {
            SearchPOS(dateTimePicker1.Value.ToString("yyyy"));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string YEARS = dateTimePicker2.Value.ToString("yyyy");
            SearchPOS_DG6(YEARS);
        }


        #endregion

      
    }
}
