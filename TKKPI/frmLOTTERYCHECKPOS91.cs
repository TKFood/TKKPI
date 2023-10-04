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
    public partial class frmLOTTERYCHECKPOS91 : Form
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

        public frmLOTTERYCHECKPOS91()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
            comboBox3load();
        }

        #region FUNCTION
        public void comboBox1load()
        {
            ComboBox CBX = comboBox1;
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[KINDS],[NAMES],[VALUE] FROM [TKKPI].[dbo].[TBPARA] WHERE [KINDS]='TBLOTTERYCHECKPOS91' ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);

            CBX.DataSource = dt.DefaultView;
            CBX.ValueMember = "NAMES";
            CBX.DisplayMember = "NAMES";
            sqlConn.Close();

            CBX.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }
        public void comboBox2load()
        {
            ComboBox CBX = comboBox2;
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[KINDS],[NAMES],[VALUE] FROM [TKKPI].[dbo].[TBPARA] WHERE [KINDS]='TBLOTTERYCHECKPOS91' ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);

            CBX.DataSource = dt.DefaultView;
            CBX.ValueMember = "NAMES";
            CBX.DisplayMember = "NAMES";
            sqlConn.Close();

            CBX.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }

        public void comboBox3load()
        {
            ComboBox CBX = comboBox3;
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[KINDS],[NAMES],[VALUE] FROM [TKKPI].[dbo].[TBPARA] WHERE [KINDS]='TBLOTTERYCHECKPOS91CHECKNAME' ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);

            CBX.DataSource = dt.DefaultView;
            CBX.ValueMember = "NAMES";
            CBX.DisplayMember = "NAMES";
            sqlConn.Close();

            CBX.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }
        public void Search()
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();
            StringBuilder SQLQUERY1 = new StringBuilder();
            StringBuilder SQLQUERY2 = new StringBuilder();


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

                if(!string.IsNullOrEmpty(comboBox1.Text.ToString())&& !comboBox1.Text.ToString().Equals("全部"))
                {
                    SQLQUERY1.AppendFormat(@"AND [ISCHECK] IN ('{0}') ", comboBox1.Text.ToString());
                }
                else if (comboBox1.Text.ToString().Equals("全部"))
                {
                    SQLQUERY1.AppendFormat(@" ");
                }
                else
                {
                    SQLQUERY1.AppendFormat(@" ");
                }
                if (!string.IsNullOrEmpty(comboBox2.Text.ToString()) && !comboBox2.Text.ToString().Equals("全部"))
                {
                    SQLQUERY2.AppendFormat(@"AND [ISCHECK2] IN ('{0}') ", comboBox2.Text.ToString());
                }
                else if(comboBox2.Text.ToString().Equals("全部"))
                {
                    SQLQUERY2.AppendFormat(@" ");
                }
                else 
                {
                    SQLQUERY2.AppendFormat(@" ");
                }

                sbSql.Clear();
                sbSql.AppendFormat(@" 
                                    SELECT 
                                     [ID] AS '登錄時間'
                                    ,[KINDS] AS '通路' 
                                    ,[BILLPOS] AS '發票'
                                    ,[BILL91] AS '購物車'
                                    ,[NUMS] AS '購買件數'
                                    ,[ISCHECK] AS '是否檢查1'
                                    ,[CHECKNAME]  AS '檢查人1'
                                    ,[CHECKTIME]  AS '檢查時間1'
                                    ,[ISCHECK2]  AS '是否檢查2'
                                    ,[CHECKNAME2] AS '檢查時間2'
                                    ,[CHECKTIME2] AS '是否檢查2'
                                    FROM [TKKPI].[dbo].[TBLOTTERYCHECKPOS91]
                                    WHERE 1=1
                                    {0}
                                    {1}
                                    ORDER BY [KINDS],[ID]
                                    ", SQLQUERY1.ToString(), SQLQUERY2.ToString());


                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();

                
                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    dataGridView1.DataSource = ds.Tables["ds"];
                    dataGridView1.AutoResizeColumns();                 

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
            string KEY1 = null;
            string KEY2 = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    KEY1 = row.Cells["發票"].Value.ToString();
                    KEY2 = row.Cells["購物車"].Value.ToString();

                    if (!string.IsNullOrEmpty(KEY1) || !string.IsNullOrEmpty(KEY2))
                    {
                        Search_POS_91(KEY1, KEY2);
                    }
                }
                else
                {

                }
            }

        }

        public void Search_POS_91(string KEY1,string KEY2)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            StringBuilder SQLQUERY1 = new StringBuilder();
            StringBuilder SQLQUERY2 = new StringBuilder();


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

                if (!string.IsNullOrEmpty(KEY1) )
                {
                    SQLQUERY1.AppendFormat(@"   AND TB008='{0}'  ", KEY1);
                }               
                else
                {
                    SQLQUERY1.AppendFormat(@" AND 1=0 ");
                }
                if (!string.IsNullOrEmpty(KEY2))
                {
                    SQLQUERY2.AppendFormat(@" AND (TG020='{0}' OR TG029='{0}') ", KEY2);
                }              
                else
                {
                    SQLQUERY2.AppendFormat(@" AND 1=0 ");
                }


                sbSql.Clear();

                if (!string.IsNullOrEmpty(KEY1) || !string.IsNullOrEmpty(KEY2))
                {                   
                    sbSql.AppendFormat(@" 
                                    SELECT 
                                    TB008 AS '發票號碼+購物車'
                                    ,TB001 AS '交易日期'
                                    ,TB002 AS '店號'
                                    ,TB010 AS '品號'
                                    ,MB002 AS '品名'
                                    ,SUM(TB019)  AS '銷售數量'
                                    FROM[TK].dbo.POSTB  WITH(NOLOCK)
                                    LEFT JOIN [TK].dbo.INVMB ON MB001=TB010
                                    WHERE 1=1
                                    {0}
                                    GROUP BY TB008
                                    ,TB001
                                    ,TB002
                                    ,TB010
                                    ,MB002

                                    UNION ALL
                                    SELECT 
                                    TG020 AS '發票號碼+購物車'
                                    ,TG003 AS '交易日期'
                                    ,TG007 AS '店號'
                                    ,TH004 AS '品號'
                                    ,TH005  AS '品名'
                                    ,SUM(TH008+TH024)  AS '銷售數量'
                                    FROM [TK].dbo.COPTG,[TK].dbo.COPTH
                                    WHERE 1=1
                                    AND TG001=TH001 AND TG002=TH002
                                    {1}
                                    GROUP BY  TG020,TG003,TG007,TH004,TH005

                                    ", SQLQUERY1.ToString(), SQLQUERY2.ToString());

                }


                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "ds");
                sqlConn.Close();


                if (ds.Tables["ds"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    dataGridView2.DataSource = ds.Tables["ds"];
                    dataGridView2.AutoResizeColumns();
                 

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
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        #endregion

     
    }
}
