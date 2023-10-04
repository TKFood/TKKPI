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
            Sequel.AppendFormat(@"SELECT  [ID],[KINDS],[NAMES],[VALUE] FROM [TKKPI].[dbo].[TBPARA] WHERE [KINDS]='TBLOTTERYCHECKPOS91' ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);

            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "NAMES";
            comboBox1.DisplayMember = "NAMES";
            sqlConn.Close();

            comboBox1.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
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
            Sequel.AppendFormat(@"SELECT  [ID],[KINDS],[NAMES],[VALUE] FROM [TKKPI].[dbo].[TBPARA] WHERE [KINDS]='TBLOTTERYCHECKPOS91' ORDER BY ID ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAMES", typeof(string));
            da.Fill(dt);

            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "NAMES";
            comboBox2.DisplayMember = "NAMES";
            sqlConn.Close();

            comboBox2.Font = new Font("Arial", 10); // 使用 "Arial" 字體，字體大小為 12
        }
        public void Search()
        {
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
                                    ORDER BY [ID]
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

       
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        #endregion
    }
}
