using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;

namespace TKKPI
{
    public partial class frmMARKETMONTHSETDETAIL : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        string EDITEquipmentID;
        int result;
        Thread TD;
        int rownum = 0;
        string YEARSMONTH;

        public frmMARKETMONTHSETDETAIL()
        {
            InitializeComponent();
        }
        public frmMARKETMONTHSETDETAIL(string ID)
        {
            InitializeComponent();           

            if (!string.IsNullOrEmpty(ID))
            {
                YEARSMONTH = ID;
                Search(ID);

            }
        }
        public void Search(string ID)
        {
            StringBuilder Query = new StringBuilder();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.Clear();
                sbSql.AppendFormat("SELECT  [YEARMONTH] AS '活動年月',[MB001] AS '品號',[MB002] AS '品名',[MONTHSET] AS '活動內容',[ID] FROM [TKKPI].[dbo].[MARKETMONTHSET] WHERE [YEARMONTH]='{0}'", ID);


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    label1.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

                        //dataGridView1.CurrentCell = dataGridView1[0, 2];
                    

                    }
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
            if (dataGridView1.Rows.Count >= 1)
            {
                //dateTimePicker1.Value =Convert.ToDateTime(YEARSMONTH.Substring(0,4).ToString()+"/"+ YEARSMONTH.Substring(4, 2).ToString() + "/01");
                textBox1.Text = dataGridView1.CurrentRow.Cells["品號"].Value.ToString();
                textBox2.Text = dataGridView1.CurrentRow.Cells["品名"].Value.ToString();
                textBox3.Text = dataGridView1.CurrentRow.Cells["活動內容"].Value.ToString();
                textBoxID.Text = dataGridView1.CurrentRow.Cells["ID"].Value.ToString();

            }

        }
        public void SETADDUPDATE()
        {
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
        }
        public void SETFINISH()
        {
            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;
            textBox3.ReadOnly = true;
        }

        public void UPDATE()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" UPDATE [TKKPI].[dbo].[MARKETMONTHSET] ");
                sbSql.AppendFormat(" SET [YEARMONTH]='{1}',[MB001]='{2}',[MB002]='{3}',[MONTHSET]='{4}' WHERE [ID]='{0}' ", textBoxID.Text.ToString(),dateTimePicker1.Value.ToString("yyyyMM"),textBox1.Text.ToString().Trim(), textBox2.Text.ToString().Trim(), textBox3.Text.ToString());
                sbSql.Append("   ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                    

                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        public void ADD()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.Append(" INSERT INTO [TKKPI].[dbo].[MARKETMONTHSET] ");
                sbSql.Append("  ([ID],[YEARMONTH],[MB001],[MB002],[MONTHSET] )  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}','{3}','{4}') ", Guid.NewGuid(), dateTimePicker1.Value.ToString("yyyyMM"), textBox1.Text.ToString().Trim(), textBox2.Text.ToString().Trim(), textBox3.Text.ToString());

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  
                    

                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        #region FUNCTION

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
            textBoxID.Text = null;
            SETADDUPDATE();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETADDUPDATE();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBoxID.Text.ToString()))
            {
                UPDATE();
            }
            else
            {
                ADD();
            }
            Search(dateTimePicker1.Value.ToString("yyyyMM"));

        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Search(dateTimePicker1.Value.ToString("yyyyMM"));
        }
        #endregion


    }
}
