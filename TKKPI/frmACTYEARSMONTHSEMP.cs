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

namespace TKKPI
{
    public partial class frmACTYEARSMONTHSEMP : Form
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
        string tablename = null;
        int rownum = 0;
        string Dep;
        SqlTransaction tran;
        int result;
        DateTime YEARSMONTHS;

        public frmACTYEARSMONTHSEMP()
        {
            InitializeComponent();
        }
        public frmACTYEARSMONTHSEMP(DateTime dt)
        {
            YEARSMONTHS = dt;
            InitializeComponent();
            dateTimePicker1.Value = dt;
            Search(YEARSMONTHS.ToString("yyyyMM"));
        }

        #region FUNCTION
        public void Search(string ID)
        {
            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSql.AppendFormat(" SELECT [ID],[YEARSMONTH],[EMP] FROM [TKKPI].[dbo].[ACTYEARSMONTHSEMP] WHERE [YEARSMONTH]='{0}'", ID.ToString());

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    button1.Visible = true;
                    button2.Visible = false;

                }
                else
                {
                    button1.Visible = false;
                    button2.Visible = true;

                    numericUpDown1.Value = Convert.ToInt32(ds.Tables["TEMPds1"].Rows[0]["EMP"].ToString());
                    textBox1.Text = ds.Tables["TEMPds1"].Rows[0]["ID"].ToString();

                }



            }
            catch
            {

            }
            finally
            {

            }

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
                sbSql.Append(" UPDATE [TKKPI].[dbo].[ACTYEARSMONTHSEMP] ");
                sbSql.AppendFormat("  SET [EMP]='{1}' WHERE [ID]='{0}' ", textBox1.Text.ToString(),numericUpDown1.Value.ToString());
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
                sbSql.Append(" INSERT INTO [TKKPI].[dbo].[ACTYEARSMONTHSEMP] ");
                sbSql.Append(" ([ID],[YEARSMONTH],[EMP])  ");
                sbSql.AppendFormat("  VALUES ('{0}','{1}','{2}') ", Guid.NewGuid(), dateTimePicker1.Value.ToString("yyyyMM"),numericUpDown1.Value.ToString());

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

        public void SETADD()
        {
            numericUpDown1.ReadOnly = false;
            numericUpDown1.Value = 0;
        }

        public void SETUPDATE()
        {
            numericUpDown1.ReadOnly = false;
        }

        public void SETFINISH()
        {
            numericUpDown1.ReadOnly = true;
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETADD();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETUPDATE();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text.ToString()))
            {
                UPDATE();
            }
            else
            {
                ADD();
            }
            Search(dateTimePicker1.Value.ToString("yyyyMM"));
            SETFINISH();
            this.Close();
        }
        #endregion


    }
}
