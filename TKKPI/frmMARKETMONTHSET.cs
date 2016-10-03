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

                 connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

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
