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
    public partial class frmDAILYCHECK : Form
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
        int rownum = 0;

        public frmDAILYCHECK()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            try
            {
                talbename = "TEMPds1";
                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT TD011,ISNULL(TB009,0) TB009,TC053,TD001,TD002,TD003,TD004,TD005,TD005,TD008,TD009,TD010,TD011,TD012,TD013,TD017,TD018,TD019,'',TB004,TB005,TB006,TB007,TB008,TB009,TB010");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.COPTB ON TB001=TD017 AND TB002=TD018 AND TB003=TD019 AND TB004=TD004 ");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND COPTD.MODIFIER='160115'");
                sbSql.AppendFormat(@"  AND TD002 LIKE '{0}%'",dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TD011<>ISNULL(TB009,0) ");
                sbSql.AppendFormat(@"  ");

                textBox1.Text = null;
                textBox1.Text = sbSql.ToString();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);



                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, talbename);
                    sqlConn.Close();

                   
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
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }

        }

        public void Search2()
        {
            try
            {
                talbename = "TEMPds2";
                sbSql.Clear();

                
                sbSql.AppendFormat(@"  SELECT COPTH.MODI_DATE,TH012,ISNULL(TD011,0) TD011,TG007,TH001,TH002,TH003,TH004,TH005,TH006,TH007,TH008,TH009,TH012,TH013,TH014,TH015,TH016,TD004,TD011");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTG,[TK].dbo.COPTH");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.COPTD ON TD001=TH014 AND TD002=TH015 AND TD003=TH016 AND TD004=TH004");
                sbSql.AppendFormat(@"  WHERE  TG001=TH001 AND TG002=TH002");
                sbSql.AppendFormat(@"  AND COPTH.MODIFIER='160115'");
                sbSql.AppendFormat(@"  AND COPTH.TH002 LIKE '{0}%'", dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TH012<>ISNULL(TD011,0) ");
                sbSql.AppendFormat(@"  AND TH001 IN ('A231','A232')");
                sbSql.AppendFormat(@"  ");

                textBox2.Text = null;
                textBox2.Text = sbSql.ToString();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);



                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, talbename);
                    sqlConn.Close();


                    if (ds.Tables[talbename].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        dataGridView2.DataSource = ds.Tables[talbename];
                        dataGridView2.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView2.CurrentCell = dataGridView2.Rows[rownum].Cells[0];

                        //dataGridView1.CurrentCell = dataGridView1[0, 2];

                    }
                }
                else
                {

                }



            }
            catch
            {

            }
            finally
            {

            }

        }

        public void Search3()
        {
            try
            {
                talbename = "TEMPds3";
                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT TJ011,ISNULL(TH012,0) TH012,TI021,TJ001,TJ002,TJ003,TJ004,TJ005,TJ006,TJ007,TJ008,TJ011,TJ012,TJ015,TJ016,TJ017,TH004,TH012");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTI,[TK].dbo.COPTJ");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.COPTH ON TH001=TJ015 AND TH002=TJ016 AND TH003=TJ017 AND TH004=TJ004");
                sbSql.AppendFormat(@"  WHERE  TI001=TJ001 AND TI002=TJ002");
                sbSql.AppendFormat(@"  AND COPTJ.MODIFIER='160115'");
                sbSql.AppendFormat(@"  AND COPTJ.TJ002 LIKE '{0}%'",dateTimePicker3.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND TJ011<>ISNULL(TH012,0) ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                textBox3.Text = null;
                textBox3.Text = sbSql.ToString();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);



                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, talbename);
                    sqlConn.Close();


                    if (ds.Tables[talbename].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        dataGridView3.DataSource = ds.Tables[talbename];
                        dataGridView3.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView3.CurrentCell = dataGridView3.Rows[rownum].Cells[0];

                        //dataGridView1.CurrentCell = dataGridView1[0, 2];

                    }
                }
                else
                {

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
            MessageBox.Show("QUERY");
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Search2();
            MessageBox.Show("QUERY");
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Search3();
            MessageBox.Show("QUERY");
        }

        #endregion


    }
}
