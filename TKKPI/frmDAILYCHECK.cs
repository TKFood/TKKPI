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

                sbSql.AppendFormat(@"  SELECT TD011,ISNULL(TB009,0) TB009,(TD008*TD011*TD026-TD012) AS DIFF,TC053,TD001,TD002,TD003,TD004,TD005,TD005,TD008,TD009,TD010,TD011,TD012,TD013,TD017,TD018,TD019,'',TB004,TB005,TB006,TB007,TB008,TB009,TB010");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTC,[TK].dbo.COPTD");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.COPTB ON TB001=TD017 AND TB002=TD018 AND TB003=TD019 AND TB004=TD004 ");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND COPTD.MODIFIER='160115'");
                sbSql.AppendFormat(@"  AND TD002 LIKE '{0}%'",dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND (TD011<>ISNULL(TB009,0) OR  (TD008*TD011*TD026-TD012)<>0)  ");
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

                
                sbSql.AppendFormat(@"  SELECT COPTH.MODI_DATE,TH012,ISNULL(TD011,0) TD011,(TH008*TH012*TH025-TH013) AS DIFF,TG007,TH001,TH002,TH003,TH004,TH005,TH006,TH007,TH008,TH009,TH012,TH013,TH014,TH015,TH016,TH025,'',TD004,TD011");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTG,[TK].dbo.COPTH");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.COPTD ON TD001=TH014 AND TD002=TH015 AND TD003=TH016 AND TD004=TH004");
                sbSql.AppendFormat(@"  WHERE  TG001=TH001 AND TG002=TH002");
                sbSql.AppendFormat(@"  AND COPTH.MODIFIER='160115'");
                sbSql.AppendFormat(@"  AND COPTH.TH002 LIKE '{0}%'", dateTimePicker2.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND (TH012<>ISNULL(TD011,0) OR (TH008*TH012*TH025-TH013)<>0)    ");
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

                sbSql.AppendFormat(@"  SELECT TJ011,ISNULL(TH012,0) TH012,   (TJ007*TJ011*TH025-TJ012) AS DIFF,TI021,TJ001,TJ002,TJ003,TJ004,TJ005,TJ006,TJ007,TJ008,TJ011,TJ012,TJ015,TJ016,TJ017,TH004,TH012");
                sbSql.AppendFormat(@"  FROM [TK].dbo.COPTI,[TK].dbo.COPTJ");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.COPTH ON TH001=TJ015 AND TH002=TJ016 AND TH003=TJ017 AND TH004=TJ004");
                sbSql.AppendFormat(@"  WHERE  TI001=TJ001 AND TI002=TJ002");
                sbSql.AppendFormat(@"  AND COPTJ.MODIFIER='160115'");
                sbSql.AppendFormat(@"  AND COPTJ.TJ002 LIKE '{0}%'",dateTimePicker3.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND (TJ011<>ISNULL(TH012,0) OR   (TJ007*TJ011*TH025-TJ012)<>0)   ");
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

        public void Search4()
        {
            try
            {
                talbename = "TEMPds4";
                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT MA002,ISNULL(TEMP.TM010,0) AS TM010,TD010,(TD010-ISNULL(TEMP.TM010,0)) AS PDIFF,ISNULL(TB009,0) AS TB009,TD011,(TD011-TD008*TD010) AS DIFF,TC025,TD001,TD002,TD003,TD004,TD005,TD006,TD008,TD009 ,TD013,TD021,TD023     ");
                sbSql.AppendFormat(@"  FROM [TK].dbo.PURMA ,[TK].dbo.PURTC,[TK].dbo.PURTD ");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.PURTB ON TB004=TD004 AND TD013=TB001 AND TD021=TB002 AND TD023=TB003");
                sbSql.AppendFormat(@"  LEFT JOIN (SELECT TL004,TM004,TM010 FROM [TK].dbo.PURTL,[TK].dbo.PURTM WHERE TL001=TM001 AND TL002=TM002 AND TM011='Y') AS TEMP ON TEMP.TM004=TD004");
                sbSql.AppendFormat(@"  WHERE TC001=TD001 AND TC002=TD002");
                sbSql.AppendFormat(@"  AND MA001=TC004");
                sbSql.AppendFormat(@"  AND PURTD.MODIFIER='160115'");
                sbSql.AppendFormat(@"   AND (ISNULL(TEMP.TM010,0)<>TD010 OR ISNULL(TB009,0)<>TD011 OR (TD011-TD008*TD010)<>0)");
                sbSql.AppendFormat(@"  AND PURTD.TD002 LIKE '{0}%'",dateTimePicker4.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ORDER BY MA002");
                sbSql.AppendFormat(@"  ");

                textBox4.Text = null;
                textBox4.Text = sbSql.ToString();

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
                        dataGridView4.DataSource = ds.Tables[talbename];
                        dataGridView4.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView4.CurrentCell = dataGridView4.Rows[rownum].Cells[0];

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

        public void Search5()
        {
            try
            {
                talbename = "TEMPds5";
                sbSql.Clear();

                sbSql.AppendFormat(@"  SELECT TG021,(TH016*TH018-TH019) AS DIFF,TD010,TH018,TH001,TH002,TH003,TH004,TH005,TH006,TH007,TH008,TH011,TH012,TH013,TH014,TH015,TH016,TH017,TH018,TH019,(TH016*TH018-TH019) AS DIFF,TD010");
                sbSql.AppendFormat(@"  FROM [TK].dbo.PURTG,[TK].dbo.PURTH");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.PURTD ON TD004=TH004 AND TH011=TD001 AND TH012=TD002 AND TH013=TD003");
                sbSql.AppendFormat(@"  WHERE TG001=TH001 AND TG002=TH002");
                sbSql.AppendFormat(@"  AND PURTH.MODIFIER='160115'");
                sbSql.AppendFormat(@"  AND PURTH.TH002 LIKE '{0}%'",dateTimePicker5.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND (TH018<>TD010 OR (TH016*TH018-TH019) <>0)");
                sbSql.AppendFormat(@"  ORDER BY TG021");
                sbSql.AppendFormat(@"  ");

                textBox5.Text = null;
                textBox5.Text = sbSql.ToString();

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
                        dataGridView5.DataSource = ds.Tables[talbename];
                        dataGridView5.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView5.CurrentCell = dataGridView5.Rows[rownum].Cells[0];

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
        public void Search6()
        {
            try
            {
                talbename = "TEMPds6";
                sbSql.Clear();

             
                sbSql.AppendFormat(@"  SELECT TI016,TJ009,ISNULL(TH018,0) AS TH018,(TJ008*TJ009-TJ010) AS DIFF,TJ001,TJ002,TJ003,TJ004,TJ005,TJ006,TJ007,TJ008,TJ009,TJ010,TJ013,TJ014,TJ015");
                sbSql.AppendFormat(@"  FROM  [TK].dbo.PURTI,[TK].dbo.PURTJ");
                sbSql.AppendFormat(@"  LEFT JOIN [TK].dbo.PURTH ON TH004=TJ004 AND TJ013=TH001 AND TJ014=TH002 AND TJ015=TH003");
                sbSql.AppendFormat(@"  WHERE TI001=TJ001 AND TI002=TJ002");
                sbSql.AppendFormat(@"  AND PURTJ.TJ002 LIKE '{0}%'",dateTimePicker6.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND PURTH.MODIFIER='160115'");
                sbSql.AppendFormat(@"  AND (TJ009<>ISNULL(TH018,0) OR (TJ008*TJ009-TJ010)<>0)");
                sbSql.AppendFormat(@"  ORDER BY TI016");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                textBox6.Text = null;
                textBox6.Text = sbSql.ToString();

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
                        dataGridView6.DataSource = ds.Tables[talbename];
                        dataGridView6.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView6.CurrentCell = dataGridView6.Rows[rownum].Cells[0];

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
        private void button4_Click(object sender, EventArgs e)
        {
            Search4();
            MessageBox.Show("QUERY");
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Search5();
            MessageBox.Show("QUERY");
        }
        private void button6_Click(object sender, EventArgs e)
        {
            Search6();
            MessageBox.Show("QUERY");
        }

        #endregion


    }
}
