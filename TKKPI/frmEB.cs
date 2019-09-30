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
    public partial class frmEB : Form
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

        public frmEB()
        {
            InitializeComponent();
            SETDATETIME();
        }

        #region FUNCTION
        public void SETDATETIME()
        {
            DateTime dt = DateTime.Now;
            DateTime startMonth = dt.AddDays(1 - dt.Day);

            dateTimePicker2.Value = startMonth;
        }
        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

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

        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder();
            string Queryday = null;

            if (comboBox1.Text.ToString().Equals("業績"))
            {
               
                STR.AppendFormat(@" SELECT *");
                STR.AppendFormat(@" FROM (");
                STR.AppendFormat(@" SELECT '1' AS 'SEQ','官網' AS 'KIND' ,CAST(SUM(TH037) AS INT) AS 'MONEY' FROM  [TK].dbo.COPTG,[TK].dbo.COPTH WITH (NOLOCK) WHERE  TG001=TH001 AND TG002=TH002 AND  TH001='A233' AND TH020='Y' AND TG005 IN ('102300','114000','116300') AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' ", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@" UNION ALL");
                STR.AppendFormat(@" SELECT '2' AS 'SEQ','現銷' AS 'KIND' ,CAST(SUM(TH037) AS INT) AS 'MONEY' FROM  [TK].dbo.COPTG,[TK].dbo.COPTH WITH (NOLOCK) WHERE  TG001=TH001 AND TG002=TH002 AND  TH001='A230' AND TH020='Y' AND TG005 IN ('102300','114000','116300') AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' ", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@" UNION ALL");
                STR.AppendFormat(@" SELECT '3' AS 'SEQ',MA002 AS 'KIND' ,CAST(SUM(TH037) AS INT) AS 'MONEY' FROM  [TK].dbo.COPTG,[TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.COPMA WHERE MA001=TG004 AND  TG001=TH001 AND TG002=TH002 AND  TH001='A234' AND TH020='Y' AND TG005 IN ('102300','114000','116300') AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' GROUP BY MA002  ", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@" ) AS TEMP  ORDER BY SEQ ");
                STR.AppendFormat(@" ");

                talbename = "TEMPds1";
            }
            else if (comboBox1.Text.ToString().Equals(""))
            {
                STR.AppendFormat(@" ");

                talbename = "TEMPds2";
            }
         
            

            return STR;
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
