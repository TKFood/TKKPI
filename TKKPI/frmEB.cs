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
using TKITDLL;

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
            //try
            //{
            //    sbSql.Clear();
            //    sbSql = SETsbSql();

            //    if (!string.IsNullOrEmpty(sbSql.ToString()))
            //    {
            //        connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            //        sqlConn = new SqlConnection(connectionString);



            //        adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
            //        sqlCmdBuilder = new SqlCommandBuilder(adapter);

            //        sqlConn.Open();
            //        ds.Clear();
            //        adapter.Fill(ds, talbename);
            //        sqlConn.Close();

            //        label1.Text = "資料筆數:" + ds.Tables[talbename].Rows.Count.ToString();

            //        if (ds.Tables[talbename].Rows.Count == 0)
            //        {

            //        }
            //        else
            //        {
            //            dataGridView1.DataSource = ds.Tables[talbename];
            //            dataGridView1.AutoResizeColumns();
            //            //rownum = ds.Tables[talbename].Rows.Count - 1;
            //            dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

            //            //dataGridView1.CurrentCell = dataGridView1[0, 2];

            //        }
            //    }
            //    else
            //    {

            //    }



            //}
            //catch
            //{

            //}
            //finally
            //{

            //}

        }

        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder();
            string Queryday = null;

            if (comboBox1.Text.ToString().Equals("業績"))
            {              

                STR.AppendFormat(@" 
                                SELECT *
                                FROM (
                                SELECT '1' AS 'SEQ','官網' AS 'KIND' ,CAST(SUM(TH037) AS INT) AS 'MONEY' FROM  [TK].dbo.COPTG,[TK].dbo.COPTH WITH (NOLOCK) WHERE  TG001=TH001 AND TG002=TH002 AND  TH001 IN ('A233') AND TH020='Y' AND TG005 IN ('102300','114000','116300','117300') AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' 
                                UNION ALL
                                SELECT '2' AS 'SEQ','現銷' AS 'KIND' ,CAST(SUM(TH037) AS INT) AS 'MONEY' FROM  [TK].dbo.COPTG,[TK].dbo.COPTH WITH (NOLOCK) WHERE  TG001=TH001 AND TG002=TH002 AND  TH001 IN ('A230') AND TH020='Y' AND TG005 IN ('102300','114000','116300','117300') AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' 
                                UNION ALL
                                SELECT '3' AS 'SEQ','預購' AS 'KIND' ,CAST(SUM(TH037) AS INT) AS 'MONEY' FROM  [TK].dbo.COPTG,[TK].dbo.COPTH WITH (NOLOCK) WHERE  TG001=TH001 AND TG002=TH002 AND  TH001 IN ('A23E','A23F') AND TH020='Y' AND TG005 IN ('102300','114000','116300','117300') AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' 
                                UNION ALL
                                SELECT '99' AS 'SEQ',MA002 AS 'KIND' ,CAST(SUM(TH037) AS INT) AS 'MONEY' FROM  [TK].dbo.COPTG,[TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.COPMA WHERE MA001=TG004 AND  TG001=TH001 AND TG002=TH002 AND  TH001='A234' AND TH020='Y' AND TG005 IN ('102300','114000','116300','117300') AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' GROUP BY MA002  
                                ) AS TEMP  ORDER BY SEQ 
                                ", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));

                talbename = "TEMPds1";
            }
            else if (comboBox1.Text.ToString().Equals(""))
            {
                STR.AppendFormat(@" ");

                talbename = "TEMPds2";
            }
         
            

            return STR;
        }

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report4 = new Report();
            report4.Load(@"REPORT\業績-電商部.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report4.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            
            TableDataSource table = report4.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            //report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            //report1.SetParameterValue("P2", dateTimePicker2.Value.ToString("yyyyMMdd"));
            report4.Preview = previewControl1;
            report4.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();


            if (comboBox1.Text.ToString().Equals("業績"))
            {

                SB.AppendFormat(@" 
                                SELECT *
                                FROM (
                                SELECT '1' AS 'SEQ','官網' AS 'KIND' ,ISNULL(CAST(SUM(TH037) AS INT),0) AS 'MONEY' FROM  [TK].dbo.COPTG,[TK].dbo.COPTH WITH (NOLOCK) WHERE  TG001=TH001 AND TG002=TH002 AND  TH001 IN ('A233') AND TH020='Y' AND TG005 IN ('102300','114000','116300','117300') AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' 
                                UNION ALL
                                SELECT '2' AS 'SEQ','現銷' AS 'KIND' ,ISNULL(CAST(SUM(TH037) AS INT),0) AS 'MONEY' FROM  [TK].dbo.COPTG,[TK].dbo.COPTH WITH (NOLOCK) WHERE  TG001=TH001 AND TG002=TH002 AND  TH001 IN ('A230') AND TH020='Y' AND TG005 IN ('102300','114000','116300','117300') AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' 
                                UNION ALL
                                SELECT '3' AS 'SEQ','預購' AS 'KIND' ,ISNULL(CAST(SUM(TH037) AS INT),0) AS 'MONEY' FROM  [TK].dbo.COPTG,[TK].dbo.COPTH WITH (NOLOCK) WHERE  TG001=TH001 AND TG002=TH002 AND  TH001 IN ('A23E','A23F') AND TH020='Y' AND TG005 IN ('102300','114000','116300','117300') AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' 
                                UNION ALL
                                SELECT '4' AS 'SEQ','業績' AS 'KIND' ,ISNULL(CAST(SUM(TH037) AS INT),0) AS 'MONEY' FROM  [TK].dbo.COPTG,[TK].dbo.COPTH WITH (NOLOCK) WHERE  TG001=TH001 AND TG002=TH002 AND  TH001 NOT IN ('A230','A233','A23E','A23F') AND TH020='Y' AND TG005 IN ('102300','114000','116300','117300') AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' 

                                ) AS TEMP  ORDER BY SEQ 

                                ", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));

                talbename = "TEMPds1";
            }
            else if (comboBox1.Text.ToString().Equals(""))
            {
                SB.AppendFormat(@" ");

                talbename = "TEMPds2";
            }



            return SB;

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            //Search();
            SETFASTREPORT();
        }

        #endregion
    }
}
