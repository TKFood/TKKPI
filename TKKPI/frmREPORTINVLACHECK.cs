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
using System.Data.SQLite;

namespace TKKPI
{
    public partial class frmREPORTINVLACHECK : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        SqlTransaction tran;
        DataSet ds1 = new DataSet();
        int result;

        public frmREPORTINVLACHECK()
        {
            InitializeComponent();

            comboBox1load();
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
            Sequel.AppendFormat(@"SELECT MC001,MC002,MC001+MC002 AS MC0012 FROM [TK].dbo.CMSMC    WHERE MC001 LIKE '2%'  ORDER BY  MC001 ");

            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("MC001", typeof(string));
            dt.Columns.Add("MC002", typeof(string));
            dt.Columns.Add("MC0012", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "MC001";
            comboBox1.DisplayMember = "MC0012";
            sqlConn.Close();


        }

        public void ADDTBINVLACHECK()
        {
            int days = new TimeSpan(dateTimePicker2.Value.Ticks - dateTimePicker1.Value.Ticks).Days;
            days = days + 1;
            DateTime dt = new DateTime();
            dt = dateTimePicker1.Value;

            string MC001 = comboBox1.SelectedValue.ToString();

            DataTable DTMB001 = SEARSHMB001(textBox1.Text.Trim(), textBox2.Text.Trim());

            if(DTMB001.Rows.Count>0)
            {
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



                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();

                    sbSql.AppendFormat(@"  
                                       DELETE [TKKPI].[dbo].[TBINVLACHECK]
                                        ");


                    foreach (DataRow dr in DTMB001.Rows)
                    {
                        dt = dateTimePicker1.Value;
                        string MB001 = dr["MB001"].ToString().Trim(); 

                        for (int i = 0; i <= days; i++)
                        {

                            sbSql.AppendFormat(@"  
                                                INSERT INTO [TKKPI].[dbo].[TBINVLACHECK]
                                                ([SDATE],[LA009],[LA001],[MC002],[MB002],[NUMS])
                                                (
                                                SELECT '{0}',LA009,LA001,[MC002],MB002,ISNULL(SUM(LA005*LA011),0) AS 'NUMS'
                                                FROM [TK].dbo.INVLA WITH (NOLOCK) ,[TK].dbo.INVMB WITH (NOLOCK) ,[TK].dbo.CMSMC
                                                WHERE LA009=MC001
                                                AND LA001=MB001
                                                AND LA009='{2}'                                                
                                                AND LA001='{1}'
                                                AND LA004<='{0}'
                                                GROUP BY LA009,LA001,MB002,[MC002]  
                                                )
                                                ", dt.ToString("yyyyMMdd"), MB001, MC001);

                            dt = dt.AddDays(1);
                        }
                    }
                    


                    cmd.Connection = sqlConn;
                    cmd.CommandTimeout = 60;
                    cmd.CommandText = sbSql.ToString();
                    cmd.Transaction = tran;
                    result = cmd.ExecuteNonQuery();

                    if (result == 0)
                    {
                        tran.Rollback();    //交易取消

                        MessageBox.Show("失敗");
                    }
                    else
                    {
                        tran.Commit();      //執行交易  

                        MessageBox.Show("成功");
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
            
        }

        public DataTable SEARSHMB001(string SMB001,string EMB001)
        {
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


                sbSql.Clear();
                sbSqlQuery.Clear();
                                
                sbSql.AppendFormat(@"  
                                    SELECT MB001
                                    FROM [TK].dbo.INVMB
                                    WHERE MB001>='{0}' AND MB001<='{1}'
                                    ORDER BY MB001
                                    ", SMB001, EMB001);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["TEMPds1"];
                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {

            }
        }

        #endregion

        #region BUTTON
        private void button4_Click(object sender, EventArgs e)
        {
            ADDTBINVLACHECK();
        }
        #endregion

    }
}
