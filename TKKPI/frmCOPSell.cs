﻿using System;
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
using TKITDLL;

namespace TKKPI
{
    public partial class frmCOPSell : Form
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
       

        public frmCOPSell()
        {
            InitializeComponent();
        }

        #region FUNCTION
       

        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    ds.Tables.Clear();

                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);


                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                   
                    adapter.Fill(ds, tablename);
                    sqlConn.Close();

                    label1.Text = "資料筆數:" + ds.Tables[tablename].Rows.Count.ToString();

                    if (ds.Tables[tablename].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        dataGridView1.Columns.Clear();
                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                        //rownum = ds.Tables[tablename].Rows.Count - 1;
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
            DateTime dt = dateTimePicker1.Value;
            string ThisMonthDay = dt.ToString("yyyyMM") + "25";
            string LastMonthDay= dt.AddMonths(-1).ToString("yyyyMM") + "26";

            if (comboBox1.Text.ToString().Equals("業務業績表"))
            {
               
                STR.AppendFormat(@"  SELECT '{0}'  AS  '年月',MV002  AS '名稱' ", dt.ToString("yyyyMM"));
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(LA011) FROM [TK].dbo.COPTG,[TK].dbo.COPTH,[TK].dbo.INVLA WHERE TG001=TH001 AND TG002=TH002 AND LA006=TH001 AND LA007=TH002 AND LA008=TH003 AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}' AND TG006=MV001),0) AS INT) AS '銷貨數量' ",LastMonthDay,ThisMonthDay);
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(TH013) FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002  AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}'  AND TG006=MV001),0) AS INT) AS '銷貨金額' ", LastMonthDay, ThisMonthDay);
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(LA011) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ,[TK].dbo.INVLA WHERE TI001=TJ001 AND TI002=TJ002 AND LA006=TJ001 AND LA007=TJ002 AND LA008=TJ003 AND SUBSTRING(TJ002,1,8)>='{0}' AND SUBSTRING(TJ002,1,8)<='{1}'  AND TI006=MV001),0) AS INT) AS '銷退數量' ", LastMonthDay, ThisMonthDay);
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(TJ012) FROM [TK].dbo.COPTI,[TK].dbo.COPTJ WHERE TI001=TJ001 AND TI002=TJ002  AND SUBSTRING(TJ002,1,8)>='{0}' AND SUBSTRING(TJ002,1,8)<='{1}' AND TI006=MV001),0) AS INT) AS '銷退金額' ", LastMonthDay, ThisMonthDay); ;
                STR.Append(@"  FROM [TK].dbo.CMSMV     ");
                STR.Append(@"  WHERE MV001 IN ('070005','090002','140020','140049','140078')");
                STR.Append(@"  ORDER BY MV001");
                STR.Append(@"  ");

                tablename = "TEMPds1";
            }
            else if (comboBox1.Text.ToString().Equals("各月客戶銷售表"))
            {
                STR.AppendFormat(@"  SELECT DISTINCT '{0}' AS '年月', TG004 AS '客戶代號',TG007  AS '客戶名稱' ", dt.ToString("yyyyMM"));
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(LA.LA011) FROM [TK].dbo.COPTG TG WITH (NOLOCK),[TK].dbo.COPTH TH WITH (NOLOCK) ,[TK].dbo.INVLA LA WITH (NOLOCK)WHERE TG.TG001=TH.TH001 AND TG.TG002=TH.TH002 AND LA.LA006=TH.TH001 AND LA.LA007=TH.TH002 AND LA.LA008=TH.TH003 AND TG.TG004=TEMP.TG004 AND TG.TG007=TEMP.TG007  AND TG.TG006 IN (SELECT ID FROM [TKKPI].dbo.[SALESMAN]) AND SUBSTRING(TG.TG002,1,8)>='{0}' AND SUBSTRING(TG.TG002,1,8)<='{1}'),0)  AS DECIMAL(18,2)) AS '銷貨數量' ", LastMonthDay, ThisMonthDay);
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(TH.TH037+TH.TH038) FROM [TK].dbo.COPTG  TG WITH (NOLOCK),[TK].dbo.COPTH TH WITH (NOLOCK) WHERE TG.TG001=TH.TH001 AND TG.TG002=TH.TH002  AND TG.TG004=TEMP.TG004 AND TG.TG007=TEMP.TG007  AND TG.TG006 IN (SELECT ID FROM [TKKPI].dbo.[SALESMAN]) AND SUBSTRING(TG.TG002,1,8)>='{0}' AND SUBSTRING(TG.TG002,1,8)<='{1}'),0)  AS DECIMAL(18,2)) AS '銷貨金額' ", LastMonthDay, ThisMonthDay);
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(LA.LA011) FROM [TK].dbo.COPTI TI WITH (NOLOCK),[TK].dbo.COPTJ TJ WITH (NOLOCK),[TK].dbo.INVLA LA WITH (NOLOCK) WHERE TI.TI001=TJ.TJ001 AND TI.TI002=TJ.TJ002 AND LA.LA006=TJ.TJ001 AND LA.LA007=TJ.TJ002 AND LA.LA008=TJ.TJ003 AND TI.TI004=TEMP.TG004 AND TI.TI021=TEMP.TG007  AND TI.TI006 IN (SELECT ID FROM [TKKPI].dbo.[SALESMAN]) AND SUBSTRING(TI.TI002,1,8)>='{0}' AND SUBSTRING(TI.TI002,1,8)<='{1}'),0)  AS DECIMAL(18,2)) AS '銷退數量' ", LastMonthDay, ThisMonthDay);
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(TJ.TJ033+TJ.TJ034) FROM [TK].dbo.COPTI TI WITH (NOLOCK),[TK].dbo.COPTJ TJ WITH (NOLOCK)  WHERE TI.TI001=TJ.TJ001 AND TI.TI002=TJ.TJ002  AND TI.TI004=TEMP.TG004 AND TI.TI021=TEMP.TG007  AND TI.TI006 IN (SELECT ID FROM [TKKPI].dbo.[SALESMAN]) AND SUBSTRING(TI.TI002,1,8)>='{0}' AND SUBSTRING(TI.TI002,1,8)<='{1}'),0)   AS DECIMAL(18,2)) AS '銷退金額' ", LastMonthDay, ThisMonthDay); ;
                STR.Append(@"  FROM (");
                STR.Append(@"  SELECT TG004 ,TG007 FROM  [TK].dbo.COPTG  WITH (NOLOCK)");
                STR.AppendFormat(@"  WHERE  SUBSTRING(TG002,1,8)>='{0}' AND SUBSTRING(TG002,1,8)<='{1}'", LastMonthDay, ThisMonthDay); ;
                STR.Append(@"  AND TG006 IN (SELECT ID FROM [TKKPI].dbo.[SALESMAN])");
                STR.Append(@"  UNION ALL");
                STR.Append(@"  SELECT TI004,TI021 FROM  [TK].dbo.COPTI WITH (NOLOCK)");
                STR.AppendFormat(@"  WHERE  SUBSTRING(TI002,1,8)>='{0}' AND SUBSTRING(TI002,1,8)<='{1}'", LastMonthDay, ThisMonthDay);
                STR.Append(@"  AND TI006 IN (SELECT ID FROM [TKKPI].dbo.[SALESMAN]) ");
                STR.Append(@"  ) AS TEMP ");
                STR.Append(@"  ");

                tablename = "TEMPds2";
            }
            else if (comboBox1.Text.ToString().Equals("各月品號銷售表"))
            {
                STR.AppendFormat(@"  SELECT  DISTINCT '{0}' AS  '年月',TH004 AS  '品號',MB002 AS  '品名',MB003 AS  '規格'", dt.ToString("yyyyMM"));
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(LA.LA011) FROM [TK].dbo.COPTH  TH WITH (NOLOCK),[TK].dbo.INVLA LA WITH (NOLOCK) WHERE TH.TH001=LA.LA006 AND TH.TH002=LA.LA007 AND TH.TH003=LA.LA008 AND TH.TH004=TEMP.TH004 AND  SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}'),0)  AS DECIMAL(18,2)) AS '銷貨數量'", LastMonthDay, ThisMonthDay);
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(TH.TH037+TH.TH038) FROM [TK].dbo.COPTH  TH WITH (NOLOCK) WHERE  TH.TH004=TEMP.TH004 AND  SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}'),0)  AS DECIMAL(18,2)) AS '銷貨金額'", LastMonthDay, ThisMonthDay);
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(LA.LA011) FROM [TK].dbo.COPTJ  TJ WITH (NOLOCK),[TK].dbo.INVLA LA WITH (NOLOCK) WHERE TJ.TJ001=LA.LA006 AND TJ.TJ002=LA.LA007 AND TJ.TJ003=LA.LA008 AND TJ.TJ004=TEMP.TH004 AND  SUBSTRING(TJ002,1,8)>='{0}' AND SUBSTRING(TJ002,1,8)<='{1}'),0)  AS DECIMAL(18,2)) AS '銷退數量'", LastMonthDay, ThisMonthDay);
                STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(TJ.TJ033+TJ.TJ034) FROM [TK].dbo.COPTJ  TJ WITH (NOLOCK) WHERE  TJ.TJ004=TEMP.TH004 AND  SUBSTRING(TJ002,1,8)>='{0}' AND SUBSTRING(TJ002,1,8)<='{1}'),0) AS DECIMAL(18,2)) AS '銷退金額'", LastMonthDay, ThisMonthDay); 
                STR.Append(@"  FROM (");
                STR.Append(@"  SELECT TH004 FROM [TK].dbo.COPTH WITH (NOLOCK)");
                STR.AppendFormat(@"  WHERE SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}'", LastMonthDay, ThisMonthDay); ;
                STR.Append(@"  UNION ALL ");
                STR.Append(@"  SELECT TJ004 FROM [TK].dbo.COPTJ WITH (NOLOCK)");
                STR.AppendFormat(@"  WHERE SUBSTRING(TJ002,1,8)>='{0}' AND SUBSTRING(TJ002,1,8)<='{1}'", LastMonthDay, ThisMonthDay);
                STR.Append(@"  ) AS TEMP");
                STR.Append(@"  LEFT JOIN [TK].dbo.INVMB WITH (NOLOCK) ON MB001=TH004");
                STR.Append(@"  ");

                tablename = "TEMPds3";
            }

            return STR;
        }

        public void ExcelExport()
        {
            Search();

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds.Tables[tablename];
            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }


            int j = 0;
            if (tablename.Equals("TEMPds1"))
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString()));
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));

                    j++;
                }

            }
            else if (tablename.Equals("TEMPds2"))
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));

                    j++;
                }

            }
            else if (tablename.Equals("TEMPds3"))
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));

                    j++;
                }

            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\業務指標{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        #endregion
    }
}
