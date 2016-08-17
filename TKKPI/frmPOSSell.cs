using System;
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
    public partial class frmPOSSell : Form
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

        public frmPOSSell()
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
                        rownum = ds.Tables[talbename].Rows.Count - 1;
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
            string ThisYear = null;
            string ThisMonth = null;
            string LastMonth = null;
            string LastYear = null;
            string LastYearMonth = null;

            ThisYear = dateTimePicker1.Value.ToString("yyyy");
            ThisMonth = dateTimePicker1.Value.ToString("MM");
            LastMonth = dt.AddMonths(-1).ToString("MM");
            LastYear = dt.AddYears(-1).ToString("yyyy");
            LastYearMonth = dt.AddYears(-1).AddMonths(1).ToString("MM");

            if (comboBox1.Text.ToString().Equals("方城市銷售門市"))
            {

                STR.Append(@"  SELECT '方城市銷售門市' AS 門市,row_number() over(order by SUM(NUM) desc) AS '排名'");
                STR.Append(@"  , 年月,TH004  AS '品號',TH005 AS '品名',MB003 AS '規格',MB004 AS '單位',CAST(SUM(NUM) AS INT) AS '銷售數量',CAST(SUM(MM)AS INT)  AS '銷售金額(含稅)',SUM(TOTALCOST) AS '製造成本',SUM(EARNMONEY) AS '毛利'");
                STR.Append(@"  FROM (");
                STR.Append(@"  SELECT SUBSTRING(TP001,1,6) AS '年月', TP004 TH004,MB002 TH005,MB003,MB004 ,TP008  NUM,TP021 MM,1 MD004,ISNULL((SELECT AVG(LB010) FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008  AS TOTALCOST,TP021-(ISNULL((SELECT LB010 FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008          )  AS EARNMONEY");
                STR.Append(@"  FROM [TK].dbo.POSTP WITH (NOLOCK)");
                STR.Append(@"  LEFT JOIN [TK].dbo.INVMB  WITH (NOLOCK) ON TP004=MB001");
                STR.AppendFormat(@"  WHERE   SUBSTRING(TP001,1,6)='{0}'",dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@"  AND TP002='106701'");
                STR.Append(@"  ) AS TEMP");
                STR.Append(@"  GROUP BY 年月,TH004,TH005,MB003,MB004");
                STR.Append(@"  ORDER BY SUM(NUM) DESC");
                STR.Append(@"  ");

                talbename = "TEMPds1";
            }
            else if (comboBox1.Text.ToString().Equals("方城市餐飲門市"))
            {

                STR.Append(@"  SELECT '方城市餐飲門市' AS 門市,row_number() over(order by SUM(NUM) desc) AS '排名'");
                STR.Append(@"  , 年月,TH004  AS '品號',TH005 AS '品名',MB003 AS '規格',MB004 AS '單位',CAST(SUM(NUM) AS INT) AS '銷售數量',CAST(SUM(MM)AS INT)  AS '銷售金額(含稅)',SUM(TOTALCOST) AS '製造成本',SUM(EARNMONEY) AS '毛利'");
                STR.Append(@"  FROM (");
                STR.Append(@"  SELECT SUBSTRING(TP001,1,6) AS '年月', TP004 TH004,MB002 TH005,MB003,MB004 ,TP008  NUM,TP021 MM,1 MD004,ISNULL((SELECT AVG(LB010) FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008  AS TOTALCOST,TP021-(ISNULL((SELECT LB010 FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008          )  AS EARNMONEY");
                STR.Append(@"  FROM [TK].dbo.POSTP WITH (NOLOCK)");
                STR.Append(@"  LEFT JOIN [TK].dbo.INVMB  WITH (NOLOCK) ON TP004=MB001");
                STR.AppendFormat(@"  WHERE   SUBSTRING(TP001,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@"  AND TP002='106702'");
                STR.Append(@"  ) AS TEMP");
                STR.Append(@"  GROUP BY 年月,TH004,TH005,MB003,MB004");
                STR.Append(@"  ORDER BY SUM(NUM) DESC");
                STR.Append(@"  ");

                talbename = "TEMPds2";
            }
            else if (comboBox1.Text.ToString().Equals("老楊五村"))
            {

                STR.Append(@"  SELECT '老楊五村' AS 門市,row_number() over(order by SUM(NUM) desc) AS '排名'");
                STR.Append(@"  , 年月,TH004  AS '品號',TH005 AS '品名',MB003 AS '規格',MB004 AS '單位',CAST(SUM(NUM) AS INT) AS '銷售數量',CAST(SUM(MM)AS INT)  AS '銷售金額(含稅)',SUM(TOTALCOST) AS '製造成本',SUM(EARNMONEY) AS '毛利'");
                STR.Append(@"  FROM (");
                STR.Append(@"  SELECT SUBSTRING(TP001,1,6) AS '年月', TP004 TH004,MB002 TH005,MB003,MB004 ,TP008  NUM,TP021 MM,1 MD004,ISNULL((SELECT AVG(LB010) FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008  AS TOTALCOST,TP021-(ISNULL((SELECT LB010 FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008          )  AS EARNMONEY");
                STR.Append(@"  FROM [TK].dbo.POSTP WITH (NOLOCK)");
                STR.Append(@"  LEFT JOIN [TK].dbo.INVMB  WITH (NOLOCK) ON TP004=MB001");
                STR.AppendFormat(@"  WHERE   SUBSTRING(TP001,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@"  AND TP002='111101'");
                STR.Append(@"  ) AS TEMP");
                STR.Append(@"  GROUP BY 年月,TH004,TH005,MB003,MB004");
                STR.Append(@"  ORDER BY SUM(NUM) DESC");
                STR.Append(@"  ");

                talbename = "TEMPds3";
            }
            else if (comboBox1.Text.ToString().Equals("站前四店"))
            {

                STR.Append(@"  SELECT '站前四店' AS 門市,row_number() over(order by SUM(NUM) desc) AS '排名'");
                STR.Append(@"  , 年月,TH004  AS '品號',TH005 AS '品名',MB003 AS '規格',MB004 AS '單位',CAST(SUM(NUM) AS INT) AS '銷售數量',CAST(SUM(MM)AS INT)  AS '銷售金額(含稅)',SUM(TOTALCOST) AS '製造成本',SUM(EARNMONEY) AS '毛利'");
                STR.Append(@"  FROM (");
                STR.Append(@"  SELECT SUBSTRING(TP001,1,6) AS '年月', TP004 TH004,MB002 TH005,MB003,MB004 ,TP008  NUM,TP021 MM,1 MD004,ISNULL((SELECT AVG(LB010) FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008  AS TOTALCOST,TP021-(ISNULL((SELECT LB010 FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008          )  AS EARNMONEY");
                STR.Append(@"  FROM [TK].dbo.POSTP WITH (NOLOCK)");
                STR.Append(@"  LEFT JOIN [TK].dbo.INVMB  WITH (NOLOCK) ON TP004=MB001");
                STR.AppendFormat(@"  WHERE   SUBSTRING(TP001,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@"  AND TP002='106504'");
                STR.Append(@"  ) AS TEMP");
                STR.Append(@"  GROUP BY 年月,TH004,TH005,MB003,MB004");
                STR.Append(@"  ORDER BY SUM(NUM) DESC");
                STR.Append(@"  ");

                talbename = "TEMPds4";
            }
            else if (comboBox1.Text.ToString().Equals("中山一店"))
            {

                STR.Append(@"  SELECT '中山一店' AS 門市,row_number() over(order by SUM(NUM) desc) AS '排名'");
                STR.Append(@"  , 年月,TH004  AS '品號',TH005 AS '品名',MB003 AS '規格',MB004 AS '單位',CAST(SUM(NUM) AS INT) AS '銷售數量',CAST(SUM(MM)AS INT)  AS '銷售金額(含稅)',SUM(TOTALCOST) AS '製造成本',SUM(EARNMONEY) AS '毛利'");
                STR.Append(@"  FROM (");
                STR.Append(@"  SELECT SUBSTRING(TP001,1,6) AS '年月', TP004 TH004,MB002 TH005,MB003,MB004 ,TP008  NUM,TP021 MM,1 MD004,ISNULL((SELECT AVG(LB010) FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008  AS TOTALCOST,TP021-(ISNULL((SELECT LB010 FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008          )  AS EARNMONEY");
                STR.Append(@"  FROM [TK].dbo.POSTP WITH (NOLOCK)");
                STR.Append(@"  LEFT JOIN [TK].dbo.INVMB  WITH (NOLOCK) ON TP004=MB001");
                STR.AppendFormat(@"  WHERE   SUBSTRING(TP001,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@"  AND TP002='106501'");
                STR.Append(@"  ) AS TEMP");
                STR.Append(@"  GROUP BY 年月,TH004,TH005,MB003,MB004");
                STR.Append(@"  ORDER BY SUM(NUM) DESC");
                STR.Append(@"  ");

                talbename = "TEMPds5";
            }
            else if (comboBox1.Text.ToString().Equals("民國二店"))
            {

                STR.Append(@"  SELECT '民國二店' AS 門市,row_number() over(order by SUM(NUM) desc) AS '排名'");
                STR.Append(@"  , 年月,TH004  AS '品號',TH005 AS '品名',MB003 AS '規格',MB004 AS '單位',CAST(SUM(NUM) AS INT) AS '銷售數量',CAST(SUM(MM)AS INT)  AS '銷售金額(含稅)',SUM(TOTALCOST) AS '製造成本',SUM(EARNMONEY) AS '毛利'");
                STR.Append(@"  FROM (");
                STR.Append(@"  SELECT SUBSTRING(TP001,1,6) AS '年月', TP004 TH004,MB002 TH005,MB003,MB004 ,TP008  NUM,TP021 MM,1 MD004,ISNULL((SELECT AVG(LB010) FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008  AS TOTALCOST,TP021-(ISNULL((SELECT LB010 FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008          )  AS EARNMONEY");
                STR.Append(@"  FROM [TK].dbo.POSTP WITH (NOLOCK)");
                STR.Append(@"  LEFT JOIN [TK].dbo.INVMB  WITH (NOLOCK) ON TP004=MB001");
                STR.AppendFormat(@"  WHERE   SUBSTRING(TP001,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@"  AND TP002='106502'");
                STR.Append(@"  ) AS TEMP");
                STR.Append(@"  GROUP BY 年月,TH004,TH005,MB003,MB004");
                STR.Append(@"  ORDER BY SUM(NUM) DESC");
                STR.Append(@"  ");

                talbename = "TEMPds6";
            }
            else if (comboBox1.Text.ToString().Equals("北港三店"))
            {

                STR.Append(@"  SELECT '北港三店' AS 門市,row_number() over(order by SUM(NUM) desc) AS '排名'");
                STR.Append(@"  , 年月,TH004  AS '品號',TH005 AS '品名',MB003 AS '規格',MB004 AS '單位',CAST(SUM(NUM) AS INT) AS '銷售數量',CAST(SUM(MM)AS INT)  AS '銷售金額(含稅)',SUM(TOTALCOST) AS '製造成本',SUM(EARNMONEY) AS '毛利'");
                STR.Append(@"  FROM (");
                STR.Append(@"  SELECT SUBSTRING(TP001,1,6) AS '年月', TP004 TH004,MB002 TH005,MB003,MB004 ,TP008  NUM,TP021 MM,1 MD004,ISNULL((SELECT AVG(LB010) FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008  AS TOTALCOST,TP021-(ISNULL((SELECT LB010 FROM [TK].dbo.INVLB WITH (NOLOCK) WHERE LB001=TP004 AND LB002=SUBSTRING(TP001,1,6)),0)*TP008          )  AS EARNMONEY");
                STR.Append(@"  FROM [TK].dbo.POSTP WITH (NOLOCK)");
                STR.Append(@"  LEFT JOIN [TK].dbo.INVMB  WITH (NOLOCK) ON TP004=MB001");
                STR.AppendFormat(@"  WHERE   SUBSTRING(TP001,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@"  AND TP002='106503'");
                STR.Append(@"  ) AS TEMP");
                STR.Append(@"  GROUP BY 年月,TH004,TH005,MB003,MB004");
                STR.Append(@"  ORDER BY SUM(NUM) DESC");
                STR.Append(@"  ");

                talbename = "TEMPds7";
            }




            return STR;
        }

        public void ExcelExport()
        {
            Search();

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds.Tables[talbename];
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
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
                ws.GetRow(j + 1).CreateCell(9).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString()));
                ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));


                j++;
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
            filename.AppendFormat(@"c:\temp\門市銷售{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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
