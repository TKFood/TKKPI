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
    public partial class frmECOMMERCE : Form
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

        public frmECOMMERCE()
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


            if(comboBox1.Text.ToString().Equals("電商每日銷貨金額及預估比較表"))
            {
                Queryday = dateTimePicker1.Value.ToString("yyyyMM");
                Queryday = Queryday + "01";

                STR.Append(@" SELECT ID ");
                STR.AppendFormat(@" ,CONVERT(varchar(6),'{0}',112)+ID AS '日期'", Queryday);
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(LA011) FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA  WITH (NOLOCK) WHERE TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND TH020='Y' AND   TH001='A233' AND SUBSTRING(TH002,1,8)=CONVERT(varchar(6),'{0}',112)+ID),0) AS INT) AS '日出貨量'", Queryday);
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(TH013) FROM [TK].dbo.COPTH WITH (NOLOCK) WHERE   TH020='Y' AND TH001='A233' AND SUBSTRING(TH002,1,8)=CONVERT(varchar(6),'{0}',112)+ID),0)AS INT) AS '日出貨金額'", Queryday); ;
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(LA011) FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA  WITH (NOLOCK) WHERE TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND TH020='Y' AND TH001='A233' AND SUBSTRING(TH002,1,6)=CONVERT(varchar(6),'{0}',112) AND   SUBSTRING(TH002,1,8)<=CONVERT(varchar(6),'{0}',112)+ID),0) AS INT) AS '累積出貨量'", Queryday);
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(TH013) FROM [TK].dbo.COPTH WITH (NOLOCK) WHERE TH020='Y' AND TH001='A233' AND SUBSTRING(TH002,1,6)=CONVERT(varchar(6),'{0}',112) AND   SUBSTRING(TH002,1,8)<=CONVERT(varchar(6),'{0}',112)+ID),0) AS INT)AS '累積出貨金額'", Queryday);
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(PREOrderNum)  FROM [TKECOMMERCE].[dbo].[ZTKECOMMERCEFrmMPRECOPTC] WHERE YEARMONTH=CONVERT(varchar(6),'{0}',112))/Day(dateadd(dd,-1,DATEADD(mm, DATEDIFF(m,0,'{0}')+1, 0)))*ID,0) AS INT)AS '預計出貨量'", Queryday);
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(PREOrderNum*MB047)  FROM [TKECOMMERCE].[dbo].[ZTKECOMMERCEFrmMPRECOPTC],[TK].dbo.[INVMB]  WHERE  INVMB.MB001=ZTKECOMMERCEFrmMPRECOPTC.MB001 AND  YEARMONTH=CONVERT(varchar(6),'{0}',112))/Day(dateadd(dd,-1,DATEADD(mm, DATEDIFF(m,0,'{0}')+1, 0)))*ID,0)AS INT) AS '預計出貨金額'", Queryday);
                STR.AppendFormat(@" ,CAST (ROUND((ISNULL((SELECT SUM(TH013) FROM [TK].dbo.COPTH WITH (NOLOCK) WHERE TH020='Y' AND TH001='A233' AND SUBSTRING(TH002,1,6)=CONVERT(varchar(6),'{0}',112) AND   SUBSTRING(TH002,1,8)<=CONVERT(varchar(6),'{0}',112)+ID),0))/(ISNULL((SELECT SUM(PREOrderNum*MB047)  FROM [TKECOMMERCE].[dbo].[ZTKECOMMERCEFrmMPRECOPTC],[TK].dbo.[INVMB]  WHERE  INVMB.MB001=ZTKECOMMERCEFrmMPRECOPTC.MB001 AND  YEARMONTH=CONVERT(varchar(6),'{0}',112))/Day(dateadd(dd,-1,DATEADD(mm, DATEDIFF(m,0,'{0}')+1, 0)))*ID,0))*100,2)  AS DECIMAL(18,2)) AS '完成率'", Queryday);
                STR.Append(@" FROM [TKECOMMERCE].dbo.BASEDAY");
                STR.AppendFormat(@" WHERE ID<=DAY(DATEADD(mm,  1, DATEADD(dd, -1, DATEADD(mm, DATEDIFF(mm,0,'{0}'), 0))))", Queryday);

                talbename = "TEMPds1";
            }
            else if(comboBox1.Text.ToString().Equals("電話訂購每日銷貨金額及預估比較表"))
            {
                Queryday = dateTimePicker1.Value.ToString("yyyyMM");
                Queryday = Queryday + "01";

                STR.Append(@" SELECT ID ");
                STR.AppendFormat(@" ,CONVERT(varchar(6),'{0}',112)+ID AS '日期'", Queryday);
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(LA011) FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA  WITH (NOLOCK) WHERE TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND TH020='Y' AND   TH001='A230' AND SUBSTRING(TH002,1,8)=CONVERT(varchar(6),'{0}',112)+ID),0) AS INT) AS '日出貨量'", Queryday);
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(TH013) FROM [TK].dbo.COPTH WITH (NOLOCK) WHERE   TH020='Y' AND TH001='A230' AND SUBSTRING(TH002,1,8)=CONVERT(varchar(6),'{0}',112)+ID),0)AS INT) AS '日出貨金額'", Queryday); ;
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(LA011) FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA  WITH (NOLOCK) WHERE TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND TH020='Y' AND TH001='A230' AND SUBSTRING(TH002,1,6)=CONVERT(varchar(6),'{0}',112) AND   SUBSTRING(TH002,1,8)<=CONVERT(varchar(6),'{0}',112)+ID),0) AS INT) AS '累積出貨量'", Queryday);
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(TH013) FROM [TK].dbo.COPTH WITH (NOLOCK) WHERE TH020='Y' AND TH001='A230' AND SUBSTRING(TH002,1,6)=CONVERT(varchar(6),'{0}',112) AND   SUBSTRING(TH002,1,8)<=CONVERT(varchar(6),'{0}',112)+ID),0) AS INT)AS '累積出貨金額'", Queryday);
                STR.Append(@" FROM [TKECOMMERCE].dbo.BASEDAY");
                STR.AppendFormat(@" WHERE ID<=DAY(DATEADD(mm,  1, DATEADD(dd, -1, DATEADD(mm, DATEDIFF(mm,0,'{0}'), 0))))", Queryday);

                talbename = "TEMPds2";
            }
            else if (comboBox1.Text.ToString().Equals("電話客服內容"))
            {
                Queryday = dateTimePicker1.Value.ToString("yyyyMM");
                

                STR.Append(@" SELECT CallDate AS '日期',[BASETYPE].TypeID AS '類別',TypeName AS '類名',CallName AS '姓名',CallPhone AS '手機',CallText AS '來電內容',CallTextRe AS '回覆',OrderID AS '訂單',ShipID AS '出貨單',InvoiceNo AS '發票'");
                STR.Append(@" FROM [TKCUSTOMERSERVICE].[dbo].[CALLRECORD]");
                STR.Append(@" LEFT JOIN [TKCUSTOMERSERVICE].[dbo].[BASETYPE] ON[CALLRECORD].[TypeID]=[BASETYPE].[TypeID]");
                STR.AppendFormat(@" WHERE SUBSTRING(CallDate,1,6)='{0}'", Queryday);
                STR.Append(@" ORDER BY CallDate,[BASETYPE].TypeID");

                talbename = "TEMPds3";
            }
            else if (comboBox1.Text.ToString().Equals("今年電商成長率"))
            {
                STR.Append(@" SELECT 月份,今年,CAST(今年出貨量 AS DECIMAL(18,2)) AS '今年出貨量'  ,CAST(今年退貨量 AS DECIMAL(18,2))  AS '今年退貨量',CAST((今年出貨量-今年退貨量) AS DECIMAL(18,2)) AS '今年實出量'  ,CAST(今年出貨金額 AS DECIMAL(18,2)) AS '今年出貨金額',CAST(今年退貨金額 AS DECIMAL(18,2)) AS '今年退貨金額',CAST((今年出貨金額-今年退貨金額) AS DECIMAL(18,2))AS '今年實出金額'  ,去年,CAST(去年出貨量 AS DECIMAL(18,2)) AS '去年出貨量',CAST(去年退貨量 AS DECIMAL(18,2)) AS '去年退貨量',CAST((去年出貨量-去年退貨量) AS DECIMAL(18,2)) AS '去年實出量'  ,CAST(去年出貨金額 AS DECIMAL(18,2))  AS '去年出貨金額',CAST(去年退貨金額 AS DECIMAL(18,2)) AS '去年退貨金額'  ,CAST((去年出貨金額-去年退貨金額) AS DECIMAL(18,2))  AS '去年實出金額' ");
                STR.Append(@" FROM (");
                STR.Append(@" SELECT ID AS '月份'");
                STR.Append(@" ,CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AS  '今年'");
                STR.Append(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH,[TK].dbo.INVLA WITH (NOLOCK) WHERE TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TH001='A233')  ,0) AS '今年出貨量'");
                STR.Append(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ,[TK].dbo.INVLA WITH (NOLOCK) WHERE TJ021 ='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001='A246') ,0)  AS '今年退貨量'");
                STR.Append(@" ,ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK) WHERE  TH020='Y' AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AND TH001='A233'),0) AS '今年出貨金額'");
                STR.Append(@" ,ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK) WHERE  TJ021='Y' AND SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001='A246'),0)  '今年退貨金額'");
                STR.Append(@" ,CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AS '去年'");
                STR.Append(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH,[TK].dbo.INVLA WITH (NOLOCK) WHERE   TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AND TH001='A233'),0) AS '去年出貨量'");
                STR.Append(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ,[TK].dbo.INVLA WITH (NOLOCK) WHERE TJ021='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001='A246'),0)  AS '去年退貨量'");
                STR.Append(@" ,ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK) WHERE TH020='Y' AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AND TH001='A233'),0)  AS '去年出貨金額'");
                STR.Append(@" ,ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK) WHERE  TJ021='Y' AND SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001='A246'),0) AS '去年退貨金額'");
                STR.Append(@" FROM [TKECOMMERCE].dbo.BASEMONTH ) AS TEMP");
                STR.Append(@" ");

                talbename = "TEMPds4";
            }
            else if(comboBox1.Text.ToString().Equals("今年電商累積銷貨"))
            {
        
                STR.Append(@" SELECT 月份,今年,CAST (今年累積出貨量 AS DECIMAL(18,2)) AS '今年累積出貨量',CAST (今年累積退貨量 AS DECIMAL(18,2)) AS '今年累積退貨量',CAST ((今年累積出貨量-今年累積退貨量) AS DECIMAL(18,2)) AS '今年實出貨量',CAST (今年累積出貨金額 AS DECIMAL(18,2))  AS '今年累積出貨金額',CAST (今年累積退貨金額 AS DECIMAL(18,2))  AS '今年累積退貨金額',CAST ((今年累積出貨金額-今年累積退貨金額) AS DECIMAL(18,2)) AS '今年實出金額',去年,CAST (去年累積出貨量 AS DECIMAL(18,2))  AS '去年累積出貨量',CAST (去年累積退貨量 AS DECIMAL(18,2)) AS '去年累積退貨量',CAST ((去年累積出貨量-去年累積退貨量) AS DECIMAL(18,2)) AS '去年實出貨量',去年,CAST(去年累積出貨金額 AS DECIMAL(18,2)) AS '去年累積出貨金額',CAST (去年累積退貨金額 AS DECIMAL(18,2)) AS '去年累積退貨金額',CAST ((去年累積出貨金額-去年累積退貨金額) AS DECIMAL(18,2)) AS '去年實出金額' ");
                STR.Append(@" FROM (");
                STR.Append(@" SELECT ID AS '月份'");
                STR.Append(@" ,CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AS  '今年'");
                STR.Append(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE  TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TH002,1,6)<=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TH001='A233'),0) ELSE 0 END AS '今年累積出貨量'");
                STR.Append(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE TJ021='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND   SUBSTRING(TJ002,1,6)<=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001='A246'),0)  ELSE 0 END AS '今年累積退貨量'");
                STR.Append(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK) WHERE  TH020='Y' AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TH002,1,6)<=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AND TH001='A233'),0) ELSE 0 END AS '今年累積出貨金額'");
                STR.Append(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK) WHERE  TJ021='Y' AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TJ002,1,6)<=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001='A246'),0)  ELSE 0 END AS '今年累積退貨金額'");
                STR.Append(@" ,CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AS '去年'");
                STR.Append(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE  TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TH002,1,6)<=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TH001='A233'),0) AS '去年累積出貨量'");
                STR.Append(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE  TJ021='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND   SUBSTRING(TJ002,1,6)<=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001='A246'),0)   AS '去年累積退貨量'");
                STR.Append(@" ,ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK) WHERE  TH020='Y' AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TH002,1,6)<=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AND TH001='A233'),0)  AS '去年累積出貨金額'");
                STR.Append(@" ,ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK) WHERE  TJ021='Y' AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TJ002,1,6)<=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001='A246'),0)  AS '去年累積退貨金額'");
                STR.Append(@" FROM [TKECOMMERCE].dbo.BASEMONTH ) AS TEMP");
                STR.Append(@" ");

                talbename = "TEMPds5";
            }
            else if (comboBox1.Text.ToString().Equals("本月客服統計"))
            {
                STR.Append(@" SELECT [CALLRECORD].[TypeID] AS TypeID,[BASETYPE].[TypeName] AS '名稱',COUNT([CALLRECORD].[TypeID]) AS '次數'");
                STR.Append(@" FROM [TKCUSTOMERSERVICE].[dbo].[CALLRECORD]");
                STR.Append(@" LEFT JOIN [TKCUSTOMERSERVICE].[dbo].[BASETYPE] ON[CALLRECORD].[TypeID]=[BASETYPE].[TypeID]");
                STR.AppendFormat(@" WHERE SUBSTRING(CallDate,1,6)='{0}'",dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@" GROUP BY [CALLRECORD].[TypeID],[BASETYPE].[TypeName]");
                STR.Append(@" ORDER BY COUNT([CALLRECORD].[TypeID]) DESC");
                STR.Append(@" ");

                talbename = "TEMPds6";
            }
            else if(comboBox1.Text.ToString().Equals("本月銷貨毛利"))
            {
                STR.Append(@" SELECT   品號,MB002 AS '品名',CAST (SUM(銷售數量) AS DECIMAL(18,2)) AS 銷售數量,CAST (SUM(銷售金額) AS DECIMAL(18,2)) AS 銷售金額,CAST (ISNULL(SUM(成本),0) AS DECIMAL(18,2)) AS '成本',CAST ((SUM(銷售金額)-ISNULL(SUM(成本),0)) AS DECIMAL(18,2)) AS '毛利'");
                STR.Append(@" FROM (");
                STR.Append(@" SELECT SUBSTRING(TH002,1,6) AS 'YM',TH004 AS '品號',LA011  AS '銷售數量',TH013 AS '銷售金額',LA013 AS '成本'");
                STR.Append(@" FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK)");
                STR.Append(@" WHERE TH020='Y' AND  TH001=LA006 AND TH002=LA007 AND TH003=LA008");
                STR.Append(@" AND SUBSTRING(TH002,1,6)=CONVERT(varchar(6),DATEADD(MONTH,-1, CONVERT(datetime, GETDATE())) , 112)");
                STR.Append(@" AND TH001='A233'");
                STR.Append(@" ) AS TEMP");
                STR.Append(@" LEFT JOIN [TK].dbo.INVMB ON MB001=品號");
                STR.Append(@" GROUP BY 品號,MB002");
                STR.Append(@" ORDER BY (SUM(銷售金額)-ISNULL(SUM(成本),0))  DESC");
                STR.Append(@" ");

                talbename = "TEMPds7";
            }
            else if (comboBox1.Text.ToString().Equals("銷貨明細"))
            {
                STR.Append(@" SELECT  品號,品名,CAST(SUM(銷售量) AS DECIMAL(18,2)) AS 銷售量,CAST(SUM(銷售金額) AS DECIMAL(18,2)) AS 銷售金額");
                STR.Append(@" FROM (");
                STR.Append(@" SELECT TH004  AS '品號',TH005  AS '品名',LA011 AS '銷售量',TH013 AS '銷售金額' ");
                STR.Append(@" FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK)");
                STR.Append(@" WHERE TH020='Y' AND  TH001=LA006 AND TH002=LA007 AND TH003=LA008");
                STR.AppendFormat(@" AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}'",dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                STR.Append(@" AND TH001='A233'");
                STR.Append(@"  ) AS TEMP");
                STR.Append(@" GROUP BY 品號,品名");
                STR.Append(@" ORDER BY SUM(銷售金額) DESC");
                STR.Append(@" ");

                talbename = "TEMPds8";
            }
        

            return STR;
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

        #endregion

    }
}
