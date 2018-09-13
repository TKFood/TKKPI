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


            if (comboBox1.Text.ToString().Equals("電商每日銷貨金額及預估比較表"))
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
                STR.AppendFormat(@" WHERE ID<=DAY(DATEADD(day, -1, DATEADD(mm, DATEDIFF(mm, '', '{0}')+1, '')))", Queryday);

                talbename = "TEMPds1";
            }
            else if (comboBox1.Text.ToString().Equals("電話訂購每日銷貨金額及預估比較表"))
            {
                Queryday = dateTimePicker1.Value.ToString("yyyyMM");
                Queryday = Queryday + "01";

                STR.Append(@" SELECT ID ");
                STR.AppendFormat(@" ,CONVERT(varchar(6),'{0}',112)+ID AS '日期'", Queryday);
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(LA011) FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA  WITH (NOLOCK) WHERE TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND TH020='Y' AND   TH001='A230' AND SUBSTRING(TH002,1,8)=CONVERT(varchar(6),'{0}',112)+ID),0) AS INT) AS '日出貨量'", Queryday);
                STR.AppendFormat(@" ,CAST(ISNULL((SELECT SUM(TH013) FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK) WHERE TG001=TH001 AND TG002=TH002 AND TG005 IN ('106400','102300','114000') AND TH020='Y' AND TH001='A230' AND SUBSTRING(TH002,1,6)=CONVERT(varchar(6),'{0}',112) AND   SUBSTRING(TH002,1,8)<=CONVERT(varchar(6),'{0}',112)+ID),0) AS INT)AS '累積出貨金額'", Queryday);
                STR.Append(@" FROM [TKECOMMERCE].dbo.BASEDAY");
                STR.AppendFormat(@" WHERE ID<=DAY(DATEADD(day, -1, DATEADD(mm, DATEDIFF(mm, '', '{0}')+1, '')))", Queryday);

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
            else if (comboBox1.Text.ToString().Equals("今年電商累積銷貨"))
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
                STR.AppendFormat(@" WHERE SUBSTRING(CallDate,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@" GROUP BY [CALLRECORD].[TypeID],[BASETYPE].[TypeName]");
                STR.Append(@" ORDER BY COUNT([CALLRECORD].[TypeID]) DESC");
                STR.Append(@" ");

                talbename = "TEMPds6";
            }
            else if (comboBox1.Text.ToString().Equals("本月銷貨毛利"))
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
            else if (comboBox1.Text.ToString().Equals("官網銷貨明細"))
            {
                if (!string.IsNullOrEmpty(textBox1.Text.ToString()))
                {
                    STR.Append(@" SELECT  品號,品名,CAST(SUM(銷售量) AS DECIMAL(18,2)) AS 銷售量,CAST(SUM(銷售金額) AS DECIMAL(18,2)) AS 銷售金額");
                    STR.Append(@" FROM (");
                    STR.Append(@" SELECT TH004  AS '品號',TH005  AS '品名',LA011 AS '銷售量',TH013 AS '銷售金額' ");
                    STR.Append(@" FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK)");
                    STR.Append(@" WHERE TH020='Y' AND  TH001=LA006 AND TH002=LA007 AND TH003=LA008");
                    STR.AppendFormat(@" AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}'", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                    STR.Append(@" AND TH001='A233'");
                    STR.AppendFormat(@" AND (TH004 LIKE '%{0}%' OR TH005 LIKE '%{0}%')", textBox1.Text.ToString());
                    STR.Append(@"  ) AS TEMP");
                    STR.Append(@" GROUP BY 品號,品名");
                    STR.Append(@" ORDER BY SUM(銷售金額) DESC");
                    STR.Append(@" ");

                }
                else
                {
                    STR.Append(@" SELECT  品號,品名,CAST(SUM(銷售量) AS DECIMAL(18,2)) AS 銷售量,CAST(SUM(銷售金額) AS DECIMAL(18,2)) AS 銷售金額");
                    STR.Append(@" FROM (");
                    STR.Append(@" SELECT TH004  AS '品號',TH005  AS '品名',LA011 AS '銷售量',TH013 AS '銷售金額' ");
                    STR.Append(@" FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK)");
                    STR.Append(@" WHERE TH020='Y' AND  TH001=LA006 AND TH002=LA007 AND TH003=LA008");
                    STR.AppendFormat(@" AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}'", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                    STR.Append(@" AND TH001='A233'");
                    STR.Append(@"  ) AS TEMP");
                    STR.Append(@" GROUP BY 品號,品名");
                    STR.Append(@" ORDER BY SUM(銷售金額) DESC");
                    STR.Append(@" ");
                }


                talbename = "TEMPds8";
            }
            else if (comboBox1.Text.ToString().Equals("今年現銷累積銷貨"))
            {

                STR.Append(@" SELECT 月份,今年,CAST (今年累積出貨量 AS DECIMAL(18,2)) AS '今年累積出貨量',CAST (今年累積退貨量 AS DECIMAL(18,2)) AS '今年累積退貨量',CAST ((今年累積出貨量-今年累積退貨量) AS DECIMAL(18,2)) AS '今年實出貨量',CAST (今年累積出貨金額 AS DECIMAL(18,2))  AS '今年累積出貨金額',CAST (今年累積退貨金額 AS DECIMAL(18,2))  AS '今年累積退貨金額',CAST ((今年累積出貨金額-今年累積退貨金額) AS DECIMAL(18,2)) AS '今年實出金額',去年,CAST (去年累積出貨量 AS DECIMAL(18,2))  AS '去年累積出貨量',CAST (去年累積退貨量 AS DECIMAL(18,2)) AS '去年累積退貨量',CAST ((去年累積出貨量-去年累積退貨量) AS DECIMAL(18,2)) AS '去年實出貨量',去年,CAST(去年累積出貨金額 AS DECIMAL(18,2)) AS '去年累積出貨金額',CAST (去年累積退貨金額 AS DECIMAL(18,2)) AS '去年累積退貨金額',CAST ((去年累積出貨金額-去年累積退貨金額) AS DECIMAL(18,2)) AS '去年實出金額' ");
                STR.Append(@" FROM (");
                STR.Append(@" SELECT ID AS '月份'");
                STR.Append(@" ,CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AS  '今年'");
                STR.Append(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)  WHERE TG001=TH001 AND TG002=TH002 AND  TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TH002,1,6)<=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TH001='A230' AND TG005 IN ('102300','114000')          ),0) ELSE 0 END AS '今年累積出貨量'");
                STR.Append(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTI WITH (NOLOCK)  WHERE TI001=TJ001 AND TI002=TJ002 AND TJ021='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND   SUBSTRING(TJ002,1,6)<=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001='A240'  AND TI005 IN ('102300','114000')  ),0)  ELSE 0 END AS '今年累積退貨量'");
                STR.Append(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)  WHERE TG001=TH001 AND TG002=TH002 AND   TH020='Y' AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TH002,1,6)<=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AND TH001='A230'  AND TG005 IN ('102300','114000') ),0) ELSE 0 END AS '今年累積出貨金額'");
                STR.Append(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.COPTI WITH (NOLOCK)  WHERE TI001=TJ001 AND TI002=TJ002 AND  TJ021='Y' AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TJ002,1,6)<=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001='A240' AND TI005 IN ('102300','114000')  ),0)  ELSE 0 END AS '今年累積退貨金額'");
                STR.Append(@" ,CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AS '去年'");
                STR.Append(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)  WHERE TG001=TH001 AND TG002=TH002 AND  TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TH002,1,6)<=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TH001='A230'  AND TG005 IN ('102300','114000') ),0) AS '去年累積出貨量'");
                STR.Append(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTI WITH (NOLOCK)  WHERE TI001=TJ001 AND TI002=TJ002 AND  TJ021='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND   SUBSTRING(TJ002,1,6)<=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001='A240' AND TI005 IN ('102300','114000') ),0)   AS '去年累積退貨量'");
                STR.Append(@" ,ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)  WHERE TG001=TH001 AND TG002=TH002 AND  TH020='Y' AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TH002,1,6)<=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AND TH001='A230' AND TG005 IN ('102300','114000') ),0)  AS '去年累積出貨金額'");
                STR.Append(@" ,ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.COPTI WITH (NOLOCK)  WHERE TI001=TJ001 AND TI002=TJ002 AND  TJ021='Y' AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TJ002,1,6)<=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001='A240' AND TI005 IN ('102300','114000') ),0)  AS '去年累積退貨金額'");
                STR.Append(@" FROM [TKECOMMERCE].dbo.BASEMONTH ) AS TEMP");
                STR.Append(@" ");

                talbename = "TEMPds9";
            }
            else if (comboBox1.Text.ToString().Equals("會計消貨者業績"))
            {
                STR.AppendFormat(@" DECLARE @YEARS varchar(4)= '{0}' ;",dateTimePicker1.Value.ToString("yyyy"));
                STR.AppendFormat(@" SELECT 今年,月份,(總合-橘點子) '總計',(WEB-蝦皮-YAHOO-PCHOME-MOMO-神坊-樂天-自游邦) AS '官網',(總合-橘點子-WEB-蝦皮-YAHOO-PCHOME-MOMO-神坊-樂天-自游邦) AS '消費者',蝦皮,MOMO,YAHOO,PCHOME,神坊,樂天,自游邦 ");
                STR.AppendFormat(@" FROM ( ");
                STR.AppendFormat(@" SELECT ID AS '月份' ,CAST(YEAR(@YEARS) AS NVARCHAR)+ID  AS  '今年' ");
                STR.AppendFormat(@" ,CONVERT(INT,ISNULL((SELECT SUM(TB004*TB007)*-1 FROM [TK].[dbo].[ACTTB] WITH (NOLOCK), [TK].[dbo].[ACTTA] WITH (NOLOCK) WHERE TA001=TB001 AND TA002=TB002 AND TA010='Y' AND TB005='411104' AND TB006 NOT IN ('106200') AND TB002 LIKE CAST(YEAR(@YEARS) AS NVARCHAR)+ID+'%'),0)) AS '總合' ");
                STR.AppendFormat(@" ,CONVERT(INT,ISNULL((SELECT SUM(TB019) FROM [TK].[dbo].[ACRTB] WITH (NOLOCK) ,[TK].[dbo].[ACRTA] WITH (NOLOCK) WHERE TA001=TB001 AND TA002=TB002 AND TA004 IN ('2209400100','11110775') AND TB002 LIKE CAST(YEAR(@YEARS) AS NVARCHAR)+ID+'%'),0))  AS '橘點子' ");
                STR.AppendFormat(@" ,CONVERT(INT,ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK) WHERE  TH020='Y' AND TH005 NOT LIKE '%手續%'AND TH005 NOT LIKE '%運費%' AND (TH001='A233' OR TH001='A234') AND TH002 LIKE  CAST(YEAR(@YEARS) AS NVARCHAR)+ID+'%'  ),0)) 'WEB' ");
                STR.AppendFormat(@" ,CONVERt(INT,ISNULL((SELECT SUM(TB004*TB007)*-1 FROM [TK].[dbo].[ACTTB] WITH (NOLOCK) WHERE TB005='411104' AND TB006 IN ('106400','102300','114000') AND (TB010 LIKE '%蝦皮%')  AND TB002 LIKE CAST(YEAR(@YEARS) AS NVARCHAR)+ID+'%'),0))  AS '蝦皮' ");
                STR.AppendFormat(@" ,CONVERT(INT,ISNULL((SELECT SUM(TB004*TB007)*-1 FROM [TK].[dbo].[ACTTB] WITH (NOLOCK) WHERE TB005='411104' AND TB006 IN ('106400','102300','114000') AND (TB010 LIKE '%YAHOO%' OR TB010 LIKE '%yahoo%')  AND TB002 LIKE CAST(YEAR(@YEARS) AS NVARCHAR)+ID+'%'),0))  AS 'YAHOO' ");
                STR.AppendFormat(@" ,CONVERT(INT,ISNULL((SELECT SUM(TB004*TB007)*-1 FROM [TK].[dbo].[ACTTB] WITH (NOLOCK) WHERE TB005='411104' AND TB006 IN ('106400','102300','114000') AND (TB010 LIKE '%MOMO%')  AND TB002 LIKE CAST(YEAR(@YEARS) AS NVARCHAR)+ID+'%'),0))  AS 'MOMO'                 ");
                STR.AppendFormat(@" ,CONVERT(INT,ISNULL((SELECT SUM(TB004*TB007)*-1 FROM [TK].[dbo].[ACTTB] WITH (NOLOCK) WHERE TB005='411104' AND TB006 IN ('106400','102300','114000') AND (TB010 LIKE '%pc%' OR  TB010 LIKE '%Pc%'  OR  TB010 LIKE '%PC%' ) AND TB002 LIKE CAST(YEAR(@YEARS) AS NVARCHAR)+ID+'%'),0)) AS 'PCHOME' ");
                STR.AppendFormat(@" ,CONVERT(INT,ISNULL((SELECT SUM(TB004*TB007)*-1 FROM [TK].[dbo].[ACTTB] WITH (NOLOCK) WHERE TB005='411104' AND TB006 IN ('106400','102300','114000') AND (TB010 LIKE '%神坊%')  AND TB002 LIKE CAST(YEAR(@YEARS) AS NVARCHAR)+ID+'%'),0)) AS '神坊' ");
                STR.AppendFormat(@" ,CONVERT(INT,ISNULL((SELECT SUM(TB004*TB007)*-1 FROM [TK].[dbo].[ACTTB] WITH (NOLOCK) WHERE TB005='411104' AND TB006 IN ('106400','102300','114000') AND (TB010 LIKE '%樂天%')  AND TB002 LIKE CAST(YEAR(@YEARS) AS NVARCHAR)+ID+'%'),0))  AS '樂天' ");
                STR.AppendFormat(@" ,CONVERT(INT,ISNULL((SELECT SUM(TB004*TB007)*-1 FROM [TK].[dbo].[ACTTB] WITH (NOLOCK) WHERE TB005='411104' AND TB006 IN ('106400','102300','114000') AND (TB010 LIKE '%自游邦%')  AND TB002 LIKE CAST(YEAR(@YEARS) AS NVARCHAR)+ID+'%'),0))  AS '自游邦' ");
                STR.AppendFormat(@"  FROM [TKECOMMERCE].dbo.BASEMONTH ) AS TEMP");
                STR.AppendFormat(@" ");
                STR.AppendFormat(@" ");

                talbename = "TEMPds10";
            }
            else if (comboBox1.Text.ToString().Equals("該月平均每筆銷售未稅金額"))
            {
                STR.AppendFormat(@" SELECT ROUND(AVG(TG045),2) AS '未稅金額' FROM [TK].dbo.COPTG WITH (NOLOCK)");
                STR.AppendFormat(@" WHERE TG001='A233'");
                STR.AppendFormat(@" AND TG002 LIKE '{0}%'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.AppendFormat(@" ");

                talbename = "TEMPds11";
            }
            else if(comboBox1.Text.ToString().Equals("每月官網銷貨"))
            {

                STR.AppendFormat(@" SELECT 月份,今年");
                STR.AppendFormat(@" ,CAST ((本月出貨金額-本月退貨金額) AS DECIMAL(18,2)) AS '本月實出金額'");
                STR.AppendFormat(@" ,CAST ((本月出貨量-本月退貨量) AS DECIMAL(18,2)) AS '本月實出貨量'");
                STR.AppendFormat(@" ,CAST (本月出貨量 AS DECIMAL(18,2)) AS '本月出貨量'");
                STR.AppendFormat(@" ,CAST (本月退貨量 AS DECIMAL(18,2)) AS '本月退貨量'");
                STR.AppendFormat(@" ,CAST (本月出貨金額 AS DECIMAL(18,2))  AS '本月出貨金額'");
                STR.AppendFormat(@" ,CAST (本月退貨金額 AS DECIMAL(18,2))  AS '本月退貨金額'");
                STR.AppendFormat(@" ,去年");
                STR.AppendFormat(@" ,CAST ((去年本月出貨金額-去年本月退貨金額) AS DECIMAL(18,2)) AS '去年本月實出金額'");
                STR.AppendFormat(@" ,CAST ((去年本月出貨量-去年本月退貨量) AS DECIMAL(18,2)) AS '去年本月實出貨量'");
                STR.AppendFormat(@" ,CAST (去年本月出貨量 AS DECIMAL(18,2))  AS '去年本月出貨量'");
                STR.AppendFormat(@" ,CAST (去年本月退貨量 AS DECIMAL(18,2)) AS '去年本月退貨量'");
                STR.AppendFormat(@" ,CAST(去年本月出貨金額 AS DECIMAL(18,2)) AS '去年本月出貨金額'");
                STR.AppendFormat(@" ,CAST (去年本月退貨金額 AS DECIMAL(18,2)) AS '去年本月退貨金額'");
                STR.AppendFormat(@" FROM (");
                STR.AppendFormat(@" SELECT ID AS '月份' ,CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AS  '今年' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE  TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TH001='A233'),0) ELSE 0 END AS '本月出貨量' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE TJ021='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND   SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001='A246'),0)  ELSE 0 END AS '本月退貨量' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK) WHERE  TH020='Y' AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AND TH001='A233'),0) ELSE 0 END AS '本月出貨金額' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK) WHERE  TJ021='Y' AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001='A246'),0)  ELSE 0 END AS '本月退貨金額' ");
                STR.AppendFormat(@" ,CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AS '去年' ");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE  TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TH002,1,6)<=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TH001='A233'),0) AS '去年本月出貨量' ");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE  TJ021='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND   SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001='A246'),0)   AS '去年本月退貨量' ");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK) WHERE  TH020='Y' AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AND TH001='A233'),0)  AS '去年本月出貨金額' ");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK) WHERE  TJ021='Y' AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001='A246'),0)  AS '去年本月退貨金額' ");
                STR.AppendFormat(@"  FROM [TKECOMMERCE].dbo.BASEMONTH ");
                STR.AppendFormat(@"  ) AS TEMP ");
                STR.AppendFormat(@" ");
                STR.AppendFormat(@" ");
                STR.AppendFormat(@" ");

                talbename = "TEMPds12";
            }
            else if (comboBox1.Text.ToString().Equals("每月現銷銷貨"))
            {
                STR.AppendFormat(@" SELECT 月份,今年");
                STR.AppendFormat(@" ,CAST ((本月出貨金額-本月退貨金額) AS DECIMAL(18,2)) AS '本月實出金額'");
                STR.AppendFormat(@" ,CAST ((本月出貨量-本月退貨量) AS DECIMAL(18,2)) AS '本月實出貨量'");
                STR.AppendFormat(@" ,CAST (本月出貨金額 AS DECIMAL(18,2))  AS '本月出貨金額'");
                STR.AppendFormat(@" ,CAST (本月退貨金額 AS DECIMAL(18,2))  AS '本月退貨金額'");
                STR.AppendFormat(@" ,CAST (本月出貨量 AS DECIMAL(18,2)) AS '本月出貨量'");
                STR.AppendFormat(@" ,CAST (本月退貨量 AS DECIMAL(18,2)) AS '本月退貨量'");
                STR.AppendFormat(@" ,去年");
                STR.AppendFormat(@" ,CAST ((去年本月出貨金額-去年本月退貨金額) AS DECIMAL(18,2)) AS '去年本月實出金額' ");
                STR.AppendFormat(@" ,CAST ((去年本月出貨量-去年本月退貨量) AS DECIMAL(18,2)) AS '去年本月實出貨量'");
                STR.AppendFormat(@" ,CAST(去年本月出貨金額 AS DECIMAL(18,2)) AS '去年本月出貨金額'");
                STR.AppendFormat(@" ,CAST (去年本月退貨金額 AS DECIMAL(18,2)) AS '去年本月退貨金額'");
                STR.AppendFormat(@" ,CAST (去年本月出貨量 AS DECIMAL(18,2))  AS '去年本月出貨量'");
                STR.AppendFormat(@" ,CAST (去年本月退貨量 AS DECIMAL(18,2)) AS '去年本月退貨量'");
                STR.AppendFormat(@" FROM ( ");
                STR.AppendFormat(@" SELECT ID AS '月份' ");
                STR.AppendFormat(@" ,CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AS  '今年' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)  WHERE TG001=TH001 AND TG002=TH002 AND  TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TH001='A230' AND TG005 IN ('102300','114000','116300')          ),0) ELSE 0 END AS '本月出貨量' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTI WITH (NOLOCK)  WHERE TI001=TJ001 AND TI002=TJ002 AND TJ021='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND   SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001='A240'  AND TI005 IN ('102300','114000','116300')   ),0)  ELSE 0 END AS '本月退貨量' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)  WHERE TG001=TH001 AND TG002=TH002 AND   TH020='Y' AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AND TH001='A230'  AND TG005 IN ('102300','114000','116300')  ),0) ELSE 0 END AS '本月出貨金額' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.COPTI WITH (NOLOCK)  WHERE TI001=TJ001 AND TI002=TJ002 AND  TJ021='Y' AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001='A240' AND TI005 IN ('102300','114000','116300')   ),0)  ELSE 0 END AS '本月退貨金額' ");
                STR.AppendFormat(@" ,CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AS '去年' ");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)  WHERE TG001=TH001 AND TG002=TH002 AND  TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TH001='A230'  AND TG005 IN ('102300','114000','116300') ),0) AS '去年本月出貨量' ");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTI WITH (NOLOCK)  WHERE TI001=TJ001 AND TI002=TJ002 AND  TJ021='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND   SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001='A240' AND TI005 IN ('102300','114000','116300')  ),0)   AS '去年本月退貨量'");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)  WHERE TG001=TH001 AND TG002=TH002 AND  TH020='Y' AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AND TH001='A230' AND TG005 IN ('102300','114000','116300')  ),0)  AS '去年本月出貨金額' ");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.COPTI WITH (NOLOCK)  WHERE TI001=TJ001 AND TI002=TJ002 AND  TJ021='Y' AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001='A240' AND TI005 IN ('102300','114000','116300')  ),0)  AS '去年本月退貨金額' ");
                STR.AppendFormat(@" FROM [TKECOMMERCE].dbo.BASEMONTH");
                STR.AppendFormat(@" ) AS TEMP ");
                STR.AppendFormat(@" ");
                STR.AppendFormat(@" ");


                talbename = "TEMPds13";
            }
            else if (comboBox1.Text.ToString().Equals("現銷銷貨明細"))
            {
                if (!string.IsNullOrEmpty(textBox1.Text.ToString()))
                {
                    STR.Append(@" SELECT  品號,品名,CAST(SUM(銷售量) AS DECIMAL(18,2)) AS 銷售量,CAST(SUM(銷售金額) AS DECIMAL(18,2)) AS 銷售金額");
                    STR.Append(@" FROM (");
                    STR.Append(@" SELECT TH004  AS '品號',TH005  AS '品名',LA011 AS '銷售量',TH013 AS '銷售金額' ");
                    STR.Append(@" FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)");
                    STR.Append(@" WHERE TH020='Y' AND  TH001=LA006 AND TH002=LA007 AND TH003=LA008");
                    STR.AppendFormat(@" AND TG001=TH001 AND TG002=TH002");
                    STR.AppendFormat(@" AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}'", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                    STR.Append(@" AND TH001='A230'  AND TG005 IN ('102300','114000')  ");
                    STR.AppendFormat(@" AND (TH004 LIKE '%{0}%' OR TH005 LIKE '%{0}%')", textBox1.Text.ToString());
                    STR.Append(@"  ) AS TEMP");
                    STR.Append(@" GROUP BY 品號,品名");
                    STR.Append(@" ORDER BY SUM(銷售金額) DESC");
                    STR.AppendFormat(@" ");


                }
                else
                {
                    STR.Append(@" SELECT  品號,品名,CAST(SUM(銷售量) AS DECIMAL(18,2)) AS 銷售量,CAST(SUM(銷售金額) AS DECIMAL(18,2)) AS 銷售金額");
                    STR.Append(@" FROM (");
                    STR.Append(@" SELECT TH004  AS '品號',TH005  AS '品名',LA011 AS '銷售量',TH013 AS '銷售金額' ");
                    STR.Append(@" FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)");
                    STR.Append(@" WHERE TH020='Y' AND  TH001=LA006 AND TH002=LA007 AND TH003=LA008");
                    STR.AppendFormat(@" AND TG001=TH001 AND TG002=TH002");
                    STR.AppendFormat(@" AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}'", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                    STR.Append(@" AND TH001='A230'  AND TG005 IN ('102300','114000')  ");
                    STR.Append(@"  ) AS TEMP");
                    STR.Append(@" GROUP BY 品號,品名");
                    STR.Append(@" ORDER BY SUM(銷售金額) DESC");
                    STR.AppendFormat(@" ");
                }


                talbename = "TEMPds8";
            }

            else if (comboBox1.Text.ToString().Equals("每月電商銷貨"))
            {

                STR.AppendFormat(@" SELECT 月份,今年");
                STR.AppendFormat(@" ,CAST ((本月出貨金額-本月退貨金額) AS DECIMAL(18,2)) AS '本月實出金額'");
                STR.AppendFormat(@" ,CAST ((本月出貨量-本月退貨量) AS DECIMAL(18,2)) AS '本月實出貨量'");
                STR.AppendFormat(@" ,CAST (本月出貨量 AS DECIMAL(18,2)) AS '本月出貨量'");
                STR.AppendFormat(@" ,CAST (本月退貨量 AS DECIMAL(18,2)) AS '本月退貨量'");
                STR.AppendFormat(@" ,CAST (本月出貨金額 AS DECIMAL(18,2))  AS '本月出貨金額'");
                STR.AppendFormat(@" ,CAST (本月退貨金額 AS DECIMAL(18,2))  AS '本月退貨金額'");
                STR.AppendFormat(@" ,去年");
                STR.AppendFormat(@" ,CAST ((去年本月出貨金額-去年本月退貨金額) AS DECIMAL(18,2)) AS '去年本月實出金額'");
                STR.AppendFormat(@" ,CAST ((去年本月出貨量-去年本月退貨量) AS DECIMAL(18,2)) AS '去年本月實出貨量'");
                STR.AppendFormat(@" ,CAST (去年本月出貨量 AS DECIMAL(18,2))  AS '去年本月出貨量'");
                STR.AppendFormat(@" ,CAST (去年本月退貨量 AS DECIMAL(18,2)) AS '去年本月退貨量'");
                STR.AppendFormat(@" ,CAST(去年本月出貨金額 AS DECIMAL(18,2)) AS '去年本月出貨金額'");
                STR.AppendFormat(@" ,CAST (去年本月退貨金額 AS DECIMAL(18,2)) AS '去年本月退貨金額'");
                STR.AppendFormat(@" FROM (");
                STR.AppendFormat(@" SELECT ID AS '月份' ,CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AS  '今年' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE  TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TH001='A234'),0) ELSE 0 END AS '本月出貨量' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE TJ021='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND   SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001=''),0)  ELSE 0 END AS '本月退貨量' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK) WHERE  TH020='Y' AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID  AND TH001='A234'),0) ELSE 0 END AS '本月出貨金額' ");
                STR.AppendFormat(@" ,CASE WHEN ID<=MONTH(GETDATE()) THEN ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK) WHERE  TJ021='Y' AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE()) AS NVARCHAR) AND SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE()) AS NVARCHAR)+ID AND TJ001=''),0)  ELSE 0 END AS '本月退貨金額' ");
                STR.AppendFormat(@" ,CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AS '去年' ");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE  TH020='Y' AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TH002,1,6)<=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TH001='A234'),0) AS '去年本月出貨量' ");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(LA011) FROM  [TK].dbo.COPTJ WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK) WHERE  TJ021='Y' AND TJ001=LA006 AND TJ002=LA007 AND TJ003=LA008 AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND   SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001=''),0)   AS '去年本月退貨量' ");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(TH013) FROM  [TK].dbo.COPTH WITH (NOLOCK) WHERE  TH020='Y' AND SUBSTRING(TH002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TH002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID  AND TH001='A234'),0)  AS '去年本月出貨金額' ");
                STR.AppendFormat(@" ,ISNULL((SELECT SUM(TJ012) FROM  [TK].dbo.COPTJ WITH (NOLOCK) WHERE  TJ021='Y' AND SUBSTRING(TJ002,1,4)=CAST(YEAR(GETDATE())-1 AS NVARCHAR) AND SUBSTRING(TJ002,1,6)=CAST(YEAR(GETDATE())-1 AS NVARCHAR)+ID AND TJ001=''),0)  AS '去年本月退貨金額' ");
                STR.AppendFormat(@"  FROM [TKECOMMERCE].dbo.BASEMONTH ");
                STR.AppendFormat(@"  ) AS TEMP ");
                STR.AppendFormat(@" ");
                STR.AppendFormat(@" ");
                STR.AppendFormat(@" ");

                talbename = "TEMPds14";
            }

            else if (comboBox1.Text.ToString().Equals("各電商銷貨金額及數量"))
            {
                STR.AppendFormat(@" SELECT TG007 AS '平台商',SUM(TH013) AS '銷貨金額',SUM(LA011) AS '銷貨數量'");
                STR.AppendFormat(@" FROM [TK].dbo.COPTH,[TK].dbo.COPTG,[TK].dbo.INVLA");
                STR.AppendFormat(@" WHERE TG001=TH001 AND TG002=TH002");
                STR.AppendFormat(@" AND TH001=LA006 AND TH002=LA007 AND TH003=LA008");
                STR.AppendFormat(@" AND TH001='A234'");
                STR.AppendFormat(@" AND TH020='Y'");
                STR.AppendFormat(@" AND SUBSTRING(TH002,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.AppendFormat(@" GROUP BY TG007");
                STR.AppendFormat(@" ORDER BY TG007");
                STR.AppendFormat(@" ");

                talbename = "TEMPds15";
            }
            else if (comboBox1.Text.ToString().Equals("電商銷貨明細"))
            {
                if (!string.IsNullOrEmpty(textBox1.Text.ToString()))
                {
                    STR.Append(@" SELECT 平台商, 品號,品名,CAST(SUM(銷售量) AS DECIMAL(18,2)) AS 銷售量,CAST(SUM(銷售金額) AS DECIMAL(18,2)) AS 銷售金額");
                    STR.Append(@" FROM (");
                    STR.Append(@" SELECT TG007 AS '平台商',TH004  AS '品號',TH005  AS '品名',LA011 AS '銷售量',TH013 AS '銷售金額' ");
                    STR.Append(@" FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)");
                    STR.Append(@" WHERE TH020='Y' AND  TH001=LA006 AND TH002=LA007 AND TH003=LA008");
                    STR.AppendFormat(@" AND TG001=TH001 AND TG002=TH002");
                    STR.AppendFormat(@" AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}'", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                    STR.Append(@" AND TH001='A234'  AND TG005 IN ('102300','114000')  ");
                    STR.AppendFormat(@" AND (TH004 LIKE '%{0}%' OR TH005 LIKE '%{0}%')", textBox1.Text.ToString());
                    STR.Append(@"  ) AS TEMP");
                    STR.Append(@" GROUP BY 平台商,品號,品名");
                    STR.Append(@" ORDER BY 平台商,SUM(銷售金額) DESC");
                    STR.AppendFormat(@" ");


                }
                else
                {
                    STR.Append(@" SELECT  平台商,品號,品名,CAST(SUM(銷售量) AS DECIMAL(18,2)) AS 銷售量,CAST(SUM(銷售金額) AS DECIMAL(18,2)) AS 銷售金額");
                    STR.Append(@" FROM (");
                    STR.Append(@" SELECT TG007 AS '平台商',TH004  AS '品號',TH005  AS '品名',LA011 AS '銷售量',TH013 AS '銷售金額' ");
                    STR.Append(@" FROM [TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.COPTG WITH (NOLOCK)");
                    STR.Append(@" WHERE TH020='Y' AND  TH001=LA006 AND TH002=LA007 AND TH003=LA008");
                    STR.AppendFormat(@" AND TG001=TH001 AND TG002=TH002");
                    STR.AppendFormat(@" AND SUBSTRING(TH002,1,8)>='{0}' AND SUBSTRING(TH002,1,8)<='{1}'", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                    STR.Append(@" AND TH001='A234'  AND TG005 IN ('102300','114000')  ");
                    STR.Append(@"  ) AS TEMP");
                    STR.Append(@" GROUP BY 平台商,品號,品名");
                    STR.Append(@" ORDER BY 平台商,SUM(銷售金額) DESC");
                    STR.AppendFormat(@" ");
                }


                talbename = "TEMPds16";
            }

            else if (comboBox1.Text.ToString().Equals("品號彙總"))
            {

                STR.AppendFormat(@" SELECT TG001,TH004 AS '品號',TH005 AS '品名',CONVERT(real, SUM(TH008)) AS '數量'");
                STR.AppendFormat(@" ,CONVERT(real, SUM(TH024)) AS '贈品',TH009 AS '單位',CONVERT(real, SUM(TH013)) AS '金額' ");
                STR.AppendFormat(@" FROM [TK].dbo.COPTG,[TK].dbo.COPTH WHERE TG001=TH001 AND TG002=TH002  ");
                STR.AppendFormat(@" AND     (TG001='A233'  OR (TG001='A230'  AND TG006  IN ('160092','170007') ) OR TG001='A234') ");
                STR.AppendFormat(@" AND   TG003>='{0}' AND TG003<='{1}' ", dateTimePicker2.Value.ToString("yyyyMMdd"), dateTimePicker3.Value.ToString("yyyyMMdd"));
                STR.AppendFormat(@" GROUP BY TG001,TH004,TH005,TH009");
                STR.AppendFormat(@" ORDER BY TG001,SUM(TH008) DESC");
                STR.AppendFormat(@" ");
                STR.AppendFormat(@" ");
                STR.AppendFormat(@" ");

                talbename = "TEMPds17";
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

        public void ExcelExport()
        {
            Search();
            string TABLENAME = "報表";

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
            if (!string.IsNullOrEmpty(talbename))
            {
                TABLENAME = talbename+"報表";

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    ws.CreateRow(i + 1);
                    for (int rows = 0; rows < dt.Columns.Count; rows++)
                    {
                        ws.GetRow(i + 1).CreateCell(rows).SetCellValue(ds.Tables[talbename].Rows[i][rows].ToString());
                    }
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
            filename.AppendFormat(@"c:\temp\{0}-{1}.xlsx", TABLENAME, DateTime.Now.ToString("yyyyMMdd"));

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

            //Thread TD;

            //TD = new Thread(showwaitfrm);
            //TD.Start();
            //Thread.Sleep(2000);   //此行可以不需要，主要用於等待主窗體填充數據
            
            //TD.Abort(); //主窗體加載完成數據後，線程結束，關閉等待窗體。
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }

        #endregion


    }
}
