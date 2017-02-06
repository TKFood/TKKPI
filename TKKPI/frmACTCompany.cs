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
    public partial class frmACTCompany : Form
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
        DataSet ds2 = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        string Dep;
        DateTime YEARSMONTHS;

        public frmACTCompany()
        {
            InitializeComponent();
            DateTime dt = DateTime.Now.AddMonths(-1);
            dateTimePicker1.Value = dt;
            comboboxload();
            SearchACTYEARSMONTHSEMP();
        }

        #region FUNCTION
        public void comboboxload()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT ME001,ME002 FROM [TK].dbo.CMSME WITH (NOLOCK) UNION ALL  SELECT '000000','全部' ORDER BY ME001";
            adapter = new SqlDataAdapter(Sequel, sqlConn);
            dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ME001", typeof(string));
            dt.Columns.Add("ME002", typeof(string));
            adapter.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "ME001";
            comboBox2.DisplayMember = "ME002";
            sqlConn.Close();


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
                    adapter.Fill(ds, tablename);
                    sqlConn.Close();
                 

                    label1.Text = "資料筆數:" + ds.Tables[tablename].Rows.Count.ToString();

                    if (ds.Tables[tablename].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        dataGridView1.DataSource = ds.Tables[tablename];
                       
                        //string s_sum = ds.Tables[tablename].Compute("SUM(當月預算)", "").ToString();
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
            string ThisYear = null;
            string ThisMonth = null;
            string LastMonth = null;
            string LastYear = null;
            string LastYearMonth = null;

            ThisYear= dateTimePicker1.Value.ToString("yyyy");
            ThisMonth = dateTimePicker1.Value.ToString("MM");
            LastMonth=dt.AddMonths(-1).ToString("MM");
            LastYear = dt.AddYears(-1).ToString("yyyy");
            LastYearMonth = dt.AddYears(-1).AddMonths(1).ToString("MM");

            if (comboBox1.Text.ToString().Equals("公司整體指標"))
            {
                STR.Append(@"  SELECT 年,月,CAST (ROUND(總負債/總資產,2) AS DECIMAL(18,2)) AS '負債佔資產比率'  ");
                STR.Append(@"  ,CAST (ROUND(流動資產/流動負債,2) AS DECIMAL(18,2))  AS '流動比率' ");
                STR.Append(@"  ,CAST (ROUND(速動資產/流動負債,2) AS DECIMAL(18,2))  AS '速動比率' ");
                STR.Append(@"  ,CAST (ROUND(本12個月營業收入/((本月期初應收帳款+本月期末應收帳款+本月期初應收票據+本月期末應收票據)/2),2)  AS DECIMAL(18,2)) AS '應收帳款週轉率'  ");
                STR.Append(@"  ,CAST (ROUND(365/ROUND(本12個月營業收入/((本月期初應收帳款+本月期末應收帳款+本月期初應收票據+本月期末應收票據)/2),2),2) AS decimal(18,2)) AS '平均收款日數'  ");
                STR.Append(@"  ,CAST (ROUND(銷貨成本/((期初存貨+期末存貨)/2),2) AS DECIMAL(18,2))  AS '存貨週轉率' ");
                STR.Append(@"  ,CAST (ROUND(銷貨成本/((製成品期初存貨+製成品期末存貨)/2),2) AS DECIMAL(18,2))  AS '平均製成品週轉率'  ");
                STR.Append(@"  ,CAST (ROUND(365/ROUND(本12個月營業收入/((本月期初應收帳款+本月期末應收帳款+本月期初應收票據+本月期末應收票據)/2),2),2) AS decimal(18,2))+CAST (ROUND(365/ROUND(銷貨成本/((期初存貨+期末存貨)/2),2),2) AS decimal(18,2))   AS '平均銷貨日數' ");
                STR.Append(@"  ,CAST (ROUND(銷貨成本/應付帳款餘額,2)  AS DECIMAL(18,2)) AS '應付款項週轉率' ");
                STR.Append(@"  ,CAST (ROUND(365/ROUND(銷貨成本/應付帳款餘額,2),2)   AS DECIMAL(18,2))  AS '應付款項週轉天數' ");
                STR.Append(@"  ,CAST (ROUND(365/ROUND(本12個月營業收入/((本月期初應收帳款+本月期末應收帳款+本月期初應收票據+本月期末應收票據)/2),2),2) AS decimal(18,2))-CAST (ROUND(365/ROUND(銷貨成本/應付帳款餘額,2),2) AS decimal(18,2)) AS '應收.付天數' ");
                STR.AppendFormat(@"  ,CAST (ROUND((營業收入類/固定資產淨額)/{0}*12,2)  AS DECIMAL(18,2)) AS '固定資產週轉率' ",Convert.ToInt16(ThisMonth));
                STR.AppendFormat(@"  ,CAST (ROUND((營業收入類/總資產淨額)/{0}*12,2) AS DECIMAL(18,2))  AS '總資產週轉率'", Convert.ToInt16(ThisMonth));
                STR.Append(@"  ,CAST (ROUND((當年度本期損利+年利息費用)/總資產淨額,2) AS DECIMAL(18,2))  AS '資產報酬率' ");
                STR.Append(@"  ,CAST (ROUND(當年度本期損利/股東權益,2) AS DECIMAL(18,2))  AS '權益報酬率(稅後)' ");
                STR.Append(@"  ,CAST (ROUND(損益表本期淨利/(股本/10),2) AS DECIMAL(18,2))  AS '每股盈餘(稅後)' ");
                STR.AppendFormat(@"  ,CAST (ROUND(當月營業收入類/{0},2)  AS DECIMAL(18,2)) AS '每人營業淨額'  ",numericUpDown1.Value.ToString());
                STR.AppendFormat(@"  ,CAST (ROUND((當月損益表本期淨利4+當月損益表本期淨利5+當月損益表本期淨利6)/{0},2) AS DECIMAL(18,2))  AS '每人營業利益'  ", numericUpDown1.Value.ToString());
                STR.Append(@"  ,CAST (ROUND(總負債/股東權益,2) AS DECIMAL(18,2))  AS '負債/淨值'  ");
                STR.Append(@"  ,CAST (ROUND((當月損益表本期淨利4+當月損益表本期淨利5+當月損益表本期淨利6+當月損益表本期淨利7+當期所得稅費用)/利息費用,2)  AS DECIMAL(18,2)) AS '利息保障倍數(稅前)' ");
                STR.Append(@"  FROM (");
                STR.AppendFormat(@"  SELECT '{0}' AS '年','{1}' AS '月'",ThisYear,ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) From [TK].dbo.ACTMB Where MB001 LIKE '1%' and MB002='{0}' and MB003<='{1}'  AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3'))) AS '總資產'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) From [TK].dbo.ACTMB Where MB001 LIKE '2%' and MB002='{0}' and MB003<='{1}'  AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3'))) AS '總負債' ", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) From [TK].dbo.ACTMB Where MB001 LIKE '1%' and MB002='{0}' and MB003<='{1}'  AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) AND MB001 IN (SELECT ID FROM [TKKPI].dbo.ACTCurrentRatio WHERE ID LIKE '1%')) AS '流動資產' ", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) From [TK].dbo.ACTMB Where MB001 LIKE '2%' and MB002='{0}' and MB003<='{1}'  AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) AND MB001 IN (SELECT ID FROM [TKKPI].dbo.ACTCurrentRatio WHERE ID LIKE '2%')) AS '流動負債'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) From [TK].dbo.ACTMB Where MB001 LIKE '1%' and MB002='{0}' and MB003<='{1}'  AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) AND MB001 IN (SELECT ID FROM [TKKPI].dbo.ACTQuickRatio WHERE ID LIKE '1%')) AS '速動資產'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'4%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002+MB003>='{0}'and MB002+MB003<='{1}') AS '本12個月營業收入'", LastYear+ LastYearMonth, ThisYear+ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'115%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' AND MB003<='{1}') AS '本月期初應收帳款'", LastYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'115%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' AND MB003<='{1}') AS '本月期末應收帳款'",ThisYear,ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'117%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' AND MB003<='{1}') AS '本月期初應收票據'", LastYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'117%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' AND MB003<='{1}') AS '本月期末應收票據'", ThisYear, ThisMonth); 
                STR.AppendFormat(@"  ,(SELECT SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) WHERE  MB001='51111' AND  SUBSTRING(LTRIM(RTRIM(MB002))+LTRIM(RTRIM(MB003)),1,6)>='{0}' AND  SUBSTRING(LTRIM(RTRIM(MB002))+LTRIM(RTRIM(MB003)),1,6)<='{1}') AS '銷貨成本'", LastYear + LastYearMonth, ThisYear + ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'13%' and MB002='{0}'and MB003<='{1}') AS '期初存貨'",LastYear,LastYearMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'13%' and MB002='{0}'and MB003<='{1}') AS '期末存貨'",ThisYear,ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'1311%' and MB002='{0}'and MB003<='{1}') AS '製成品期初存貨'", LastYear, LastYearMonth); ;
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'1311%' and MB002='{0}'and MB003<='{1}') AS '製成品期末存貨'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'2171%' and MB002='{0}' and MB003<='{1}') AS '應付帳款餘額'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'4%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003<='{1}') AS '營業收入類'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where( MB001 like'16%' OR MB001 like'17%' OR MB001 like'191%')AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}'and MB003<='{1}') AS '固定資產淨額' ", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'1%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003<='{1}') AS '總資產淨額' ", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'3353%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003<='{1}') AS '當年度本期損利'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where( MB001 like '7511%')AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}'and MB003<='{1}') AS '年利息費用'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'3%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003<='{1}') AS '股東權益'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'3353%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003='{1}') '損益表本期淨利'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'31%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003<='{1}') AS '股本'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'4%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003='{1}') AS '當月營業收入類'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'4%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003='{1}') AS '當月損益表本期淨利4'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'5%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003='{1}') AS '當月損益表本期淨利5'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK)  Where MB001 like'6%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003='{1}' ) AS '當月損益表本期淨利6'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB005)-SUM(MB004) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'7%'  AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3'))  and MB002='{0}' and MB003='{1}' ) AS '當月損益表本期淨利7'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'7951%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003='{1}') AS '當期所得稅費用'", ThisYear, ThisMonth);
                STR.AppendFormat(@"  ,(Select SUM(MB004)-SUM(MB005) FROM [TK].dbo.ACTMB WITH (NOLOCK) Where MB001 like'7511%' AND MB001 IN (SELECT MA001 FROM [TK].dbo.ACTMA WHERE (MA008='2' OR MA008='3')) and MB002='{0}' and MB003='{1}') AS '利息費用'", ThisYear, ThisMonth);
                STR.Append(@"  ) AS TEMP");
                STR.Append(@"  ");


                tablename = "TEMPds1";
            }
            else if(comboBox1.Text.ToString().Equals("實際與預算比較費用報表"))
            {
                Dep = comboBox2.SelectedValue.ToString();

                if(Dep.Equals("000000"))
                {
                    STR.AppendFormat(@"  SELECT 年度,月份,科目,科目名稱,部門代號,部門,當月預算,當月實際費用,當月費用達成率,預算累積,實際費用累積,年度累積達成率,年度預算總額,年度預算累積達成率");
                    STR.AppendFormat(@"  FROM (");
                    STR.AppendFormat(@"  SELECT '{0}' AS '年度','{1}' AS '月份',MA001 AS '科目',MA003 AS '科目名稱',MK004 AS '部門代號',CASE WHEN ISNULL(ME002,'')='' THEN '全公司' ELSE ME002 END AS '部門',CAST( MK006 AS DECIMAL(18,2))   AS '當月預算'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST( ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004='{1}'),0) AS DECIMAL(18,2)) AS '當月實際費用'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@" ,CONVERT(DECIMAL(18,2),ISNULL(CAST( ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004='{1}'),0) AS DECIMAL(18,2))/NULLIF(CAST( MK006 AS DECIMAL(18,2)),0)*100,0)) AS '當月費用達成率' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST( ISNULL((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 AND MK.MK005<='{1}'),0) AS DECIMAL(18,2)) AS '預算累積'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST( ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0) AS DECIMAL(18,2)) AS '實際費用累積'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM")); ;
                    STR.AppendFormat(@" ,CONVERT(DECIMAL(18,2),ISNULL((CAST( NULLIF((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0) AS DECIMAL(18,2))/CAST( NULLIF((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 AND MK.MK005<='{1}'),0) AS DECIMAL(18,2))*100 ),0)) AS '年度累積達成率' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@" ,CAST( ISNULL((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 ),0) AS DECIMAL(18,2)) AS '年度預算總額' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CONVERT(DECIMAL(18,2),ISNULL((CAST( NULLIF((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0) AS DECIMAL(18,2))/CAST( NULLIF((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 ),0) AS DECIMAL(18,2))*100 ),0)) AS '年度預算累積達成率'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  FROM [TK].dbo.ACTMA WITH (NOLOCK),[TK].dbo.ACTMK WITH (NOLOCK)");
                    STR.AppendFormat(@"  LEFT JOIN [TK].dbo.CMSME WITH (NOLOCK) ON MK004=ME001");
                    STR.AppendFormat(@"  WHERE MK003=MA001 AND MK002='{0}'  AND MK005='{1}'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM")); ;
                    STR.AppendFormat(@"  AND (MA001 LIKE '5%' OR MA001 LIKE '6%' )");
                    STR.AppendFormat(@"  ) AS TEMP");
                    STR.AppendFormat(@"  UNION ALL ");
                    STR.AppendFormat(@"  SELECT 'TOTAL','','','','','',SUM(當月預算),SUM(當月實際費用),ROUND(SUM(當月實際費用)/SUM(當月預算)*100,2),SUM(預算累積),SUM(實際費用累積),ROUND(SUM(實際費用累積)/SUM(預算累積)*100,2),SUM(年度預算總額),ROUND(SUM(實際費用累積)/SUM(年度預算總額)*100,2)");
                    STR.AppendFormat(@"  FROM ( ");
                    STR.AppendFormat(@"  SELECT '{0}' AS '年度','{1}' AS '月份',MA001 AS '科目',MA003 AS '科目名稱',MK004 AS '部門代號',CASE WHEN ISNULL(ME002,'')='' THEN '全公司' ELSE ME002 END AS '部門',CAST( MK006 AS DECIMAL(18,2))   AS '當月預算'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST( ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004='{1}'),0) AS DECIMAL(18,2)) AS '當月實際費用'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@" ,CONVERT(DECIMAL(18,2),ISNULL(CAST( ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004='{1}'),0) AS DECIMAL(18,2))/NULLIF(CAST( MK006 AS DECIMAL(18,2)),0)*100,0)) AS '當月費用達成率' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST( ISNULL((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 AND MK.MK005<='{1}'),0) AS DECIMAL(18,2)) AS '預算累積'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST( ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0) AS DECIMAL(18,2)) AS '實際費用累積'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM")); ;
                    STR.AppendFormat(@" ,CONVERT(DECIMAL(18,2),ISNULL((CAST( NULLIF((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0) AS DECIMAL(18,2))/CAST( NULLIF((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 AND MK.MK005<='{1}'),0) AS DECIMAL(18,2))*100 ),0)) AS '年度累積達成率' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@" ,CAST( ISNULL((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 ),0) AS DECIMAL(18,2)) AS '年度預算總額' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CONVERT(DECIMAL(18,2),ISNULL((CAST( NULLIF((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0) AS DECIMAL(18,2))/CAST( NULLIF((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 ),0) AS DECIMAL(18,2))*100 ),0)) AS '年度預算累積達成率'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  FROM [TK].dbo.ACTMA WITH (NOLOCK),[TK].dbo.ACTMK WITH (NOLOCK)");
                    STR.AppendFormat(@"  LEFT JOIN [TK].dbo.CMSME WITH (NOLOCK) ON MK004=ME001");
                    STR.AppendFormat(@"  WHERE MK003=MA001 AND MK002='{0}'  AND MK005='{1}'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM")); ;
                    STR.AppendFormat(@"  AND (MA001 LIKE '5%' OR MA001 LIKE '6%' )");
                    STR.AppendFormat(@"  ) AS TEMP2");
                    STR.AppendFormat(@"  ORDER BY 年度,科目,部門");
                    STR.AppendFormat(@"  ");
                }
                else
                {
                    STR.AppendFormat(@"  SELECT 年度,月份,科目,科目名稱,部門代號,部門,當月預算,當月實際費用,當月費用達成率,預算累積,實際費用累積,年度累積達成率,年度預算總額,年度預算累積達成率");
                    STR.AppendFormat(@"  FROM (");
                    STR.AppendFormat(@"  SELECT '{0}' AS '年度','{1}' AS '月份',MA001 AS '科目',MA003 AS '科目名稱',MK004 AS '部門代號',CASE WHEN ISNULL(ME002,'')='' THEN '全公司' ELSE ME002 END AS '部門',CAST( MK006   AS DECIMAL(18,2)) AS '當月預算'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004='{1}'),0)  AS DECIMAL(18,2)) AS '當月實際費用'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@" ,CONVERT(DECIMAL(18,2),ISNULL(CAST( ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004='{1}'),0) AS DECIMAL(18,2))/NULLIF(CAST( MK006 AS DECIMAL(18,2)),0)*100,0)) AS '當月費用達成率' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 AND MK.MK005<='{1}'),0)   AS DECIMAL(18,2)) AS '預算累積'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0)   AS DECIMAL(18,2))AS '實際費用累積'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM")); ;
                    STR.AppendFormat(@" ,CONVERT(DECIMAL(18,2),ISNULL((CAST( NULLIF((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0) AS DECIMAL(18,2))/CAST( NULLIF((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 AND MK.MK005<='{1}'),0) AS DECIMAL(18,2))*100 ),0)) AS '年度累積達成率' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@" ,CAST( ISNULL((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 ),0) AS DECIMAL(18,2)) AS '年度預算總額' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CONVERT(DECIMAL(18,2),ISNULL((CAST( NULLIF((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0) AS DECIMAL(18,2))/CAST( NULLIF((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 ),0) AS DECIMAL(18,2))*100 ),0)) AS '年度預算累積達成率'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  FROM [TK].dbo.ACTMA WITH (NOLOCK),[TK].dbo.ACTMK WITH (NOLOCK)");
                    STR.AppendFormat(@"  LEFT JOIN [TK].dbo.CMSME WITH (NOLOCK) ON MK004=ME001");
                    STR.AppendFormat(@"  WHERE MK003=MA001 AND MK002='{0}'  AND MK005='{1}' AND MK004='{2}' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"),Dep);
                    STR.AppendFormat(@"  AND (MA001 LIKE '5%' OR MA001 LIKE '6%' )");             
                    STR.AppendFormat(@"  ) AS TEMP");
                    STR.AppendFormat(@"  UNION ALL ");
                    STR.AppendFormat(@"  SELECT 'TOTAL','','','','','',SUM(當月預算),SUM(當月實際費用),ROUND(SUM(當月實際費用)/SUM(當月預算)*100,2),SUM(預算累積),SUM(實際費用累積),ROUND(SUM(實際費用累積)/SUM(預算累積)*100,2),SUM(年度預算總額),ROUND(SUM(實際費用累積)/SUM(年度預算總額)*100,2)");
                    STR.AppendFormat(@"  FROM ( ");
                    STR.AppendFormat(@"  SELECT '{0}' AS '年度','{1}' AS '月份',MA001 AS '科目',MA003 AS '科目名稱',MK004 AS '部門代號',CASE WHEN ISNULL(ME002,'')='' THEN '全公司' ELSE ME002 END AS '部門',CAST( MK006   AS DECIMAL(18,2)) AS '當月預算'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004='{1}'),0)  AS DECIMAL(18,2)) AS '當月實際費用'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@" ,CONVERT(DECIMAL(18,2),ISNULL(CAST( ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004='{1}'),0) AS DECIMAL(18,2))/NULLIF(CAST( MK006 AS DECIMAL(18,2)),0)*100,0)) AS '當月費用達成率' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 AND MK.MK005<='{1}'),0)   AS DECIMAL(18,2)) AS '預算累積'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CAST(ISNULL((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0)   AS DECIMAL(18,2))AS '實際費用累積'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM")); ;
                    STR.AppendFormat(@" ,CONVERT(DECIMAL(18,2),ISNULL((CAST( NULLIF((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0) AS DECIMAL(18,2))/CAST( NULLIF((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 AND MK.MK005<='{1}'),0) AS DECIMAL(18,2))*100 ),0)) AS '年度累積達成率' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@" ,CAST( ISNULL((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 ),0) AS DECIMAL(18,2)) AS '年度預算總額' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  ,CONVERT(DECIMAL(18,2),ISNULL((CAST( NULLIF((SELECT SUM(MD005) FROM  [TK].dbo.ACTMD WITH (NOLOCK) WHERE MD001=MA001 AND MD002=MK004 AND MD003='{0}' AND MD004<='{1}'),0) AS DECIMAL(18,2))/CAST( NULLIF((SELECT SUM(MK006) FROM [TK].dbo.ACTMK MK WITH (NOLOCK) WHERE MK.MK003=MA001 AND MK.MK002='{0}' AND MK.MK004=ME001 ),0) AS DECIMAL(18,2))*100 ),0)) AS '年度預算累積達成率'", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"));
                    STR.AppendFormat(@"  FROM [TK].dbo.ACTMA WITH (NOLOCK),[TK].dbo.ACTMK WITH (NOLOCK)");
                    STR.AppendFormat(@"  LEFT JOIN [TK].dbo.CMSME WITH (NOLOCK) ON MK004=ME001");
                    STR.AppendFormat(@"  WHERE MK003=MA001 AND MK002='{0}'  AND MK005='{1}' AND MK004='{2}' ", dateTimePicker1.Value.ToString("yyyy"), dateTimePicker1.Value.ToString("MM"), Dep);
                    STR.AppendFormat(@"  AND (MA001 LIKE '5%' OR MA001 LIKE '6%' )");
                    STR.AppendFormat(@"  ) AS TEMP2");
                    STR.AppendFormat(@"  ORDER BY 年度,科目,部門");
                    STR.AppendFormat(@"  ");
                }
               

                tablename = "TEMPds2";
            }
            else if (comboBox1.Text.ToString().Equals("營業部門-銷貨成本金額及毛利"))
            {
                STR.Append(@"  SELECT YM AS '年月',DEP  AS '部門',ME002 AS '部門名稱',CAST(SUM(NUM) AS DECIMAL(18,2))  AS '數量',CAST(SUM(MM)  AS DECIMAL(18,2)) AS '銷售金額',CAST(SUM(COST)  AS DECIMAL(18,2)) AS '成本'");
                STR.Append(@"  FROM (");
                STR.Append(@"  SELECT SUBSTRING(TG002,1,6) AS 'YM',TG005  AS 'DEP',LA011  AS 'NUM',TH013  AS 'MM',LA013  AS 'COST'");
                STR.Append(@"  FROM [TK].dbo.COPTG WITH (NOLOCK),[TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH(NOLOCK)");
                STR.AppendFormat(@"  WHERE TG001=TH001 AND TG002=TH002 AND TH001=LA006 AND TH002=LA007 AND TH003=LA008 AND SUBSTRING(TG002,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@"  UNION ALL");
                STR.Append(@"  SELECT SUBSTRING(TP001,1,6),TP002,LA011,TP021,LA013  ");
                STR.Append(@"  FROM [TK].dbo.POSTP WITH (NOLOCK) ,[TK].dbo.INVLA WITH (NOLOCK)");
                STR.AppendFormat(@"  WHERE TP001=LA004 AND TP002=LA006 AND TP004=LA001 AND SUBSTRING(TP001,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM")); 
                STR.Append(@"  UNION ALL");
                STR.Append(@"  SELECT  SUBSTRING(TA002,1,6),TA004 ,LA011*-1,0,LA013 *-1");
                STR.Append(@"  FROM [TK].dbo.INVTA WITH (NOLOCK),[TK].dbo.INVTB WITH (NOLOCK),[TK].dbo.INVLA WITH (NOLOCK)");
                STR.AppendFormat(@"  WHERE  TA001=TB001 AND TA002=TB002 AND TB001=LA006 AND TB002=LA007 AND TB003=LA008 AND TA001='A114' AND SUBSTRING(TA002,1,6)='{0}'  AND TA004 IN (SELECT MA001 FROM [TK].dbo.WSCMA)", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@"  ) AS TEMP");
                STR.Append(@"  LEFT JOIN [TK].dbo.CMSME WITH (NOLOCK) ON ME001=DEP");
                STR.Append(@"  GROUP BY YM,DEP,ME002");
                STR.Append(@"  ORDER BY YM,DEP,ME002");
                STR.Append(@"  ");

                tablename = "TEMPds3";
            }
            else if (comboBox1.Text.ToString().Equals("每月所有負毛利產品"))
            {
                STR.Append(@"  SELECT YM AS '年月',ID AS '品號',NAME AS '品名',COM AS '規格',CAST(SUM(NUM)  AS DECIMAL(18,2)) AS '銷售數量',CAST(SUM(MM)   AS DECIMAL(18,2)) AS '銷售金額',CAST(SUM(COST)  AS DECIMAL(18,2)) AS '成本',CAST( SUM(MM)-SUM(COST)   AS DECIMAL(18,2)) AS '毛利'");
                STR.Append(@"  FROM (");
                STR.Append(@"  SELECT SUBSTRING(TG002,1,6) AS 'YM',TH004 AS 'ID',TH005  AS 'NAME',TH006 AS 'COM',LA011  AS 'NUM',TH013  AS 'MM',LA013  AS 'COST'");
                STR.Append(@"  FROM [TK].dbo.COPTG WITH (NOLOCK),[TK].dbo.COPTH WITH (NOLOCK),[TK].dbo.INVLA WITH(NOLOCK)");
                STR.Append(@"  WHERE TG001=TH001 AND TG002=TH002 AND TH001=LA006 AND TH002=LA007 AND TH003=LA008");
                STR.AppendFormat(@"  AND SUBSTRING(TG002,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@"  UNION ALL");
                STR.Append(@"  SELECT SUBSTRING(TP001,1,6),TP004,MB002,MB003,LA011,TP021,LA013 ");
                STR.Append(@"  FROM [TK].dbo.POSTP WITH (NOLOCK) ,[TK].dbo.INVLA WITH (NOLOCK),[TK].dbo.INVMB WITH (NOLOCK)");
                STR.Append(@"  WHERE TP004=MB001 AND TP001=LA004 AND TP002=LA006 AND TP004=LA001 ");
                STR.AppendFormat(@"  AND SUBSTRING(TP001,1,6)='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));
                STR.Append(@"  ) AS TEMP");
                STR.Append(@"  WHERE  ID NOT LIKE'2%'AND ID NOT LIKE'3%' AND NAME NOT LIKE '%試吃%'");
                STR.Append(@" GROUP BY YM,ID,NAME,COM ");
                STR.Append(@"  HAVING  SUM(MM)>0 AND SUM(MM)-SUM(COST)<=0 ");
                STR.Append(@"  ORDER BY YM,ID,NAME,COM");
                STR.Append(@"  ");

                tablename = "TEMPds4";
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
            if(tablename.Equals("TEMPds1"))
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
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString()));
                    ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));
                    ws.GetRow(j + 1).CreateCell(11).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString()));
                    ws.GetRow(j + 1).CreateCell(12).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString()));
                    ws.GetRow(j + 1).CreateCell(13).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString()));
                    ws.GetRow(j + 1).CreateCell(14).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[14].ToString()));
                    ws.GetRow(j + 1).CreateCell(15).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[15].ToString()));
                    ws.GetRow(j + 1).CreateCell(16).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[16].ToString()));
                    ws.GetRow(j + 1).CreateCell(17).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[17].ToString()));
                    ws.GetRow(j + 1).CreateCell(18).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[18].ToString()));
                    ws.GetRow(j + 1).CreateCell(19).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[19].ToString()));
                    ws.GetRow(j + 1).CreateCell(20).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[20].ToString()));
                    ws.GetRow(j + 1).CreateCell(21).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[21].ToString()));
                    
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
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
                    ws.GetRow(j + 1).CreateCell(9).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString()));
                   
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
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));

                    j++;
                }
            }
            else if (tablename.Equals("TEMPds4"))
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
            filename.AppendFormat(@"c:\temp\公司整體指標{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            YEARSMONTHS = dateTimePicker1.Value;
            SearchACTYEARSMONTHSEMP();
        }
        public void SearchACTYEARSMONTHSEMP()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSql.AppendFormat(" SELECT [ID],[YEARSMONTH],[EMP] FROM [TKKPI].[dbo].[ACTYEARSMONTHSEMP] WHERE [YEARSMONTH]='{0}'", dateTimePicker1.Value.ToString("yyyyMM"));

                adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();

                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    numericUpDown1.Value = 1;
                }
                else
                {
                    numericUpDown1.Value = Convert.ToInt32(ds2.Tables["TEMPds2"].Rows[0]["EMP"].ToString());
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
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            frmACTYEARSMONTHSEMP objfrmACTYEARSMONTHSEMP = new frmACTYEARSMONTHSEMP(YEARSMONTHS);
            objfrmACTYEARSMONTHSEMP.ShowDialog();
            SearchACTYEARSMONTHSEMP();
        }

        #endregion

       
    }
}
