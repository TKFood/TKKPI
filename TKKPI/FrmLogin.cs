using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Reflection;

namespace TKKPI
{
    public partial class FrmLogin : Form
    {
        public FrmLogin()
        {
            InitializeComponent();
        }
        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            LOGIN();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }
        #endregion

        #region LOGIN
        public void LOGIN()
        {
            if (txt_UserName.Text == "" || txt_Password.Text == "")
            {
                MessageBox.Show("請輸入帳號、密碼");
                return;
            }
            try
            {
                //Create SqlConnection
                String connectionString;
                SqlConnection conn;
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                conn = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand("Select * from MNU_Login where UserName=@username and Password=@password", conn);
                cmd.Parameters.AddWithValue("@username", txt_UserName.Text);
                cmd.Parameters.AddWithValue("@password", txt_Password.Text);
                conn.Open();
                SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                adapt.Fill(ds);
                conn.Close();
                int count = ds.Tables[0].Rows.Count;
                //If count is equal to 1, than show frmMain form
                if (count == 1)
                {
                    //MessageBox.Show("登入成功!");

                    FrmParent fm = new FrmParent(txt_UserName.Text.ToString());
                    fm.Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("登入失敗!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion

        #region FUNCTION
        private void txt_Password_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LOGIN();
            }
        }

        #endregion
    }
}
