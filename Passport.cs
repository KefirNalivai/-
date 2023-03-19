using MetroFramework;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Комиссия
{
    public partial class Passport : MetroFramework.Forms.MetroForm
    {
        main mn;
        int id;
        string connectString;
        SqlConnection con;
        DataSet ds;
        SqlDataAdapter da;
        Абитуриенты abitF;
        string oldValue;
        public Passport(main mn1, int _id, Абитуриенты _abitF)
        {
            InitializeComponent();
            this.StyleManager = metroStyleManager1;
            metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            mn = mn1;
            connectString = @"Data Source=" + Properties.Settings.Default.serverCon + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + Properties.Settings.Default.User + ";Password=" + mn.pas;

            id = _id;
            con = new SqlConnection(connectString);
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM Passport where [id абитуриента] = " + id, con);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[0].Visible = false;
            dataGridView1.Columns[1].Visible = false;
            con.Close();
            abitF = _abitF;
            abitF.Enabled = false;
            dataGridView1.RowHeadersBorderStyle =
                    DataGridViewHeaderBorderStyle.Single;
            dataGridView1.ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Single;
            dataGridView1.EnableHeadersVisualStyles = false;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns["Адрес"]).MaxInputLength = 50;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns["Номер паспорта"]).MaxInputLength = 9;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns["Идентификационный номер"]).MaxInputLength = 14;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns["Кем выдан"]).MaxInputLength = 70;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Times New Roman", 12);

        }

        private void metroToggle1_CheckedChanged(object sender, EventArgs e)
        {
            if (metroToggle1.Checked)
            {
                metroStyleManager1.Theme = MetroThemeStyle.Dark;
                metroStyleManager1.Style = MetroColorStyle.Blue;
                this.Theme = metroStyleManager1.Theme;
                this.Style = metroStyleManager1.Style;
                Properties.Settings.Default.themeSwitch = true;

                dataGridView1.BackgroundColor = Color.FromArgb(31, 29, 29);
                dataGridView1.ForeColor = Color.White;
                dataGridView1.GridColor = Color.FromArgb(105, 96, 96);

                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(31, 29, 29);

                dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(145, 110, 110);

                dataGridView1.RowsDefaultCellStyle.BackColor = Color.FromArgb(31, 29, 29);
                dataGridView1.RowsDefaultCellStyle.ForeColor = Color.FromArgb(145, 110, 110);
            }
            else
            {
                metroStyleManager1.Theme = MetroThemeStyle.Light;
                metroStyleManager1.Style = MetroColorStyle.Teal;
                this.Theme = metroStyleManager1.Theme;
                this.Style = metroStyleManager1.Style;
                Properties.Settings.Default.themeSwitch = false;

                dataGridView1.BackgroundColor = Color.White;
                dataGridView1.BackgroundColor = Color.White;
                dataGridView1.ForeColor = Color.Black;
                dataGridView1.GridColor = Color.Black;

                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.White;

                dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;

                dataGridView1.RowsDefaultCellStyle.BackColor = Color.White;
                dataGridView1.RowsDefaultCellStyle.ForeColor = Color.Black;
            }
            Properties.Settings.Default.Save();
           
        }

        

        private void metroButton1_Click_1(object sender, EventArgs e)
        {
            upd();
        }

        void upd()
        {
            con.Open();

            da.Update(ds);

            con.Close();
        }

        private void Passport_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                 metroButton1.PerformClick();
                
            }
            abitF.Enabled = true;
            abitF.Select();
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            oldValue = dataGridView1[e.ColumnIndex, 0].Value.ToString();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
            MessageBox.Show("Неверный формат ввода данных", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            e.Cancel = false;
            oldValue = string.Empty;
        }
    }
}
