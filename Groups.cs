using MetroFramework;
using MetroFramework.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Комиссия
{
    public partial class Groups : MetroFramework.Forms.MetroForm
    {
        string connectString;
        SqlConnection con;
        DataSet ds;
        SqlDataAdapter da;
        main mn;
        public Groups(main mn1)
        {
            InitializeComponent();

            this.StyleManager = metroStyleManager1;
            metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            mn = mn1;
            connectString = @"Data Source=" + Properties.Settings.Default.serverCon + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + Properties.Settings.Default.User + ";Password=" + mn.pas;
            con = new SqlConnection(connectString);
            con.Open();
            da = new SqlDataAdapter("select * from groups", con);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            con.Close();
            metroTextBox3.Text = string.Empty;
            metroTextBox4.Text = string.Empty;
            dataGridView1.RowHeadersBorderStyle =
                    DataGridViewHeaderBorderStyle.Single;
            dataGridView1.ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Single;
            dataGridView1.EnableHeadersVisualStyles = false;
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
                numericUpDown1.BackColor = Color.FromArgb(31, 29, 29);
                numericUpDown1.ForeColor = Color.White;

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
                numericUpDown1.BackColor = Color.White;
                numericUpDown1.ForeColor = Color.Black;

            }
            Properties.Settings.Default.Save();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            try
            {

                if (string.IsNullOrEmpty(metroTextBox1.Text) || string.IsNullOrWhiteSpace(metroTextBox1.Text))
                {
                    MessageBox.Show("Вы не ввели номер группы");
                    return;
                }

                if (string.IsNullOrEmpty(metroTextBox2.Text) || string.IsNullOrWhiteSpace(metroTextBox1.Text))
                {
                    MessageBox.Show("Вы не ввели название профессии");
                    return;
                }

                con.Open();




                SqlCommand abc = new SqlCommand("select * from groups where [Номер группы] = @nom", con);
                abc.Parameters.Add("@nom", metroTextBox1.Text);
                ds = new DataSet();
                da = new SqlDataAdapter("qw", con);
                da.SelectCommand = abc;
                da.Fill(ds);
                if (ds.Tables[0].Rows.Count != 0)
                {
                    con.Close();
                    MessageBox.Show("Введенный номер группы уже находится в базе");
                    return;
                }
                abc = new SqlCommand("insert into groups values(@mt1, @mt2, " + numericUpDown1.Value + ", '" + (metroCheckBox2.Checked ? "1" : "0") + "', @mt3,@mt4)", con);
                abc.Parameters.Add("@mt1", metroTextBox1.Text);
                abc.Parameters.Add("@mt2", metroTextBox2.Text);
                abc.Parameters.Add("@mt3", metroTextBox3.Text);
                abc.Parameters.Add("@mt4", metroTextBox4.Text);
                abc.ExecuteNonQuery();
                da.Update((DataTable)ds.Tables[0]);
                ds = new DataSet();
                da = new SqlDataAdapter("select * from groups", con);
                da.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                con.Close();
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                con.Open();
                SqlCommand abc = new SqlCommand("DELETE FROM groups WHERE [Номер группы] =@db1", con);
                abc.Parameters.Add("@db1", dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value.ToString());
                abc.ExecuteNonQuery();
                da.Update((DataTable)ds.Tables[0]);
                ds = new DataSet();
                da = new SqlDataAdapter("select * from groups", con);
                da.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                con.Close();
            }
        }

        private void Groups_Load(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                dataGridView1.Rows[dataGridView1.CurrentRow.Index].Selected = false;
                dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0];
            }
        }

        private void Groups_FormClosed(object sender, FormClosedEventArgs e)
        {
            mn.metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            mn.Show();
            mn.WindowState = this.WindowState;
            mn.Location = this.Location;
        }

        private void metroCheckBox2_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}
