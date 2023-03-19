using MetroFramework;
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
    public partial class Marks : MetroFramework.Forms.MetroForm
    {
        main mn;
        int id;
        string connectString;
        SqlConnection con;
        DataSet ds;
        SqlDataAdapter da;
        Абитуриенты abitF;
        string oldValue;
        bool sso;
        public Marks(main mn1, int _id, Абитуриенты _abitF)
        {
            InitializeComponent();
            this.StyleManager = metroStyleManager1;
            metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            mn = mn1;
            id = _id;
            connectString = @"Data Source=" + Properties.Settings.Default.serverCon + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + Properties.Settings.Default.User + ";Password=" + mn.pas;

            con = new SqlConnection(connectString);
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM Marks where [id абитуриента] = " + id, con);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[0].Visible = false;
            con.Close();
            srBall();
            abitF = _abitF;
            abitF.Enabled = false;
            sso = abitF.sso[abitF.metroComboBox1.SelectedIndex];
            if (!sso)
            {

                dataGridView1.Columns[18].Visible = false;
                dataGridView1.Columns[19].Visible = false;
                dataGridView1.Columns[20].Visible = false;
            }
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

        private void metroButton1_Click(object sender, EventArgs e)
        {

        }

        void upd()
        {
            con.Open();

            da.Update(ds);

            con.Close();
        }

        private void srBall()
        {
            if (dataGridView1.Rows.Count != 0)
            {
                double sr = 0;
                int count = 0;
                for (int i = 1; i < dataGridView1.Columns.Count; i++)
                {
                    if (!string.IsNullOrEmpty(dataGridView1[i, 0].Value.ToString()) || !string.IsNullOrWhiteSpace(dataGridView1[i, 0].Value.ToString()))
                    {
                        sr += Double.Parse(dataGridView1[i, 0].Value.ToString());
                        count++;
                    }

                }
                if (count != 0)
                {
                    metroTextBox1.Text = Math.Round((sr / count), 3).ToString();
                }
            }
            else
            {
                metroTextBox1.Text = string.Empty;
            }
        }

        private void Marks_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                if (!string.IsNullOrWhiteSpace(metroTextBox1.Text) || !string.IsNullOrEmpty(metroTextBox1.Text))
                {


                    metroButton1.PerformClick();
                    con.Open();
                    SqlCommand abc = new SqlCommand("UPDATE abit SET [ср. балл] = " + metroTextBox1.Text.Replace(',', '.') + " WHERE [id] = " + id, con);
                    abc.ExecuteNonQuery();
                    con.Close();
                    abitF.Enabled = true;
                    abitF.metroButton1.PerformClick();
                    abitF.Select();

                }
                else
                {
                    abitF.Enabled = true;
                    abitF.metroButton1.PerformClick();
                    abitF.Select();

                }

            }
            else
            {
                abitF.Enabled = true;
                abitF.Select();

            }
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
            MessageBox.Show("Неверный формат ввода данных", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            e.Cancel = false;
            oldValue = string.Empty;
        }

        private void metroButton1_Click_1(object sender, EventArgs e)
        {
            upd();
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(dataGridView1[e.ColumnIndex, 0].Value.ToString()))
            {
                if (Double.Parse(dataGridView1[e.ColumnIndex, 0].Value.ToString()) > 10 || Double.Parse(dataGridView1[e.ColumnIndex, 0].Value.ToString()) < 0)
                {
                    MessageBox.Show("Оценка не может быть больше 10 или меньше 0");
                    dataGridView1[e.ColumnIndex, 0].Value = oldValue;
                    oldValue = string.Empty;
                }
                else
                {
                    metroButton1.PerformClick();
                    srBall();
                    oldValue = string.Empty;
                }
            }
            else
            {
                metroButton1.PerformClick();
                srBall();
                oldValue = string.Empty;
            }
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            oldValue = dataGridView1[e.ColumnIndex, 0].Value.ToString();
        }
    }
}
