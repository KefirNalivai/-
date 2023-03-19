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
    public partial class Абитуриенты : MetroFramework.Forms.MetroForm
    {
        string connectString;
        SqlConnection con;
        DataSet ds;
        SqlDataAdapter da;
        main mn;
        int prov = 1;
        public List<bool> sso = new List<bool>();
        
        public Абитуриенты(main mn1)
        {
            InitializeComponent();

            this.StyleManager = metroStyleManager1;
            metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            mn = mn1;
            connectString = @"Data Source=" + Properties.Settings.Default.serverCon + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + Properties.Settings.Default.User + ";Password=" + mn.pas;
            con = new SqlConnection(connectString);
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM groups", con);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            if (ds.Tables[0].Rows.Count == 0)
            {
                con.Close();

                
               
                return;
            }

            for (int i1 = 0; i1 < ds.Tables[0].Rows.Count; i1++)
            {
                metroComboBox1.Items.Add(ds.Tables[0].Rows[i1][0].ToString() + ":  " + ds.Tables[0].Rows[i1][1].ToString());
                sso.Add(bool.Parse(ds.Tables[0].Rows[i1][3].ToString()));
               
            }
            con.Close();
            if (metroComboBox1.Items.Count != 0)
            {
                metroComboBox1.SelectedIndex = 0;
            }
            dataGridView1.RowHeadersBorderStyle =
                    DataGridViewHeaderBorderStyle.Single;
            dataGridView1.ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Single;
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Times New Roman", 12);
            ((DataGridViewTextBoxColumn)dataGridView1.Columns["ФИО"]).MaxInputLength = 50;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns["Льготы"]).MaxInputLength = 100;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns["Приемущество"]).MaxInputLength = 100;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns["Язык"]).MaxInputLength = 15;
            ((DataGridViewTextBoxColumn)dataGridView1.Columns["Примечание"]).MaxInputLength = 100;
            //  metroComboBox1.Text.Substring(0, metroComboBox1.Text.IndexOf(':'));
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

                checkedListBox1.BackColor = Color.FromArgb(31, 29, 29);
                checkedListBox1.ForeColor = Color.FromArgb(145, 110, 110);
                checkedListBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;

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
                checkedListBox1.BackColor = Color.White;
                checkedListBox1.ForeColor = Color.Black;
                checkedListBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            }
            Properties.Settings.Default.Save();

        }

      

        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT [id],[ФИО], [Целевое обучение], [ср. балл], [Лицо из числа ОПФР],[Льготы],[Приемущество],[Нуждающиеся в общаге],[Пол],[Язык],[Примечание],[Заявление],[Документ об образовании],[Медсправка],[Фото],[Паспорт],[Договор] FROM [Komissiya].[dbo].[abit] WHERE[Номер группы] IN (SELECT [Номер группы] FROM groups WHERE[Номер группы] = @m1) ORDER BY [Целевое обучение] desc, [Льготы] desc, [Приемущество] desc, [Лицо из числа ОПФР] desc ,[ср. балл] desc", con);
            cmd.Parameters.Add("@m1", metroComboBox1.Text.Substring(0, metroComboBox1.Text.IndexOf(':')));
          
            da = new SqlDataAdapter("test", con);
            da.SelectCommand = cmd;
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            
            ds = new DataSet();
            
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[dataGridView1.Columns[0].HeaderText].Visible = false;
            con.Close();
        }

        private void Абитуриенты_Load(object sender, EventArgs e)
        {
            if (metroComboBox1.Items.Count != 0)
            {

                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    checkedListBox1.Items.Add(dataGridView1.Columns[i].HeaderText, false);
                checkedListBox1.Items.Remove("id");
            }
          
        }

        private void Абитуриенты_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (prov == 1)
            {
                mn.metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
                mn.Show();
                mn.WindowState = this.WindowState;
                mn.Location = this.Location;
            }
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;
            MessageBox.Show("Неверный формат ввода данных", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            e.Cancel = false;
        }

        

        

        private void metroButton7_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                con.Open();
                SqlCommand abc = new SqlCommand("DELETE FROM abit WHERE [id] = @dt1", con);
                abc.Parameters.Add("dt1", dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
                abc.ExecuteNonQuery();
                abc = new SqlCommand("DELETE FROM Marks WHERE [id абитуриента] =@dt1", con);
                abc.Parameters.Add("dt1", dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
                abc.ExecuteNonQuery();
                abc = new SqlCommand("DELETE FROM Parents WHERE [id абитуриента] = @dt1", con);
                abc.Parameters.Add("dt1", dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[0].Value);
                abc.ExecuteNonQuery();
                da.Update((DataTable)ds.Tables[0]);
                ds = new DataSet();
                SqlCommand cmd = new SqlCommand("SELECT [id],[ФИО],[Целевое обучение], [ср. балл], [Лицо из числа ОПФР],[Льготы],[Приемущество],[Нуждающиеся в общаге],[Пол],[Язык],[Примечание],[Заявление],[Документ об образовании],[Медсправка],[Фото],[Паспорт],[Договор] FROM [Komissiya].[dbo].[abit] WHERE[Номер группы] IN (SELECT [Номер группы] FROM groups WHERE[Номер группы] = @m1)", con);
                cmd.Parameters.Add("@m1", metroComboBox1.Text.Substring(0, metroComboBox1.Text.IndexOf(':')));
                da = new SqlDataAdapter("test", con);
                da.SelectCommand = cmd;
                da.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                con.Close();
            }
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                Marks mark = new Marks(mn, (int)dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value, this);
                mark.Show();
            }
            else
            {
                MessageBox.Show("Абитуриент не выбран");
            }
        }

        private void metroButton1_Click_1(object sender, EventArgs e)
        {
            con.Open();
            string quary = "SELECT [id],[ФИО], [Целевое обучение], [ср. балл], [Лицо из числа ОПФР],[Льготы],[Приемущество],[Нуждающиеся в общаге],[Пол],[Язык],[Примечание],[Заявление],[Документ об образовании],[Медсправка],[Фото],[Паспорт],[Договор] FROM [Komissiya].[dbo].[abit] WHERE[Номер группы] IN (SELECT [Номер группы] FROM groups WHERE[Номер группы] = @m1)";
            if (checkedListBox1.CheckedItems.Count != 0)
            {
                quary += " AND (";
            }
            for (int i = 0; i < checkedListBox1.CheckedItems.Count; i++)
            {
                if (i == 0)
                {
                    quary += "[" + checkedListBox1.CheckedItems[i].ToString() + "] LIKE @mt1";

                }
                else
                {
                    quary += "OR [" + checkedListBox1.CheckedItems[i].ToString() + "] LIKE @mt1";
                }
            }
            if (checkedListBox1.CheckedItems.Count != 0)
            {
                quary += ")";
            }
            quary += " ORDER BY [Целевое обучение] desc, [Льготы] desc, [Приемущество] desc, [Лицо из числа ОПФР] desc ,[ср. балл] desc";
            SqlCommand cmd = new SqlCommand(quary, con);
            cmd.Parameters.Add("@m1", metroComboBox1.Text.Substring(0, metroComboBox1.Text.IndexOf(':')));
            cmd.Parameters.Add("@mt1", '%' + metroTextBox1.Text + '%');
            da = new SqlDataAdapter("test", con);
            da.SelectCommand = cmd;
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            
            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[dataGridView1.Columns[0].HeaderText].Visible = false;
            con.Close();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            if (metroButton2.Text == "Отметить все")
            {
                for (int i = 0; i < dataGridView1.Columns.Count - 1; i++)
                {
                    checkedListBox1.SetItemChecked(i, true);
                }
                metroButton2.Text = "Отменить выделение";
            }
            else
            {
                for (int i = 0; i < dataGridView1.Columns.Count - 1; i++)
                {
                    checkedListBox1.SetItemChecked(i, false);
                }
                metroButton2.Text = "Отметить все";
            }
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
            metroButton2.Text = "Отметить все";
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT [id],[ФИО], [Целевое обучение], [ср. балл], [Лицо из числа ОПФР],[Льготы],[Приемущество],[Нуждающиеся в общаге],[Пол],[Язык],[Примечание],[Заявление],[Документ об образовании],[Медсправка],[Фото],[Паспорт],[Договор] FROM [Komissiya].[dbo].[abit] WHERE[Номер группы] IN (SELECT [Номер группы] FROM groups WHERE[Номер группы] = @m1) ORDER BY [Целевое обучение] desc, [Льготы] desc, [Приемущество] desc, [Лицо из числа ОПФР] desc ,[ср. балл] desc", con);
            cmd.Parameters.Add("@m1", metroComboBox1.Text.Substring(0, metroComboBox1.Text.IndexOf(':')));
            da = new SqlDataAdapter("test", con);
            da.SelectCommand = cmd;
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[dataGridView1.Columns[0].HeaderText].Visible = false;
            con.Close();
            metroTextBox1.Text = "";
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            metroButton6.PerformClick();
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }
        private void metroButton6_Click(object sender, EventArgs e)
        {
            upd();
        }

        void upd()
        {
            con.Open();

            da.Update(ds);

            con.Close();
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                Parents abc = new Parents(mn, (int)dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value, this);
                abc.Show();
            }
            else
            {
                MessageBox.Show("Абитуриент не выбран");
            }
        }

        private void metroButton8_Click(object sender, EventArgs e)
        {
            AddAbit frm = new AddAbit(metroComboBox1.Text.Substring(0, metroComboBox1.Text.IndexOf(':')), mn, this);
            this.Enabled = false;
            frm.Show();
            
        }

        private void metroButton9_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                Passport abc = new Passport(mn, (int)dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value, this);
                abc.Show();
            }
            else
            {
                MessageBox.Show("Абитуриент не выбран");
            }
        }

        private void Абитуриенты_Shown(object sender, EventArgs e)
        {
            if (metroComboBox1.Items.Count == 0)
            {
                MessageBox.Show("Вы не добавили ни одной группы для внесения абитуриентов. Добавьте их на вкладке \"Группы\"");

                this.Close();
            }
        }
    }
}
