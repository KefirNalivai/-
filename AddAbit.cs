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
    public partial class AddAbit : MetroFramework.Forms.MetroForm
    {
        string oldval;
        string connectString;
        SqlConnection con;
        DataSet ds;
        SqlDataAdapter da;
        string gr;
        Абитуриенты abitF;
        main mn;
        bool sso;
        public AddAbit(string gr1, main mn1, Абитуриенты _abitF)
        {
            InitializeComponent();
           
            abitF = _abitF;
            this.StyleManager = metroStyleManager1;
            metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            dataGridView1.Rows.Add();
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1[i, 0].Value = string.Empty;
            }
            mn = mn1;
            connectString = @"Data Source=" + Properties.Settings.Default.serverCon + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + Properties.Settings.Default.User + ";Password=" + mn.pas;

            metroComboBox1.SelectedIndex = 1;
           
            con = new SqlConnection(connectString);
            gr = gr1;
            metroComboBox3.SelectedIndex = 0;
           
            metroTextBox5.Text = string.Empty;
            metroTextBox6.Text = string.Empty;
          
            metroTextBox3.Text = string.Empty;
            metroTextBox4.Text = string.Empty;
            metroTextBox10.Text = string.Empty;
            metroTextBox11.Text = string.Empty;
            metroTextBox7.Text = string.Empty;
            metroTextBox8.Text = string.Empty;
            metroTextBox12.Text = string.Empty;
            metroTextBox13.Text = string.Empty;
            metroTextBox14.Text = string.Empty;
            metroTextBox15.Text = string.Empty;
            metroTextBox16.Text = string.Empty;
            metroTextBox17.Text = string.Empty;
            metroTextBox18.Text = string.Empty;
            metroTextBox19.Text = string.Empty;
            metroTextBox20.Text = string.Empty;
            metroTextBox21.Text = string.Empty;

            sso = abitF.sso[abitF.metroComboBox1.SelectedIndex];
            if (!sso)
            {
                
                dataGridView1.Columns[17].Visible = false;
                dataGridView1.Columns[18].Visible = false;
                dataGridView1.Columns[19].Visible = false;
              
            }
            dataGridView1.RowHeadersBorderStyle =
                    DataGridViewHeaderBorderStyle.Single;
            dataGridView1.ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Single;
            dataGridView1.EnableHeadersVisualStyles = false;


           
            con = new SqlConnection(connectString);
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM Yasiki", con);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                metroComboBox2.Items.Add(ds.Tables[0].Rows[i][0].ToString());
            }

            if (metroComboBox2.Items.Count != 0)
            {
                metroComboBox2.SelectedIndex = 0;
            }


            da = new SqlDataAdapter("SELECT * FROM lgoti", con);
           cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                metroComboBox3.Items.Add(ds.Tables[0].Rows[i][0].ToString());
            }
            metroComboBox3.SelectedIndex = 0;

            con.Close();
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
                radMaskedEditBox1.BackColor = Color.FromArgb(31, 29, 29);
                radMaskedEditBox1.ForeColor= Color.FromArgb(145, 110, 110);
                radMaskedEditBox2.BackColor = Color.FromArgb(31, 29, 29);
                radMaskedEditBox2.ForeColor = Color.FromArgb(145, 110, 110);

            }
            else
            {
                metroStyleManager1.Theme = MetroThemeStyle.Light;
                metroStyleManager1.Style = MetroColorStyle.Teal;
                this.Theme = metroStyleManager1.Theme;
                this.Style = metroStyleManager1.Style;
              Properties.Settings.Default.themeSwitch=false;

                dataGridView1.BackgroundColor = Color.White;
                dataGridView1.BackgroundColor = Color.White;
                dataGridView1.ForeColor = Color.Black;
                dataGridView1.GridColor = Color.Black;

                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.White;

                dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;

                dataGridView1.RowsDefaultCellStyle.BackColor = Color.White;
                dataGridView1.RowsDefaultCellStyle.ForeColor = Color.Black;
                radMaskedEditBox1.BackColor = Color.White;
                radMaskedEditBox1.ForeColor = Color.Black;
                radMaskedEditBox2.BackColor = Color.White;
                radMaskedEditBox2.ForeColor = Color.Black;
                
            }
            Properties.Settings.Default.Save();
        }

       

        private void srBall()
        {
            double sr = 0;
            int count = 0;
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                if (!string.IsNullOrEmpty(dataGridView1[i, 0].Value.ToString()))
                {
                    sr += Double.Parse(dataGridView1[i, 0].Value.ToString());
                    count++;
                }

            }
            if (count != 0)
            {
                metroTextBox9.Text = Math.Round((sr / count), 3).ToString();
            }
            else
            {
                metroTextBox9.Text = string.Empty;
            }
        }

        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            oldval = dataGridView1[e.ColumnIndex, 0].Value.ToString();
        }

       

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int index = e.ColumnIndex;
            int rez = 0;
            if (dataGridView1[e.ColumnIndex, 0].Value == null)
            {
                dataGridView1[e.ColumnIndex, 0].Value = string.Empty;
                oldval = string.Empty;
                srBall();
            }
            dataGridView1[e.ColumnIndex, 0].Value = dataGridView1[e.ColumnIndex, 0].Value.ToString().Replace('.', ',');

            if (!string.IsNullOrEmpty(dataGridView1.Rows[e.RowIndex].Cells[index].Value.ToString()))
            {

                if (int.TryParse(dataGridView1.Rows[e.RowIndex].Cells[index].Value.ToString(), out rez))
                {
                    if (rez < 0 || rez > 10)
                    {
                        MessageBox.Show("Вы ввели некорректную оценку");
                        dataGridView1.Rows[e.RowIndex].Cells[index].Value = oldval;
                        oldval = string.Empty;
                        srBall();
                    }
                    else
                    {
                        dataGridView1.Rows[e.RowIndex].Cells[index].Value = rez;
                        srBall();
                    }
                }
                else
                {
                    MessageBox.Show("Введено не число. Повторите ввод");
                    dataGridView1.Rows[e.RowIndex].Cells[index].Value = oldval;
                    oldval = string.Empty;
                    srBall();
                }
            }
            else
            {
                oldval = string.Empty;
                srBall();
            }
        }

        private void AddAbit_FormClosed(object sender, FormClosedEventArgs e)
        {
            abitF.metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            abitF.Enabled = true;
            abitF.Select();
          
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            try
            {
                string momPhone = radMaskedEditBox1.Text.Replace("+", "");
                string dadPhone = radMaskedEditBox2.Text.Replace("+", "");
                momPhone = momPhone.Replace("(", "");
                momPhone = momPhone.Replace(")", "");
                momPhone = momPhone.Replace("-", "");
                dadPhone = dadPhone.Replace("(", "");
                dadPhone = dadPhone.Replace(")", "");
                dadPhone = dadPhone.Replace("-", "");
                momPhone = momPhone.Replace("_", "");
                dadPhone = momPhone.Replace("_", "");


                if (string.IsNullOrEmpty(metroTextBox1.Text))
                {
                    MessageBox.Show("Вы не ввели ФИО абитуриента");
                    return;
                }

                metroTextBox9.Text = metroTextBox9.Text.Replace('.', ',');
                if (!Double.TryParse(metroTextBox9.Text, out double qw))
                {
                    MessageBox.Show("Вы ввели не число в среднем балле!");
                    return;
                }

                if (qw < 0 || qw > 10)
                {
                    MessageBox.Show("Не корректно введен средний балл");
                    return;
                }


                con.Open();

                SqlCommand abc = new SqlCommand("select * from abit where [Номер группы] = @nom and [ФИО] = @fio", con);
                abc.Parameters.Add("@nom", gr);
                abc.Parameters.Add("@fio", metroTextBox1.Text);
                ds = new DataSet();
                da = new SqlDataAdapter("qw", con);
                da.SelectCommand = abc;
                da.Fill(ds);
                if (ds.Tables[0].Rows.Count != 0)
                {
                    con.Close();
                    MessageBox.Show("Абитуриент с таким именем уже находится в группе");
                    return;
                }


                abc = new SqlCommand("INSERT INTO abit VALUES(@gr,@FIO," + (metroCheckBox10.Checked ? "1" : "0") + ", @ball,'" + (metroCheckBox1.Checked ? "1" : "0") + "',@m3,'" + (metroCheckBox2.Checked ? "1" : "0") + "','" + (metroCheckBox3.Checked ? "1" : "0") + "',@m1,@m2,@mt2,'" + (metroCheckBox4.Checked ? "1" : "0") + "','" + (metroCheckBox5.Checked ? "1" : "0") + "','" + (metroCheckBox6.Checked ? "1" : "0") + "','" + (metroCheckBox7.Checked ? "1" : "0") + "','" + (metroCheckBox8.Checked ? "1" : "0") + "','" + (metroCheckBox9.Checked ? "1" : "0") + "')", con);
                abc.Parameters.Add("@FIO", metroTextBox1.Text);
                abc.Parameters.Add("@gr", gr);
                abc.Parameters.Add("@ball", metroTextBox9.Text.Replace(',', '.'));
                abc.Parameters.Add("@m3", metroComboBox3.Text);
                abc.Parameters.Add("@m1", metroComboBox1.Text);
                abc.Parameters.Add("@m2", metroComboBox2.Text);
                abc.Parameters.Add("@mt2", metroTextBox2.Text);
                abc.ExecuteNonQuery();





                string quary = "INSERT INTO Marks([id абитуриента], ";
                string quary2 = " VALUES((SELECT [id] FROM abit where [Номер группы] = @gr AND [ФИО] = @FIO AND [ср. балл] = @ball), ";
                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    if (Double.TryParse(dataGridView1[i, 0].Value.ToString(), out double val))
                    {
                        quary += "[" + dataGridView1.Columns[i].HeaderText + "], ";
                        quary2 += val.ToString().Replace(',', '.') + ", ";
                    }

                }
                quary = quary.Substring(0, quary.Length - 2) + ")";
                quary2 = quary2.Substring(0, quary2.Length - 2) + ")";
                abc = new SqlCommand(quary + quary2, con);
                abc.Parameters.Add("@FIO", metroTextBox1.Text);
                abc.Parameters.Add("@gr", gr);
                abc.Parameters.Add("@ball", metroTextBox9.Text.Replace(',', '.'));
                abc.ExecuteNonQuery();

                abc = new SqlCommand("INSERT INTO Parents VALUES((SELECT [id] FROM abit where [Номер группы] = @gr AND [ФИО] = @FIO AND [ср. балл] = @ball), @mt3, '" + metroDateTime1.Value.Year.ToString() + "-" + metroDateTime1.Value.Month.ToString() + "-" + metroDateTime1.Value.Day.ToString() + "', @mt14, @mt5, " + momPhone + ", @mt10 , @mt16 , @mt19 , @mt20, '" + metroDateTime4.Value.Year.ToString() + "-" + metroDateTime4.Value.Month.ToString() + "-" + metroDateTime4.Value.Day.ToString() + "')", con);
                abc.Parameters.Add("@FIO", metroTextBox1.Text);
                abc.Parameters.Add("@gr", gr);
                abc.Parameters.Add("@ball", metroTextBox9.Text.Replace(',', '.'));
                abc.Parameters.Add("@mt3", metroTextBox3.Text);
                abc.Parameters.Add("@mt4", metroTextBox4.Text);
                abc.Parameters.Add("@mt5", metroTextBox5.Text);
                abc.Parameters.Add("@mt10", metroTextBox10.Text);
                abc.Parameters.Add("@mt14", metroTextBox14.Text);
                abc.Parameters.Add("@mt16", metroTextBox16.Text);
                abc.Parameters.Add("@mt19", metroTextBox19.Text);
                abc.Parameters.Add("@mt20", metroTextBox20.Text);
                abc.ExecuteNonQuery();


                abc = new SqlCommand("INSERT INTO Parents VALUES((SELECT [id] FROM abit where [Номер группы] = @gr AND [ФИО] = @FIO AND [ср. балл] = @ball), @mt4, '" + metroDateTime2.Value.Year.ToString() + "-" + metroDateTime2.Value.Month.ToString() + "-" + metroDateTime2.Value.Day.ToString() + "', @mt15, @mt6, " + dadPhone + ", @mt11, @mt17, @mt18, @mt21, '" + metroDateTime5.Value.Year.ToString() + "-" + metroDateTime5.Value.Month.ToString() + "-" + metroDateTime5.Value.Day.ToString() + "')", con);
                abc.Parameters.Add("@FIO", metroTextBox1.Text);
                abc.Parameters.Add("@gr", gr);
                abc.Parameters.Add("@ball", metroTextBox9.Text.Replace(',', '.'));
                abc.Parameters.Add("@mt4", metroTextBox4.Text);
                abc.Parameters.Add("@mt15", metroTextBox15.Text);
                abc.Parameters.Add("@mt6", metroTextBox6.Text);
                abc.Parameters.Add("@mt11", metroTextBox11.Text);
                abc.Parameters.Add("@mt17", metroTextBox17.Text);
                abc.Parameters.Add("@mt18", metroTextBox18.Text);
                abc.Parameters.Add("@mt21", metroTextBox21.Text);
                abc.ExecuteNonQuery();


                abc = new SqlCommand("INSERT INTO Passport VALUES((SELECT [id] FROM abit where [Номер группы] = @gr AND [ФИО] = @FIO AND [ср. балл] = @ball), @mt13, @mt8, @mt7, @mt12, '" + metroDateTime3.Value.Year.ToString() + "-" + metroDateTime3.Value.Month.ToString() + "-" + metroDateTime3.Value.Day.ToString() + "', '" + metroDateTime6.Value.Year.ToString() + "-" + metroDateTime6.Value.Month.ToString() + "-" + metroDateTime6.Value.Day.ToString() + "')", con);
                abc.Parameters.Add("@FIO", metroTextBox1.Text);
                abc.Parameters.Add("@gr", gr);
                abc.Parameters.Add("@ball", metroTextBox9.Text.Replace(',', '.'));
                abc.Parameters.Add("@mt13", metroTextBox13.Text);
                abc.Parameters.Add("@mt7", metroTextBox7.Text);
                abc.Parameters.Add("@mt12", metroTextBox12.Text);
                abc.Parameters.Add("@mt8", metroTextBox8.Text);

                abc.ExecuteNonQuery();

                con.Close();
                MessageBox.Show("Данные успешно внесены");




                abitF.Enabled = true;
                abitF.metroButton1.PerformClick();
                this.Close();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                con.Close();
            }

        }

        private void metroTextBox8_LocationChanged(object sender, EventArgs e)
        {

        }

        private void metroTextBox2_Click(object sender, EventArgs e)
        {

        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            string abc = radMaskedEditBox1.Text.Replace("+", "");
            abc = abc.Replace("(", "");
            abc = abc.Replace(")", "");
            abc = abc.Replace("-", "");
            MessageBox.Show(abc);
        }
    }
}
