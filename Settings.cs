using MetroFramework;
using MetroFramework.Controls;
using Microsoft.AspNetCore.Mvc;

using Microsoft.SqlServer.Management.Smo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Комиссия
{
    public partial class Settings : MetroFramework.Forms.MetroForm
    {
        main mn;
        SqlConnection con;
        DataSet ds;
        SqlDataAdapter da;
        string connectString;
       
        public Settings(main mn1)
        {
            InitializeComponent();
            this.StyleManager = metroStyleManager1;
            metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            mn = mn1;
            metroTextBox1.Text = Properties.Settings.Default.serverCon;
            metroTextBox2.Text = Properties.Settings.Default.User;

            if (!string.IsNullOrEmpty(mn.pas))
            {
                connectString = @"Data Source=" + Properties.Settings.Default.serverCon + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + Properties.Settings.Default.User + ";Password=" + mn.pas;
                con = new SqlConnection(connectString);
                con.Open();
                da = new SqlDataAdapter("SELECT * FROM Yasiki", con);
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                ds = new DataSet();
                da.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];
                con.Close();


                connectString = @"Data Source=" + Properties.Settings.Default.serverCon + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + Properties.Settings.Default.User + ";Password=" + mn.pas;
                con = new SqlConnection(connectString);
                con.Open();
                da = new SqlDataAdapter("SELECT * FROM lgoti", con);
                cb = new SqlCommandBuilder(da);
                ds = new DataSet();
                da.Fill(ds);
                dataGridView2.DataSource = ds.Tables[0];
                con.Close();
                metroPanel2.Enabled = true;
            }
            dataGridView1.RowHeadersBorderStyle =
                    DataGridViewHeaderBorderStyle.Single;
            dataGridView1.ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Single;
            dataGridView1.EnableHeadersVisualStyles = false;

            dataGridView2.RowHeadersBorderStyle =
                    DataGridViewHeaderBorderStyle.Single;
            dataGridView2.ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Single;
            dataGridView2.EnableHeadersVisualStyles = false;
           

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

                dataGridView2.BackgroundColor = Color.FromArgb(31, 29, 29);
                dataGridView2.ForeColor = Color.White;
                dataGridView2.GridColor = Color.FromArgb(105, 96, 96);
                            
                dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(31, 29, 29);
                            
                dataGridView2.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(145, 110, 110);
                            
                dataGridView2.RowsDefaultCellStyle.BackColor = Color.FromArgb(31, 29, 29);
                dataGridView2.RowsDefaultCellStyle.ForeColor = Color.FromArgb(145, 110, 110);
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

                dataGridView2.BackgroundColor = Color.White;
                dataGridView2.BackgroundColor = Color.White;
                dataGridView2.ForeColor = Color.Black;
                dataGridView2.GridColor = Color.Black;
                            
                dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.White;
                            
                dataGridView2.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
                            
                dataGridView2.RowsDefaultCellStyle.BackColor = Color.White;
                dataGridView2.RowsDefaultCellStyle.ForeColor = Color.Black;
            }
            Properties.Settings.Default.Save();
        }

        private void Settings_FormClosed(object sender, FormClosedEventArgs e)
        {
            mn.metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            mn.Show();
            mn.WindowState = this.WindowState;
            mn.Location = this.Location;
        }



         private async void metroButton1_Click_1(object sender, EventArgs e)
        {
          
            progressBar1.Visible = true;
            metroButton1.Enabled = false;
            Task task =  Task.Run(() => {
                string connectString = @"Data Source=" + metroTextBox1.Text + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + metroTextBox2.Text + ";Password=" + metroTextBox3.Text + "; Connection Timeout=5";

                con = new SqlConnection(connectString);
                Task abc = Task.Run(() =>
                {
                    for (int i = 1; i < 6; i++)
                    {
                        progressBar1.Invoke(new Action(() => { progressBar1.Value = i; }));
                        Thread.Sleep(1000);


                    }
                   
                });
                try
            {
                con.Open();
            }
            catch
            {
                    metroButton1.Invoke(new Action(() => { metroButton1.Enabled = true; }));
                    progressBar1.Invoke(new Action(() => { progressBar1.Visible = false; }));
                    progressBar1.Invoke(new Action(() => { progressBar1.Value = 0; }));
                    MessageBox.Show("Подключение отсутствует!", "Ошибка");
                return;
            }
                progressBar1.Invoke(new Action(() => { progressBar1.Value = 0; }));
                progressBar1.Invoke(new Action(() => { progressBar1.Visible = false; }));
                Properties.Settings.Default.serverCon = metroTextBox1.Text;
                Properties.Settings.Default.User = metroTextBox2.Text;
                mn.metroButton1.Invoke(new Action(() => { mn.metroButton1.Enabled = true; }));
                mn.metroButton2.Invoke(new Action(() => { mn.metroButton2.Enabled = true; }));
                mn.metroButton3.Invoke(new Action(() => { mn.metroButton3.Enabled = true; }));
                mn.metroLabel1.Invoke(new Action(() => { mn.metroLabel1.Visible = false; }));
                mn.pas = metroTextBox3.Text;
                MessageBox.Show("Вход успешно выполнен");
                metroPanel2.Invoke(new Action(() => { metroPanel2.Enabled = true; }));

                metroButton1.Invoke(new Action(()=> { metroButton1.Enabled = true; }));
                string dir = @"C:\РезервБазы\backup_" + DateTime.Now.Day.ToString() + "_" + DateTime.Now.Month.ToString() + "_" + DateTime.Now.Year.ToString() + ".bak";
                if (!File.Exists(@"C:\РезервБазы"))
                {
                    Directory.CreateDirectory(@"C:/РезервБазы");
                    
                }
                if (!File.Exists(dir))
                {
                    SqlCommand back = new SqlCommand($"BACKUP DATABASE Komissiya TO  DISK = '{dir}'", con);
                    back.ExecuteNonQuery();
                }
                this.Invoke(new Action(() => { this.Close(); }));
            });
        }
        private void metroButton2_Click(object sender, EventArgs e)
        {
        }
        private void metroLabel4_Click(object sender, EventArgs e)
        {
        }
        private void metroTextBox4_Click(object sender, EventArgs e)
        {
        }
        private void metroButton2_Click_1(object sender, EventArgs e)
        {
            con.Open();
            SqlCommand abc = new SqlCommand("INSERT INTO Yasiki VALUES('"+metroTextBox4.Text+"')", con);
            abc.ExecuteNonQuery();
            connectString = @"Data Source=" + Properties.Settings.Default.serverCon + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + Properties.Settings.Default.User + ";Password=" + mn.pas;
            con = new SqlConnection(connectString);
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM Yasiki", con);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            con.Close();
        }
        private void metroButton3_Click(object sender, EventArgs e)
        {
            con.Open();
            SqlCommand abc = new SqlCommand("DELETE FROM Yasiki WHERE [Язык] ='" + dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value.ToString() + "'", con);
            abc.ExecuteNonQuery();
            connectString = @"Data Source=" + Properties.Settings.Default.serverCon + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + Properties.Settings.Default.User + ";Password=" + mn.pas;
            con = new SqlConnection(connectString);
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM Yasiki", con);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            con.Close();
        }
        private void metroButton4_Click(object sender, EventArgs e)
        {
            con.Open();
            SqlCommand abc = new SqlCommand("INSERT INTO lgoti VALUES('" + metroTextBox5.Text + "')", con);
            abc.ExecuteNonQuery();
            connectString = @"Data Source=" + Properties.Settings.Default.serverCon + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + Properties.Settings.Default.User + ";Password=" + mn.pas;
            con = new SqlConnection(connectString);
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM lgoti", con);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            dataGridView2.DataSource = ds.Tables[0];
            con.Close();
        }
        private void metroButton5_Click(object sender, EventArgs e)
        {
            con.Open();
            SqlCommand abc = new SqlCommand("DELETE FROM lgoti WHERE [Льгота] = '" + dataGridView2[0, dataGridView2.CurrentCell.RowIndex].Value.ToString() + "'", con);
            abc.ExecuteNonQuery();
            connectString = @"Data Source=" + Properties.Settings.Default.serverCon + "; Initial Catalog=Komissiya; Integrated Security=False; User ID=" + Properties.Settings.Default.User + ";Password=" + mn.pas;
            con = new SqlConnection(connectString);
            con.Open();
            da = new SqlDataAdapter("SELECT * FROM lgoti", con);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            dataGridView2.DataSource = ds.Tables[0];

            con.Close();
        }
    }
}
