using MetroFramework;
using MetroFramework.Controls;
using Microsoft.Office.Interop.Word;
using Microsoft.SqlServer.Management.Smo;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;



namespace Комиссия
{
    public partial class zachislit : MetroFramework.Forms.MetroForm
    {
       
        SqlConnection con;
        DataSet ds;
        SqlDataAdapter da;
        main mn;
        string connectString;
        public List<bool> sso = new List<bool>();
        public List<string> spec = new List<string>();
        public List<string> kvalif = new List<string>();
        public zachislit(main mn1)
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
            for (int i1 = 0; i1 < ds.Tables[0].Rows.Count; i1++)
            {
                metroComboBox1.Items.Add(ds.Tables[0].Rows[i1][0].ToString() + ":  " + ds.Tables[0].Rows[i1][1].ToString());
                sso.Add(bool.Parse(ds.Tables[0].Rows[i1][3].ToString()));
                spec.Add(ds.Tables[0].Rows[i1][4].ToString());
                kvalif.Add(ds.Tables[0].Rows[i1][5].ToString());

            }
            con.Close();
            if (metroComboBox1.Items.Count != 0)
            {
                metroComboBox1.SelectedIndex = 0;
            }
            dataGridView1.RowHeadersBorderStyle =
                    DataGridViewHeaderBorderStyle.None;
            dataGridView1.ColumnHeadersBorderStyle =
                DataGridViewHeaderBorderStyle.Single;
            dataGridView1.EnableHeadersVisualStyles = false;
            metroRadioButton2.Checked = true;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Times New Roman", 12);
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


                dataGridView1.BackgroundColor = Color.FromArgb(31,29,29);
                dataGridView1.ForeColor= Color.White;
                dataGridView1.GridColor= Color.FromArgb(105, 96, 96);                           

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
            if (dataGridView1.Rows.Count != 0)
            {
                try
                {
                    string filenamep = @"Зачисления\Зачисление";
                    string shablon = filenamep;
                    int count = 2;

                    while (File.Exists(Directory.GetCurrentDirectory() + @"\" + filenamep + ".doc"))
                    {

                        filenamep = shablon + "_" + count.ToString();
                        count++;
                    }
                    filenamep = filenamep + ".doc";
                    File.Copy(@"Шаблоны\Отчет.doc", Directory.GetCurrentDirectory() + @"\" + filenamep);
                    if (!sso[metroComboBox1.SelectedIndex])
                    {

                        Word.Application app = new Word.Application();
                        Object fileName = Directory.GetCurrentDirectory() + @"\" + filenamep;
                        Object missing = Type.Missing;
                        app.Documents.Open(ref fileName);

                        var Paragraph = app.ActiveDocument.Paragraphs.Add();
                        //Получаем диапазон
                        var tableRange = Paragraph.Range;
                        //Добавляем таблицу 2х2 в указаный диапазон
                        con.Open();

                        if (dataGridView1.RowCount == 0)
                        {
                            app.ActiveDocument.Save();
                            app.ActiveDocument.Close();
                            con.Close();
                            app.Quit();
                            MessageBox.Show("Данные пустые");
                        }
                        else
                        {
                            app.ActiveDocument.Tables.Add(tableRange, dataGridView1.RowCount + 1, 7);

                            var table = app.ActiveDocument.Tables[app.ActiveDocument.Tables.Count];
                            table.set_Style("Сетка таблицы");
                            table.ApplyStyleHeadingRows = true;
                            table.ApplyStyleLastRow = false;
                            table.ApplyStyleFirstColumn = true;
                            table.ApplyStyleLastColumn = false;
                            table.ApplyStyleRowBands = true;
                            table.ApplyStyleColumnBands = false;
                            table.AllowAutoFit = true;

                            Word.Cell cell = app.ActiveDocument.Tables[1].Cell(1, 1);
                            cell.Range.Text = "ФИО";
                            cell = app.ActiveDocument.Tables[1].Cell(1, 2);
                            cell.Range.Text = "Целевое обучение";
                            cell = app.ActiveDocument.Tables[1].Cell(1, 3);
                            cell.Range.Text = "Средний балл";
                            cell = app.ActiveDocument.Tables[1].Cell(1, 4);
                            cell.Range.Text = "Льготы";
                            cell = app.ActiveDocument.Tables[1].Cell(1, 5);
                            cell.Range.Text = "Приемущество";
                            cell = app.ActiveDocument.Tables[1].Cell(1, 6);
                            cell.Range.Text = "Нуждается в общежитии";
                            cell = app.ActiveDocument.Tables[1].Cell(1, 7);
                            cell.Range.Text = "Иностранный язык";



                            for (int i = 0; i < dataGridView1.Rows.Count; i++)
                            {
                                cell = app.ActiveDocument.Tables[1].Cell(i + 2, 1);
                                cell.Range.Text = dataGridView1[1, i].Value.ToString();
                                cell = app.ActiveDocument.Tables[1].Cell(i + 2, 2);
                                cell.Range.Text = dataGridView1[2, i].Value.ToString() == "true" ? "Да" : "Нет";
                                cell = app.ActiveDocument.Tables[1].Cell(i + 2, 3);
                                cell.Range.Text = dataGridView1[3, i].Value.ToString();
                                cell = app.ActiveDocument.Tables[1].Cell(i + 2, 4);
                                cell.Range.Text = dataGridView1[5, i].Value.ToString();
                                cell = app.ActiveDocument.Tables[1].Cell(i + 2, 5);
                                cell.Range.Text = dataGridView1[6, i].Value.ToString() == "true" ? "Да" : "Нет";
                                cell = app.ActiveDocument.Tables[1].Cell(i + 2, 6);
                                cell.Range.Text = dataGridView1[7, i].Value.ToString() == "true" ? "Да" : "Нет";
                                cell = app.ActiveDocument.Tables[1].Cell(i + 2, 7);
                                cell.Range.Text = dataGridView1[9, i].Value.ToString();
                            }


                            app.ActiveDocument.Save();
                            app.ActiveDocument.Close();
                            con.Close();
                            app.Quit();
                        }
                    }
                    else
                    {
                        Word.Application app = new Word.Application();
                        Object fileName = Directory.GetCurrentDirectory() + @"\" + filenamep;
                        Object missing = Type.Missing;
                        app.Documents.Open(ref fileName);

                        var Paragraph = app.ActiveDocument.Paragraphs.Add();
                        //Получаем диапазон
                        var tableRange = Paragraph.Range;
                        //Добавляем таблицу 2х2 в указаный диапазон
                        con.Open();
                        da = new SqlDataAdapter("SELECT COUNT(*) FROM groups WHERE groups.ССО = 1", con);
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        ds = new DataSet();
                        da.Fill(ds);
                        if (int.Parse(ds.Tables[0].Rows[0][0].ToString()) == 0)
                        {
                            app.ActiveDocument.Save();
                            app.ActiveDocument.Close();
                            con.Close();
                            app.Quit();
                            MessageBox.Show("Данные пустые");
                        }
                        else
                        {
                            app.ActiveDocument.Tables.Add(tableRange, 21, int.Parse(ds.Tables[0].Rows[0][0].ToString()) + 2);

                            var table = app.ActiveDocument.Tables[app.ActiveDocument.Tables.Count];
                            table.set_Style("Сетка таблицы");
                            table.ApplyStyleHeadingRows = true;
                            table.ApplyStyleLastRow = false;
                            table.ApplyStyleFirstColumn = true;
                            table.ApplyStyleLastColumn = false;
                            table.ApplyStyleRowBands = true;
                            table.ApplyStyleColumnBands = false;
                            Word.Cell cell = app.ActiveDocument.Tables[1].Cell(1, 1);
                            cell.Range.Text = "Квалификация";
                            cell = app.ActiveDocument.Tables[1].Cell(1, app.ActiveDocument.Tables[1].Columns.Count);
                            cell.Range.Text = "ИТОГО";
                            cell = app.ActiveDocument.Tables[1].Cell(2, 1);
                            cell.Range.Text = "Группа";
                            cell = app.ActiveDocument.Tables[1].Cell(3, 1);
                            cell.Range.Text = "План приёма";
                            cell = app.ActiveDocument.Tables[1].Cell(4, 1);
                            cell.Range.Text = "Подано заявок";
                            cell = app.ActiveDocument.Tables[1].Cell(5, 1);
                            cell.Range.Text = "Вне конкурса";
                            cell = app.ActiveDocument.Tables[1].Cell(6, 1);
                            cell.Range.Text = "Баллы:";
                            cell = app.ActiveDocument.Tables[1].Cell(7, 1);
                            cell.Range.Text = "до 3";
                            cell = app.ActiveDocument.Tables[1].Cell(8, 1);
                            cell.Range.Text = "3,1 - 3,5";
                            cell = app.ActiveDocument.Tables[1].Cell(9, 1);
                            cell.Range.Text = "3,6 - 4,0";
                            cell = app.ActiveDocument.Tables[1].Cell(10, 1);
                            cell.Range.Text = "4,1 - 4,5";

                            cell = app.ActiveDocument.Tables[1].Cell(11, 1);
                            cell.Range.Text = "4,6 - 5,0";
                            cell = app.ActiveDocument.Tables[1].Cell(12, 1);
                            cell.Range.Text = "5,1 - 5,5";
                            cell = app.ActiveDocument.Tables[1].Cell(13, 1);
                            cell.Range.Text = "5,6 - 6,0";
                            cell = app.ActiveDocument.Tables[1].Cell(14, 1);
                            cell.Range.Text = "6,1 - 6,5";
                            cell = app.ActiveDocument.Tables[1].Cell(15, 1);
                            cell.Range.Text = "6,6 - 7,0";
                            cell = app.ActiveDocument.Tables[1].Cell(16, 1);
                            cell.Range.Text = "7,1 - 7,5";
                            cell = app.ActiveDocument.Tables[1].Cell(17, 1);
                            cell.Range.Text = "7,6 - 8,0";
                            cell = app.ActiveDocument.Tables[1].Cell(18, 1);
                            cell.Range.Text = "8,1 - 8,5";
                            cell = app.ActiveDocument.Tables[1].Cell(19, 1);
                            cell.Range.Text = "8,6 - 9,0";
                            cell = app.ActiveDocument.Tables[1].Cell(20, 1);
                            cell.Range.Text = "9,1 - 9,5";
                            cell = app.ActiveDocument.Tables[1].Cell(21, 1);
                            cell.Range.Text = "9,6 - 10";



                            da = new SqlDataAdapter("SELECT groups.ССО, groups.Квалификация ,groups.[Номер группы], groups.[Численность группы] , Count(abit.[Документ об образовании]) as 'подано заявок', (Select COUNT(*) FROM abit WHERE abit.[Номер группы] = groups.[Номер группы] and [Целевое обучение] = 1) as 'Вне конкурса','__', (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] <= 3 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 3.1 and abit.[ср. балл] <= 3.5 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 3.6 and abit.[ср. балл] <= 4.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 4.1 and abit.[ср. балл] <= 4.5 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 4.6 and abit.[ср. балл] <= 5.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 5.1 and abit.[ср. балл] <= 5.5 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 5.5 and abit.[ср. балл] <= 6.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 6.1 and abit.[ср. балл] <= 6.5 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 6.6 and abit.[ср. балл] <= 7.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 7.1 and abit.[ср. балл] <= 7.5 and abit.[Номер группы] = groups.[Номер группы]),(SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 7.6 and abit.[ср. балл] <= 8.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 8.1 and abit.[ср. балл] <= 8.5 and abit.[Номер группы] = groups.[Номер группы]),(SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 8.6 and abit.[ср. балл] <= 9.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 9.0 and abit.[ср. балл] <= 9.5 and abit.[Номер группы] = groups.[Номер группы]),(SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 9.6 and abit.[Номер группы] = groups.[Номер группы])  FROM groups LEFT JOIN abit on abit.[Номер группы] = groups.[Номер группы] Group by groups.[Номер группы], groups.Квалификация, groups.[Численность группы], groups.ССО HAVING groups.ССО = 1", con);
                            cb = new SqlCommandBuilder(da);
                            ds = new DataSet();
                            da.Fill(ds);

                            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                            {
                                for (int j = 1; j < ds.Tables[0].Columns.Count; j++)
                                {
                                    if (j == 6)
                                    { continue; }
                                    cell = app.ActiveDocument.Tables[1].Cell(j, i + 2);
                                    cell.Range.Text = ds.Tables[0].Rows[i][j].ToString();
                                }
                            }

                            int[] markscount = new int[15] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                            for (int i1 = 0; i1 < ds.Tables[0].Rows.Count; i1++)
                            {
                                for (int j1 = 7; j1 < ds.Tables[0].Columns.Count; j1++)
                                {
                                    if (ds.Tables[0].Rows[i1][j1].ToString() == "1")
                                    {
                                        markscount[j1 - 7]++;
                                    }
                                }
                            }

                            for (int i2 = 7; i2 <= app.ActiveDocument.Tables[1].Rows.Count; i2++)
                            {
                                cell = app.ActiveDocument.Tables[1].Cell(i2, app.ActiveDocument.Tables[1].Columns.Count);
                                cell.Range.Text = markscount[i2 - 7].ToString();
                            }

                            da = new SqlDataAdapter("SELECT  (SELECT SUM([Численность группы]) FROM groups where groups.ССО = 1), (SELECT COUNT(*) FROM abit LEFT JOIN groups on abit.[Номер группы] = groups.[Номер группы] group by groups.ССО HAVING groups.ССО = 1)", con);
                            cb = new SqlCommandBuilder(da);
                            ds = new DataSet();
                            da.Fill(ds);

                            cell = app.ActiveDocument.Tables[1].Cell(3, app.ActiveDocument.Tables[1].Columns.Count);
                            cell.Range.Text = ds.Tables[0].Rows[0][0].ToString();
                            cell = app.ActiveDocument.Tables[1].Cell(4, app.ActiveDocument.Tables[1].Columns.Count);
                            cell.Range.Text = ds.Tables[0].Rows[0][1].ToString();



                            app.ActiveDocument.Save();
                            app.ActiveDocument.Close();
                            con.Close();
                            app.Quit();
                        }

                    }
                    MessageBox.Show("Докумемнт успешно создан");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());


                }
            }
            else
            {
                MessageBox.Show("Абитуриент не выбран");
            }

        }

        private void metroComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT [id],[ФИО],[Целевое обучение], [ср. балл], [Лицо из числа ОПФР],[Льготы],[Приемущество],[Нуждающиеся в общаге],[Пол],[Язык],[Примечание],[Заявление],[Документ об образовании],[Медсправка],[Фото],[Паспорт],[Договор] FROM [Komissiya].[dbo].[abit] WHERE[Номер группы] IN (SELECT [Номер группы] FROM groups WHERE[Номер группы] = @m1) ORDER BY [Целевое обучение] desc, [Льготы] desc, [Приемущество] desc, [Лицо из числа ОПФР] desc ,[ср. балл] desc", con);
            cmd.Parameters.Add("@m1", metroComboBox1.Text.Substring(0, metroComboBox1.Text.IndexOf(':')));
            da = new SqlDataAdapter("test", con);
            da.SelectCommand = cmd;
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            ds = new DataSet();
            da.Fill(ds);
            dataGridView1.DataSource = ds.Tables[0];
            dataGridView1.Columns[dataGridView1.Columns[0].HeaderText].Visible = false;
            con.Close();

            if (!sso[metroComboBox1.SelectedIndex])
            {

                metroComboBox2.Items.Clear();

                metroComboBox2.Items.Add("ПТО после 9");
                metroComboBox2.Items.Add("ПТО после 11");
                metroComboBox2.Items.Add("Заявление ПТО");

                metroComboBox2.SelectedIndex = 0;

                metroPanel1.Visible = true;
            }
            else
            {
                metroComboBox2.Items.Clear();
                metroComboBox2.Items.Add("ССО");
               
                metroComboBox2.Items.Add("Заявление ССО");
                metroComboBox2.SelectedIndex= 0;
                metroPanel1.Visible = false;

            }

        }

        private void zachislit_FormClosed(object sender, FormClosedEventArgs e)
        {
            mn.metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            mn.Show();
            mn.WindowState = this.WindowState;
            mn.Location = this.Location;
        }

        

        private void metroButton3_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                Word.Application app = new Word.Application();
                try
                {
                   
                    switch (metroComboBox2.Text)
                    {

                        case "ПТО после 9":
                            {
                                string filenamep = @"Зачисления\ПТО после 9_" + dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value.ToString();
                                string shablon = filenamep;
                                int count = 2;

                                while (File.Exists(Directory.GetCurrentDirectory() + @"\" + filenamep + ".doc"))
                                {

                                    filenamep = shablon + "_" + count.ToString();
                                    count++;
                                }
                                filenamep = filenamep + ".doc";
                                File.Copy(Directory.GetCurrentDirectory() + @"\Шаблоны\ПТО после 9.doc", Directory.GetCurrentDirectory() + @"\" + filenamep);
                                Object fileName = Directory.GetCurrentDirectory() + @"\" + filenamep;
                                Object missing = Type.Missing;
                                app.Documents.Open(ref fileName);
                                Word.Find find = app.Selection.Find;
                                Object wrap = Word.WdFindWrap.wdFindContinue;
                                Object replace = Word.WdReplace.wdReplaceAll;

                                con.Open();
                                SqlCommand cmd = new SqlCommand("SELECT * FROM Passport WHERE [id абитуриента] = @dt1", con);
                                cmd.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                                da = new SqlDataAdapter("test", con);
                                da.SelectCommand= cmd;
                                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                                ds = new DataSet();
                                da.Fill(ds);
                                SqlCommand cmd2 = new SqlCommand("SELECT * FROM Parents WHERE [id абитуриента] = @dt1", con);
                                cmd2.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                                SqlDataAdapter da2 = new SqlDataAdapter("test", con);
                                da2.SelectCommand= cmd2;
                                SqlCommandBuilder cb2 = new SqlCommandBuilder(da);
                                DataSet ds2 = new DataSet();
                                int opek;
                                da2.Fill(ds2);
                                if (metroRadioButton1.Checked)
                                {
                                    opek = 0;
                                }
                                else
                                {
                                    opek = 1;
                                }

                                foreach (Shape sh in app.ActiveDocument.Shapes)
                                {
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.fio}", dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value.ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.adress}", ds.Tables[0].Rows[0][2].ToString());

                                    if (ds.Tables[0].Rows[0][3].ToString().Length != 0)
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn1}", ds.Tables[0].Rows[0][3].ToString().Substring(0, 2));
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn2}", ds.Tables[0].Rows[0][3].ToString().Substring(2, ds.Tables[0].Rows[0][3].ToString().Length - 2));                                      
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn1}", "  ");
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn2}", "    ");
                                    }

                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasvid}", ds.Tables[0].Rows[0][5].ToString());
                                    DateTime date;
                                    if (DateTime.TryParse(ds.Tables[0].Rows[0][6].ToString(), out date))
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasdate}", date.ToShortDateString());
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasdate}", "");
                                    }
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasln}", ds.Tables[0].Rows[0][4].ToString());

                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pfio}", ds2.Tables[0].Rows[opek][2].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.padr}", ds2.Tables[0].Rows[opek][4].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.ppasn}", ds2.Tables[0].Rows[opek][9].ToString());
                         
                                    if (DateTime.TryParse(ds2.Tables[0].Rows[opek][11].ToString(), out date))
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.ppasdate}", date.ToShortDateString());
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.ppasdate}", " ");
                                    }
                                    
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.ppasvidan}", ds2.Tables[0].Rows[opek][10].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.ppasln}", ds2.Tables[0].Rows[opek][8].ToString());




                                }

                                find.Text = "{info.specialn}";
                                find.Replacement.Text = spec[metroComboBox1.SelectedIndex];
                                find.Execute(FindText: Type.Missing, Replace: replace);

                                find.Text = "{info.cvalif}";
                                find.Replacement.Text = kvalif[metroComboBox1.SelectedIndex];
                                find.Execute(FindText: Type.Missing, Replace: replace);
                               
                               

                               


                              

                               
                               

                                

                             

                                con.Close();
                                app.ActiveDocument.Save();
                                app.ActiveDocument.Close();

                                app.Quit();

                                break;
                            }

                        case "ПТО после 11":
                            {
                                string filenamep = @"Зачисления\ПТО после 11_" + dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value.ToString();
                                string shablon = filenamep;
                                int count = 2;

                                while (File.Exists(Directory.GetCurrentDirectory() + @"\" + filenamep + ".doc"))
                                {

                                    filenamep = shablon + "_" + count.ToString();
                                    count++;
                                }
                                filenamep = filenamep + ".doc";
                                File.Copy(Directory.GetCurrentDirectory() + @"\Шаблоны\ПТО после 11.doc", Directory.GetCurrentDirectory() + @"\" + filenamep);
                                Object fileName = Directory.GetCurrentDirectory() + @"\" + filenamep;
                                Object missing = Type.Missing;
                                app.Documents.Open(ref fileName);
                                Word.Find find = app.Selection.Find;
                                Object wrap = Word.WdFindWrap.wdFindContinue;
                                Object replace = Word.WdReplace.wdReplaceAll;
                                con.Open();
                              
                                SqlCommand cmd = new SqlCommand("SELECT * FROM Passport WHERE [id абитуриента] = @dt1", con);
                                cmd.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                                da = new SqlDataAdapter("test", con);
                                da.SelectCommand = cmd;
                                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                                ds = new DataSet();
                                da.Fill(ds);
                                SqlCommand cmd2 = new SqlCommand("SELECT * FROM Parents WHERE [id абитуриента] = @dt1", con);
                                cmd2.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                                SqlDataAdapter da2 = new SqlDataAdapter("test", con);
                                da2.SelectCommand = cmd2;
                                SqlCommandBuilder cb2 = new SqlCommandBuilder(da);
                                DataSet ds2 = new DataSet();
                                int opek;
                                da2.Fill(ds2);
                                if (metroRadioButton1.Checked)
                                {
                                    opek = 0;
                                }
                                else
                                {
                                    opek = 1;
                                }

                                foreach (Shape sh in app.ActiveDocument.Shapes)
                                {
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.fio}", dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value.ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.adress}", ds.Tables[0].Rows[0][2].ToString());

                                    if (ds.Tables[0].Rows[0][3].ToString().Length != 0)
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn1}", ds.Tables[0].Rows[0][3].ToString().Substring(0, 2));
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn2}", ds.Tables[0].Rows[0][3].ToString().Substring(2, ds.Tables[0].Rows[0][3].ToString().Length - 2));
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn1}", "  ");
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn2}", "    ");
                                    }

                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasvidan}", ds.Tables[0].Rows[0][5].ToString());
                                    DateTime date;

                                    if (DateTime.TryParse(ds.Tables[0].Rows[0][6].ToString(), out date))
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasdate}", date.ToShortDateString());
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasdate}", "");
                                    }
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasln}", ds.Tables[0].Rows[0][4].ToString());

                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pfio}", ds2.Tables[0].Rows[opek][2].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.padr}", ds2.Tables[0].Rows[opek][4].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.ppasn}", ds2.Tables[0].Rows[opek][9].ToString());
                                    if (DateTime.TryParse(ds2.Tables[0].Rows[opek][11].ToString(), out date))
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.ppasdate}", date.ToShortDateString());
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.ppasdate}", " ");
                                    }
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.ppasvidan}", ds2.Tables[0].Rows[opek][10].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.ppasln}", ds2.Tables[0].Rows[opek][8].ToString());




                                }

                                find.Text = "{info.specialn}";
                                find.Replacement.Text = spec[metroComboBox1.SelectedIndex];
                                find.Execute(FindText: Type.Missing, Replace: replace);

                                find.Text = "{info.cvalif}";
                                find.Replacement.Text = kvalif[metroComboBox1.SelectedIndex];
                                find.Execute(FindText: Type.Missing, Replace: replace);

                                con.Close();
                                app.ActiveDocument.Save();
                                app.ActiveDocument.Close();

                                app.Quit();

                                break;
                            }

                        case "ССО":
                            {
                                string filenamep = @"Зачисления\ССО_" + dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value.ToString();
                                string shablon = filenamep;
                                int count = 2;

                                while (File.Exists(Directory.GetCurrentDirectory() + @"\" + filenamep + ".doc"))
                                {

                                    filenamep = shablon + "_" + count.ToString();
                                    count++;
                                }
                                filenamep = filenamep + ".doc";
                                File.Copy( Directory.GetCurrentDirectory() + @"\Шаблоны\ССО.doc", Directory.GetCurrentDirectory() + @"\" + filenamep);
                                app = new Word.Application();
                                Object fileName = Directory.GetCurrentDirectory() + @"\" + filenamep;
                                Object missing = Type.Missing;
                                app.Documents.Open(ref fileName);
                                Word.Find find = app.Selection.Find;
                                Object wrap = Word.WdFindWrap.wdFindContinue;
                                Object replace = Word.WdReplace.wdReplaceAll;
                                con.Open();
                                SqlCommand cmd = new SqlCommand("SELECT * FROM Passport WHERE [id абитуриента] = @dt1", con);
                                cmd.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                                da = new SqlDataAdapter("test", con);
                                da.SelectCommand = cmd;
                                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                                ds = new DataSet();
                                da.Fill(ds);
                                SqlCommand cmd2 = new SqlCommand("SELECT * FROM Parents WHERE [id абитуриента] = @dt1", con);
                                cmd2.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                                SqlDataAdapter da2 = new SqlDataAdapter("test", con);
                                da2.SelectCommand = cmd2;
                                SqlCommandBuilder cb2 = new SqlCommandBuilder(da);
                                DataSet ds2 = new DataSet();
                                int opek;
                                da2.Fill(ds2);
                                if (metroRadioButton1.Checked)
                                {
                                    opek = 0;
                                }
                                else
                                {
                                    opek = 1;
                                }

                                foreach (Shape sh in app.ActiveDocument.Shapes)
                                {
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.fio}", dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value.ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.adress}", ds.Tables[0].Rows[0][2].ToString());

                                    if (ds.Tables[0].Rows[0][3].ToString().Length != 0)
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn1}", ds.Tables[0].Rows[0][3].ToString().Substring(0, 2));
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn2}", ds.Tables[0].Rows[0][3].ToString().Substring(2, ds.Tables[0].Rows[0][3].ToString().Length - 2));
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn1}", "  ");
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn2}", "    ");
                                    }

                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasvid}", ds.Tables[0].Rows[0][5].ToString());
                                    DateTime date;
                                    if (DateTime.TryParse(ds.Tables[0].Rows[0][6].ToString(), out date))
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasdate}", date.ToShortDateString());
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasdate}", "");
                                    }
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasln}", ds.Tables[0].Rows[0][4].ToString());

                                    




                                }

                                find.Text = "{info.specialn}";
                                find.Replacement.Text = spec[metroComboBox1.SelectedIndex];
                                find.Execute(FindText: Type.Missing, Replace: replace);

                                find.Text = "{info.cvalif}";
                                find.Replacement.Text = kvalif[metroComboBox1.SelectedIndex];
                                find.Execute(FindText: Type.Missing, Replace: replace);

















                                con.Close();
                                app.ActiveDocument.Save();
                                app.ActiveDocument.Close();

                                app.Quit();

                                break;
                            }

                        case "Заявление ССО":
                            {
                                string filenamep = @"Зачисления\Заявление ССО_" + dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value.ToString();
                                string shablon = filenamep;
                                int count = 2;

                                while (File.Exists(Directory.GetCurrentDirectory() + @"\" + filenamep + ".doc"))
                                {

                                    filenamep = shablon + "_" + count.ToString();
                                    count++;
                                }
                                filenamep = filenamep + ".doc";
                                File.Copy(Directory.GetCurrentDirectory() + @"\Шаблоны\Заявление ССО.doc", Directory.GetCurrentDirectory() + @"\" + filenamep);
                                app = new Word.Application();
                                Object fileName = Directory.GetCurrentDirectory() + @"\" + filenamep;
                                Object missing = Type.Missing;
                                app.Documents.Open(ref fileName);
                                Word.Find find = app.Selection.Find;
                                Object wrap = Word.WdFindWrap.wdFindContinue;
                                Object replace = Word.WdReplace.wdReplaceAll;
                                con.Open();
                                SqlCommand cmd = new SqlCommand("SELECT * FROM Passport WHERE [id абитуриента] = @dt1", con);
                                cmd.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                                da = new SqlDataAdapter("test", con);
                                da.SelectCommand = cmd;
                                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                                ds = new DataSet();
                                da.Fill(ds);
                                SqlCommand cmd2 = new SqlCommand("SELECT * FROM Parents WHERE [id абитуриента] = @dt1", con);
                                cmd2.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                                SqlDataAdapter da2 = new SqlDataAdapter("test", con);
                                da2.SelectCommand = cmd2;
                                SqlCommandBuilder cb2 = new SqlCommandBuilder(da);
                                DataSet ds2 = new DataSet();
                                int opek;
                                da2.Fill(ds2);
                                if (metroRadioButton1.Checked)
                                {
                                    opek = 0;
                                }
                                else
                                {
                                    opek = 1;
                                }

                                string obsh;
                                if ((bool)dataGridView1[7, dataGridView1.CurrentCell.RowIndex].Value)
                                {
                                    obsh = "да";
                                }
                                else
                                {
                                    obsh = "нет";
                                }

                                foreach (Shape sh in app.ActiveDocument.Shapes)
                                {
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.fio}", dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value.ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.adress}", ds.Tables[0].Rows[0][2].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.lgot}", dataGridView1[5, dataGridView1.CurrentCell.RowIndex].Value.ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.obsh}", obsh);


                                    DateTime date;
                                    if (DateTime.TryParse(ds.Tables[0].Rows[0][7].ToString(), out date))
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.dater}", date.ToShortDateString());
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.dater}", "");
                                    }
                                   

                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasn}", ds.Tables[0].Rows[0][3].ToString());
                                   

                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasvid}", ds.Tables[0].Rows[0][5].ToString());
                                   
                                    if (DateTime.TryParse(ds.Tables[0].Rows[0][6].ToString(), out date))
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasdate}", date.ToShortDateString());
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasdate}", " ");
                                    }
                                    
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pasln}", ds.Tables[0].Rows[0][4].ToString());


                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pdadfio}", ds2.Tables[0].Rows[1][2].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pdadadr}", ds2.Tables[0].Rows[1][4].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pdadtel}", ds2.Tables[0].Rows[1][6].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pmamfio}", ds2.Tables[0].Rows[0][2].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pmamadr}", ds2.Tables[0].Rows[0][4].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pmamtel}", ds2.Tables[0].Rows[0][6].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.specialn}", spec[metroComboBox1.SelectedIndex]);





                                }

                               

















                                con.Close();
                                app.ActiveDocument.Save();
                                app.ActiveDocument.Close();

                                app.Quit();

                                break;
                            }

                        case "Заявление ПТО":
                            {
                                string filenamep = @"Зачисления\Заявление ПТО_" + dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value.ToString();
                                string shablon = filenamep;
                                int count = 2;

                                while (File.Exists(Directory.GetCurrentDirectory() + @"\" + filenamep + ".doc"))
                                {

                                    filenamep = shablon + "_" + count.ToString();
                                    count++;
                                }
                                filenamep = filenamep + ".doc";
                                File.Copy(Directory.GetCurrentDirectory() + @"\Шаблоны\Заявление ПТО.doc", Directory.GetCurrentDirectory() + @"\" + filenamep);
                                Object fileName = Directory.GetCurrentDirectory() + @"\" + filenamep;
                                Object missing = Type.Missing;
                                app.Documents.Open(ref fileName);
                                Word.Find find = app.Selection.Find;
                                Object wrap = Word.WdFindWrap.wdFindContinue;
                                Object replace = Word.WdReplace.wdReplaceAll;

                                con.Open();
                                SqlCommand cmd = new SqlCommand("SELECT * FROM Passport WHERE [id абитуриента] = @dt1", con);
                                cmd.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                                da = new SqlDataAdapter("test", con);
                                da.SelectCommand = cmd;
                                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                                ds = new DataSet();
                                da.Fill(ds);
                                SqlCommand cmd2 = new SqlCommand("SELECT * FROM Parents WHERE [id абитуриента] = @dt1", con);
                                cmd2.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                                SqlDataAdapter da2 = new SqlDataAdapter("test", con);
                                da2.SelectCommand = cmd2;
                                SqlCommandBuilder cb2 = new SqlCommandBuilder(da);
                                DataSet ds2 = new DataSet();
                                int opek;
                                da2.Fill(ds2);
                                if (metroRadioButton1.Checked)
                                {
                                    opek = 0;
                                }
                                else
                                {
                                    opek = 1;
                                }

                                string obsh;
                                if ((bool)dataGridView1[7, dataGridView1.CurrentCell.RowIndex].Value)
                                {
                                    obsh = "да";
                                }
                                else
                                {
                                    obsh = "нет";
                                }

                                foreach (Shape sh in app.ActiveDocument.Shapes)
                                {
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.fio}", dataGridView1[1, dataGridView1.CurrentCell.RowIndex].Value.ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.yasik}", dataGridView1[9, dataGridView1.CurrentCell.RowIndex].Value.ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.lgot}", dataGridView1[5, dataGridView1.CurrentCell.RowIndex].Value.ToString());
                                    if (obsh == "да")
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.obshy}", "_______________________");
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.obshn}", "");
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.obshn}", "_______________________");
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.obshy}", "");
                                    }
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.adress}", ds.Tables[0].Rows[0][2].ToString());
                                    DateTime date;
                                    if (DateTime.TryParse(ds.Tables[0].Rows[0][7].ToString(), out date))
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.dater}", date.ToShortDateString());
                                    }
                                    else
                                    {
                                        sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.dater}", "");
                                    }


                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pdadfio}", ds2.Tables[0].Rows[1][2].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pdadadr}", ds2.Tables[0].Rows[1][4].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pdadtel}", ds2.Tables[0].Rows[1][6].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pmamfio}", ds2.Tables[0].Rows[0][2].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pmamadr}", ds2.Tables[0].Rows[0][4].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.pmamtel}", ds2.Tables[0].Rows[0][6].ToString());
                                    sh.TextFrame.TextRange.Text = sh.TextFrame.TextRange.Text.Replace("{info.specialn}", spec[metroComboBox1.SelectedIndex]);




                                }

                               


                                con.Close();
                                app.ActiveDocument.Save();
                                app.ActiveDocument.Close();

                                app.Quit();

                                break;
                            }
                       

                    }
                    MessageBox.Show("Документ успешно создан");
                }
                catch (Exception ex)
                {
                    con.Close();
                    app.ActiveDocument.Save();
                    app.ActiveDocument.Close();

                    app.Quit();
                    MessageBox.Show("Произошла ошибка: " + ex.Message);
                    
                }
                
            
            }
            else
            {
                MessageBox.Show("Абитуриент не выбран");
            }
            
       
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                con.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM Parents WHERE [id абитуриента] = @dt1", con);
                cmd.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                da = new SqlDataAdapter("test", con);
                da.SelectCommand= cmd;
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                ds = new DataSet();
                da.Fill(ds);
                DateTime date = (DateTime)ds.Tables[0].Rows[0][3];
                MessageBox.Show("Имя: "+ ds.Tables[0].Rows[0][2].ToString() + "\nАдрес: " + ds.Tables[0].Rows[0][4].ToString() + "\nДата рождения: " + date.ToShortDateString() + "\nПримечание: " + ds.Tables[0].Rows[0][7].ToString());
                con.Close();
            }
            else
            {
                MessageBox.Show("Абитуриент не выбран");
            }
         }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                con.Open();
                SqlCommand cmd = new SqlCommand("SELECT * FROM Parents WHERE [id абитуриента] = @dt1", con);
                cmd.Parameters.Add("@dt1", dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value);
                da = new SqlDataAdapter("test", con);
                da.SelectCommand = cmd;
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                ds = new DataSet();
                da.Fill(ds);
                DateTime date = (DateTime)ds.Tables[0].Rows[1][3];
                MessageBox.Show("Имя: " + ds.Tables[0].Rows[1][2].ToString() + "\nАдрес: " + ds.Tables[0].Rows[1][4].ToString() + "\nДата рождения: " + date.ToShortDateString() + "\nПримечание: " + ds.Tables[0].Rows[1][7].ToString());
                con.Close();
            }
            else
            {
                MessageBox.Show("Абитуриент не выбран");
            }
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count != 0)
            {
                try
            {
                string filenamep = @"Зачисления\Отчет";
                string shablon = filenamep;
                int count = 2;

                while (File.Exists(Directory.GetCurrentDirectory() + @"\" + filenamep + ".doc"))
                {

                    filenamep = shablon + "_" + count.ToString();
                    count++;
                }
                filenamep = filenamep + ".doc";
                File.Copy(@"Шаблоны\Отчет.doc", Directory.GetCurrentDirectory() + @"\" + filenamep);
                if (!sso[metroComboBox1.SelectedIndex])
                {

                    Word.Application app = new Word.Application();
                    Object fileName = Directory.GetCurrentDirectory() + @"\" + filenamep;
                    Object missing = Type.Missing;
                    app.Documents.Open(ref fileName);

                    var Paragraph = app.ActiveDocument.Paragraphs.Add();
                    //Получаем диапазон
                    var tableRange = Paragraph.Range;
                    //Добавляем таблицу 2х2 в указаный диапазон
                    con.Open();
                    da = new SqlDataAdapter("SELECT COUNT(*) FROM groups WHERE groups.ССО = 0", con);
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    ds = new DataSet();
                    da.Fill(ds);
                    if (int.Parse(ds.Tables[0].Rows[0][0].ToString()) == 0)
                    {
                        app.ActiveDocument.Save();
                        app.ActiveDocument.Close();
                        con.Close();
                        app.Quit();
                        MessageBox.Show("Данные пустые");
                    }
                    else
                    {
                        app.ActiveDocument.Tables.Add(tableRange, int.Parse(ds.Tables[0].Rows[0][0].ToString()) + 1, 7);

                        var table = app.ActiveDocument.Tables[app.ActiveDocument.Tables.Count];
                        table.set_Style("Сетка таблицы");
                        table.ApplyStyleHeadingRows = true;
                        table.ApplyStyleLastRow = false;
                        table.ApplyStyleFirstColumn = true;
                        table.ApplyStyleLastColumn = false;
                        table.ApplyStyleRowBands = true;
                        table.ApplyStyleColumnBands = false;

                        Word.Cell cell = app.ActiveDocument.Tables[1].Cell(1, 1);
                        cell.Range.Text = "Квалификация";
                        cell = app.ActiveDocument.Tables[1].Cell(1, 2);
                        cell.Range.Text = "Группа";
                        cell = app.ActiveDocument.Tables[1].Cell(1, 3);
                        cell.Range.Text = "План приёма";
                        cell = app.ActiveDocument.Tables[1].Cell(1, 4);
                        cell.Range.Text = "Подано заявок";
                        cell = app.ActiveDocument.Tables[1].Cell(1, 5);
                        cell.Range.Text = "Заявок на целевое обучение";
                        cell = app.ActiveDocument.Tables[1].Cell(1, 6);
                        cell.Range.Text = "Заявок с льготами на зачисление";
                        cell = app.ActiveDocument.Tables[1].Cell(1, 7);
                        cell.Range.Text = "Значение среднего бала (минимальное - максимальное)";


                        da = new SqlDataAdapter("SELECT groups.ССО, groups.Квалификация ,groups.[Номер группы], groups.[Численность группы] , Count(abit.[Документ об образовании]) as 'подано заявок', (Select COUNT(*) FROM abit WHERE abit.[Номер группы] = groups.[Номер группы] and [Целевое обучение] = 1) as 'целевики', (Select COUNT(*) FROM abit WHERE abit.[Номер группы] = groups.[Номер группы] and Льготы <> '-' and Льготы <> 'нет') as 'льготники', CONCAT('Макс - ',max(abit.[ср. балл]),' \nМин -', min(abit.[ср. балл]))  FROM groups LEFT JOIN abit on abit.[Номер группы] = groups.[Номер группы] Group by groups.[Номер группы], groups.Квалификация, groups.[Численность группы], groups.ССО HAVING groups.ССО = 0", con);
                        cb = new SqlCommandBuilder(da);
                        ds = new DataSet();
                        da.Fill(ds);

                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            for (int j = 1; j < ds.Tables[0].Columns.Count; j++)
                            {
                                cell = app.ActiveDocument.Tables[1].Cell(i + 2, j);
                                cell.Range.Text = ds.Tables[0].Rows[i][j].ToString();
                            }
                        }


                        app.ActiveDocument.Save();
                        app.ActiveDocument.Close();
                        con.Close();
                        app.Quit();
                    }
                }
                else
                {
                    Word.Application app = new Word.Application();
                    Object fileName = Directory.GetCurrentDirectory() + @"\" + filenamep;
                    Object missing = Type.Missing;
                    app.Documents.Open(ref fileName);

                    var Paragraph = app.ActiveDocument.Paragraphs.Add();
                    //Получаем диапазон
                    var tableRange = Paragraph.Range;
                    //Добавляем таблицу 2х2 в указаный диапазон
                    con.Open();
                    da = new SqlDataAdapter("SELECT COUNT(*) FROM groups WHERE groups.ССО = 1", con);
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    ds = new DataSet();
                    da.Fill(ds);
                    if (int.Parse(ds.Tables[0].Rows[0][0].ToString()) == 0)
                    {
                        app.ActiveDocument.Save();
                        app.ActiveDocument.Close();
                        con.Close();
                        app.Quit();
                        MessageBox.Show("Данные пустые");
                    }
                    else
                    {
                        app.ActiveDocument.Tables.Add(tableRange, 21, int.Parse(ds.Tables[0].Rows[0][0].ToString()) + 2);

                        var table = app.ActiveDocument.Tables[app.ActiveDocument.Tables.Count];
                        table.set_Style("Сетка таблицы");
                        table.ApplyStyleHeadingRows = true;
                        table.ApplyStyleLastRow = false;
                        table.ApplyStyleFirstColumn = true;
                        table.ApplyStyleLastColumn = false;
                        table.ApplyStyleRowBands = true;
                        table.ApplyStyleColumnBands = false;
                        Word.Cell cell = app.ActiveDocument.Tables[1].Cell(1, 1);
                        cell.Range.Text = "Квалификация";
                        cell = app.ActiveDocument.Tables[1].Cell(1, app.ActiveDocument.Tables[1].Columns.Count);
                        cell.Range.Text = "ИТОГО";
                        cell = app.ActiveDocument.Tables[1].Cell(2, 1);
                        cell.Range.Text = "Группа";
                        cell = app.ActiveDocument.Tables[1].Cell(3, 1);
                        cell.Range.Text = "План приёма";
                        cell = app.ActiveDocument.Tables[1].Cell(4, 1);
                        cell.Range.Text = "Подано заявок";
                        cell = app.ActiveDocument.Tables[1].Cell(5, 1);
                        cell.Range.Text = "Вне конкурса";
                        cell = app.ActiveDocument.Tables[1].Cell(6, 1);
                        cell.Range.Text = "Баллы:";
                        cell = app.ActiveDocument.Tables[1].Cell(7, 1);
                        cell.Range.Text = "до 3";
                        cell = app.ActiveDocument.Tables[1].Cell(8, 1);
                        cell.Range.Text = "3,1 - 3,5";
                        cell = app.ActiveDocument.Tables[1].Cell(9, 1);
                        cell.Range.Text = "3,6 - 4,0";
                        cell = app.ActiveDocument.Tables[1].Cell(10, 1);
                        cell.Range.Text = "4,1 - 4,5";

                        cell = app.ActiveDocument.Tables[1].Cell(11, 1);
                        cell.Range.Text = "4,6 - 5,0";
                        cell = app.ActiveDocument.Tables[1].Cell(12, 1);
                        cell.Range.Text = "5,1 - 5,5";
                        cell = app.ActiveDocument.Tables[1].Cell(13, 1);
                        cell.Range.Text = "5,6 - 6,0";
                        cell = app.ActiveDocument.Tables[1].Cell(14, 1);
                        cell.Range.Text = "6,1 - 6,5";
                        cell = app.ActiveDocument.Tables[1].Cell(15, 1);
                        cell.Range.Text = "6,6 - 7,0";
                        cell = app.ActiveDocument.Tables[1].Cell(16, 1);
                        cell.Range.Text = "7,1 - 7,5";
                        cell = app.ActiveDocument.Tables[1].Cell(17, 1);
                        cell.Range.Text = "7,6 - 8,0";
                        cell = app.ActiveDocument.Tables[1].Cell(18, 1);
                        cell.Range.Text = "8,1 - 8,5";
                        cell = app.ActiveDocument.Tables[1].Cell(19, 1);
                        cell.Range.Text = "8,6 - 9,0";
                        cell = app.ActiveDocument.Tables[1].Cell(20, 1);
                        cell.Range.Text = "9,1 - 9,5";
                        cell = app.ActiveDocument.Tables[1].Cell(21, 1);
                        cell.Range.Text = "9,6 - 10";



                        da = new SqlDataAdapter("SELECT groups.ССО, groups.Квалификация ,groups.[Номер группы], groups.[Численность группы] , Count(abit.[Документ об образовании]) as 'подано заявок', (Select COUNT(*) FROM abit WHERE abit.[Номер группы] = groups.[Номер группы] and [Целевое обучение] = 1) as 'Вне конкурса','__', (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] <= 3 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 3.1 and abit.[ср. балл] <= 3.5 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 3.6 and abit.[ср. балл] <= 4.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 4.1 and abit.[ср. балл] <= 4.5 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 4.6 and abit.[ср. балл] <= 5.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 5.1 and abit.[ср. балл] <= 5.5 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 5.5 and abit.[ср. балл] <= 6.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 6.1 and abit.[ср. балл] <= 6.5 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 6.6 and abit.[ср. балл] <= 7.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 7.1 and abit.[ср. балл] <= 7.5 and abit.[Номер группы] = groups.[Номер группы]),(SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 7.6 and abit.[ср. балл] <= 8.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 8.1 and abit.[ср. балл] <= 8.5 and abit.[Номер группы] = groups.[Номер группы]),(SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 8.6 and abit.[ср. балл] <= 9.0 and abit.[Номер группы] = groups.[Номер группы]), (SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 9.0 and abit.[ср. балл] <= 9.5 and abit.[Номер группы] = groups.[Номер группы]),(SELECT COUNT(*) FROM abit WHERE abit.[ср. балл] >= 9.6 and abit.[Номер группы] = groups.[Номер группы])  FROM groups LEFT JOIN abit on abit.[Номер группы] = groups.[Номер группы] Group by groups.[Номер группы], groups.Квалификация, groups.[Численность группы], groups.ССО HAVING groups.ССО = 1", con);
                        cb = new SqlCommandBuilder(da);
                        ds = new DataSet();
                        da.Fill(ds);

                        for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                        {
                            for (int j = 1; j < ds.Tables[0].Columns.Count; j++)
                            {
                                if (j == 6)
                                { continue; }
                                cell = app.ActiveDocument.Tables[1].Cell(j, i + 2);
                                cell.Range.Text = ds.Tables[0].Rows[i][j].ToString();
                            }
                        }

                        int[] markscount = new int[15] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
                        for (int i1 = 0; i1 < ds.Tables[0].Rows.Count; i1++)
                        {
                            for (int j1 = 7; j1 < ds.Tables[0].Columns.Count; j1++)
                            {
                                if (ds.Tables[0].Rows[i1][j1].ToString() == "1")
                                {
                                    markscount[j1 - 7]++;
                                }
                            }
                        }

                        for (int i2 = 7; i2 <= app.ActiveDocument.Tables[1].Rows.Count; i2++)
                        {
                            cell = app.ActiveDocument.Tables[1].Cell(i2, app.ActiveDocument.Tables[1].Columns.Count);
                            cell.Range.Text = markscount[i2 - 7].ToString();
                        }

                        da = new SqlDataAdapter("SELECT  (SELECT SUM([Численность группы]) FROM groups where groups.ССО = 1), (SELECT COUNT(*) FROM abit LEFT JOIN groups on abit.[Номер группы] = groups.[Номер группы] group by groups.ССО HAVING groups.ССО = 1)", con);
                        cb = new SqlCommandBuilder(da);
                        ds = new DataSet();
                        da.Fill(ds);

                        cell = app.ActiveDocument.Tables[1].Cell(3, app.ActiveDocument.Tables[1].Columns.Count);
                        cell.Range.Text = ds.Tables[0].Rows[0][0].ToString();
                        cell = app.ActiveDocument.Tables[1].Cell(4, app.ActiveDocument.Tables[1].Columns.Count);
                        cell.Range.Text = ds.Tables[0].Rows[0][1].ToString();



                        app.ActiveDocument.Save();
                        app.ActiveDocument.Close();
                        con.Close();
                        app.Quit();
                    }

                }
                MessageBox.Show("Докумемнт успешно создан");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());


            }
        }
        else
            {
                MessageBox.Show("Абитуриент не выбран");
            }

}

    }
}
