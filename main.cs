using MetroFramework;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Комиссия
{
    public partial class main : MetroFramework.Forms.MetroForm
    {
        public string pas = string.Empty;
        public main()
        {
            InitializeComponent();
            this.StyleManager = metroStyleManager1;
            metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            SoundPlayer simpleSound = new SoundPlayer("EO.WinForm.Ac1.wav");
            metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
            int x = this.Size.Height - 90;
            int oneHeig = x / 5;
            metroButton1.Location = new Point(metroButton1.Location.X, 35);
            metroButton3.Location = new Point(metroButton1.Location.X, metroButton1.Location.Y + oneHeig + 10);
            metroButton2.Location = new Point(metroButton1.Location.X, metroButton3.Location.Y + oneHeig + 10);
            metroButton5.Location = new Point(metroButton1.Location.X, metroButton2.Location.Y + oneHeig + 10);
            metroButton4.Location = new Point(metroButton1.Location.X, metroButton5.Location.Y + oneHeig + 10);
            metroButton1.Height = oneHeig;
            metroButton2.Height = oneHeig;
            metroButton3.Height = oneHeig;
            metroButton4.Height = oneHeig;
            metroButton5.Height = oneHeig;
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
            }
            else
            {
                metroStyleManager1.Theme = MetroThemeStyle.Light;
                metroStyleManager1.Style = MetroColorStyle.Teal;
                this.Theme = metroStyleManager1.Theme;
                this.Style = metroStyleManager1.Style;
                Properties.Settings.Default.themeSwitch = false;
            }
            Properties.Settings.Default.Save();
        }

        private void metroButton2_Click(object sender, EventArgs e)
        {
           zachislit fm = new zachislit(this);
            fm.Show();
            fm.WindowState = this.WindowState;
            fm.Location = this.Location;
            this.Hide();
        }

        private void metroButton4_Click(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }

        private void metroButton3_Click(object sender, EventArgs e)
        {
            Groups fm = new Groups(this);
            fm.Show();
            fm.WindowState = this.WindowState;
            fm.Location = this.Location;
            this.Hide();
        }

        private void metroButton1_Click(object sender, EventArgs e)
        {
            Абитуриенты fm = new Абитуриенты(this);
            fm.Show();
            fm.WindowState = this.WindowState;
            fm.Location = this.Location;
            this.Hide();
        }

        private void main_Load(object sender, EventArgs e)
        {

        }

        private void main_FormClosed(object sender, FormClosedEventArgs e)
        {
            Environment.Exit(0);
        }

        private void metroButton5_Click(object sender, EventArgs e)
        {
            Settings fm = new Settings(this);
            fm.Show();
            fm.WindowState = this.WindowState;
            fm.Location = this.Location;
            this.Hide();
        }

        private void metroToggle1_VisibleChanged(object sender, EventArgs e)
        {
            metroToggle1.Checked = Properties.Settings.Default.themeSwitch;
        }

        private void main_Resize(object sender, EventArgs e)
        {
            int x = this.Size.Height - 90;
            int oneHeig = x / 5;
            metroButton1.Location = new Point(metroButton1.Location.X, 35);
            metroButton3.Location = new Point(metroButton1.Location.X, metroButton1.Location.Y + oneHeig + 10);
            metroButton2.Location = new Point(metroButton1.Location.X, metroButton3.Location.Y + oneHeig + 10);
            metroButton5.Location = new Point(metroButton1.Location.X, metroButton2.Location.Y + oneHeig + 10);
            metroButton4.Location = new Point(metroButton1.Location.X, metroButton5.Location.Y + oneHeig + 10);
            metroButton1.Height = oneHeig;
            metroButton2.Height = oneHeig;
            metroButton3.Height = oneHeig;
            metroButton4.Height = oneHeig;
            metroButton5.Height = oneHeig;

        }

        private void main_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.F10)
                Process.Start(Directory.GetCurrentDirectory() + @"\Справка.pdf");
        }
    }
}
