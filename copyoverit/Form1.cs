using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Timers;

namespace copyoverit
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        bool overwrite;

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            timer1.Interval = Convert.ToInt32(txtTime.Text)*1000;
            timer1.Enabled = !timer1.Enabled;
            if (timer1.Enabled)
            {
                label6.Text = "Running";
                btnStart.Text = "Stop";
            }
            else
            {
                label6.Text = "Stopped";
                btnStart.Text = "Start";
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files(.xls)|*.xls|Excel Files(.xlsx)|*.xlsx|Excel Files(.xlsm)|*.xlsm|all files(*.*)|*.*";
            ofd.InitialDirectory = "c:\\matriks\\user\\reports\\excel";
            ofd.ShowDialog();
            txtSource.Text = ofd.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Files(.xls)|*.xls|Excel Files(.xlsx)|*.xlsx|Excel Files(.xlsm)|*.xlsm|all files(*.*)|*.*";
            sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            sfd.ShowDialog();
            txtDest.Text = sfd.FileName;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            File.Copy(txtSource.Text, txtDest.Text, true);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                overwrite = true;
            }
            else
            {
                overwrite = false;
            }
        }
    }
}
