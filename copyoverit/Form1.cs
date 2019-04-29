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
using System.Diagnostics;
using System.Timers;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
namespace copyoverit
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        Excel.Application aplika = new Excel.Application();
        bool overwrite;
        Excel.Worksheet sh;
        Excel.Worksheet a;

        private void button2_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
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
            ofd.Filter = "Excel Files(.xls)|*.xls|Excel Files(.xlsx)|*.xlsx|Excel Files(.xlsm)|*.xlsm";
            ofd.InitialDirectory = "c:\\matriks\\user\\reports\\excel";
            ofd.ShowDialog();
            txtSource.Text = ofd.FileName;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Files(.xls)|*.xls|Excel Files(.xlsx)|*.xlsx|Excel Files(.xlsm)|*.xlsm";
            sfd.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            sfd.ShowDialog();
            txtDest.Text = sfd.FileName;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Thread.Sleep(Convert.ToInt16(textBox1.Text)*1000);
            Stopwatch sw=new Stopwatch();
            sw.Start();
            Excel.Workbook asil = aplika.Workbooks.Open(txtSource.Text);
            Excel.Workbook hedef = aplika.Workbooks.Open(txtDest.Text);
            sh = asil.Worksheets.get_Item(1);
            a = hedef.Worksheets.get_Item(1);
            Excel.Range kaynak = sh.Range["A1:O100"];
            Excel.Range hedefr = a.Range["A1"];
            sh.Copy(hedefr);

            asil.Close(false);
            hedef.Close(overwrite);
            aplika.Quit();
            sw.Stop();
            TimeSpan ts = sw.Elapsed;
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            MessageBox.Show(elapsedTime);
            timer1.Enabled = false;
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
