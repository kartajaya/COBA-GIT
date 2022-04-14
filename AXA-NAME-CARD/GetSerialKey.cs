using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Management;
using System.Management.Instrumentation;
using System.IO;

namespace AXA_NAME_CARD
{
    public partial class GetSerialKey : Form
    {
        axaproc proc = new axaproc();
        public GetSerialKey()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!File.Exists("InfoSn.TXT"))
            {
                MessageBox.Show("Please put File InfoSn.TXT in main Folder", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string sn = "";
            string[] sx = File.ReadAllLines("InfoSn.txt");
            for (int jj = 0; jj < sx.Length; jj++)
            {
                sn = sn+sx[jj].Trim();
            }
            string serial = proc.Encrypt(sn, "-");
            textBox1.Text = serial;
            string fout = Directory.GetCurrentDirectory() + "\\SerialNum.TXT";
            using (System.IO.StreamWriter fs = new System.IO.StreamWriter(fout, false))
            {
                fs.WriteLine(serial);
                fs.Close();
            }
            MessageBox.Show("Get Serial Succesfull", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            Application.Exit();
        }
    }
}
