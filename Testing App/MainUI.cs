using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using CDEAutomation.classes;
using System.Configuration;

namespace CDEAutomation
{
    public partial class MainUI : Form
    {
        public MainUI()
        {
            InitializeComponent();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            CDEAutomation.Properties.Settings.Default["TXTBOX1"] = textBox1.Text;
            CDEAutomation.Properties.Settings.Default["TXTBOX2"] = textBox2.Text;
            CDEAutomation.Properties.Settings.Default["TXTBOX3"] = textBox3.Text;
            CDEAutomation.Properties.Settings.Default.Save();

            if (string.IsNullOrEmpty(textBox1.Text) || string.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("Please select configuration and attributes files");
            }
            else
            {
                xmlHelper.location = textBox1.Text;
              

                if (!string.IsNullOrEmpty(textBox2.Text))
                {
                    appMethods.xlsLoc = textBox2.Text;

                }
                if (!string.IsNullOrEmpty(textBox3.Text))
                {
                  log.location = textBox3.Text;

                }
                //Application.ExitThread();
                appMethods.runCDEAuto();
                Application.ExitThread();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Show the dialog and get result.
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                textBox1.Text = openFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Show the dialog and get result.
            DialogResult result = openFileDialog2.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                textBox2.Text = openFileDialog2.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Show the dialog and get result.
            DialogResult result = openFileDialog3.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                textBox3.Text = openFileDialog3.FileName;
            }
        }

        private void MainUI_Load(object sender, EventArgs e)
        {
            textBox1.Text = CDEAutomation.Properties.Settings.Default["TXTBOX1"].ToString();
            textBox2.Text = CDEAutomation.Properties.Settings.Default["TXTBOX2"].ToString();
            textBox3.Text = CDEAutomation.Properties.Settings.Default["TXTBOX3"].ToString();
        }


    }
}
