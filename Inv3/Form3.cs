using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Inv3
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        public string partNumber
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }

        public string KMZ
        {
            get { return textBox3.Text; }
            set { textBox3.Text = value; }
        }

        public string partName
        {
            get { return textBox2.Text; }
            set { textBox2.Text = value; }
        }

        public int partCount
        {
            get { return ((int)numericUpDown1.Value); }
            set { numericUpDown1.Value = value; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ( (textBox1.Text.Length == 0) && (textBox3.Text.Length == 0) )
            {
                MessageBox.Show("JE POTŘEBA  ''K M Z''  NEBO  ''P A R T N U M B E R''\r\n" +
                    "\r\n" +
                    "ZADEJ ZNOVU ALESPOŇ JEDNU Z TĚCHTO HODNOT\r\n", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button1.DialogResult = DialogResult.None;
                this.DialogResult = DialogResult.None;
                textBox1.Focus();
            }
            else
            {
                button1.DialogResult = DialogResult.OK;
                this.DialogResult = DialogResult.OK;
            }
        }
    }
}
