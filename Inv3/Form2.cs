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

    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
            textBox1.Text = "";
        }

        public string newNodeText
        {
            get { return textBox1.Text; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ( textBox1.Text.Length > 0 )
            {
                button1.DialogResult = DialogResult.OK;
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("MUSÍŠ ZADAT NĚJAKÝ NÁZEV PRO NOVOU POZICI", "CHYBÍ NÁZEV", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox1.Focus();
                button1.DialogResult = DialogResult.None;
                this.DialogResult = DialogResult.None;
            }
        }
    }
}
