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
        }

        public string partName
        {
            get { return textBox2.Text; }
        }

        public int partCount
        {
            get { return ((int)numericUpDown1.Value); }
        }


    }
}
