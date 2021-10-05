using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.IO;

namespace Inv3
{
    public partial class Form1 : Form
    {

        public struct PART
        {
            public String ParentTag;
            public String TreeFullPath;
            public String Name;
            public String PartNumber;
            public String KMZ;
            public String Count;
        }

        public List<PART> parts = new List<PART>();

        FileStream file;
        StreamWriter writer;
        StreamReader reader;

        private void SaveTreeNonRecursive(TreeNode treeNode)
        {
            if (treeNode != null)
            {
                //Using a queue to store and process each node in the TreeView
                Queue<TreeNode> staging = new Queue<TreeNode>();
                staging.Enqueue(treeNode);

                while (staging.Count > 0)
                {
                    treeNode = staging.Dequeue();

                    //writer.WriteLine(treeNode.Tag);
                    writer.WriteLine(">");
                    writer.WriteLine(treeNode.Text);

                    foreach (TreeNode node in treeNode.Nodes)
                    {
                        staging.Enqueue(node);
                    }
                }
            }
        }



        public void SaveTree()
        {
            String FileName = Application.StartupPath.ToString() + @"\\save.ini";
            if (File.Exists(FileName))
                File.Delete(FileName);
            file = new FileStream(FileName, FileMode.CreateNew);
            file.Close();
            file.Dispose();
            writer = new StreamWriter(FileName);
            foreach (TreeNode n in treeView1.Nodes)
            {
                SaveTreeNonRecursive(n);
            }
            writer.Close();
            writer.Dispose();
        }

        public void LoadTree()
        {
            String _LineName = "";
            String _LineTag = "";
            List<TreeNode> nodes = new List<TreeNode>();
            String FileName = Application.StartupPath.ToString() + @"\\save.ini";
            //file = new FileStream(FileName, FileMode.CreateNew);
            //file.Close();
            //file.Dispose();
            reader = new StreamReader(FileName);
            int mode = 0;
            while ((_LineName = reader.ReadLine()) != null)
            {
                switch(mode)
                {
                    case 0:
                        {
                            break;
                        }
                    case 1:
                        {
                            break;
                        }

                }
                _LineTag = reader.ReadLine();
                TreeNode nod = new TreeNode();
                nod.Text = _LineName;
                nod.Tag = _LineTag;
                nodes.Add(nod);
                _LineName = "";
                _LineTag = "";
            }
            reader.Close();
            reader.Dispose();


        }


        public Form1()
        {
            InitializeComponent();

            
            treeView1.BeginUpdate();
                        
            
            TreeNode node = new TreeNode();
            TreeNode node1 = new TreeNode();
            TreeNode node2 = new TreeNode();
            TreeNode node3 = new TreeNode();
            node.Text = "SKLAD";
            node.Tag = "ROOT";
            treeView1.Nodes.Add(node);
            node1.Text = "10";
            node1.Tag = "ROOT";
            treeView1.Nodes[0].Nodes.Add(node1);
            node2.Text = "15";
            node2.Tag = "ROOT";
            treeView1.Nodes[0].Nodes.Add(node2);
            node3.Text = "20";
            node3.Tag = "15";
            treeView1.Nodes[0].Nodes[1].Nodes.Add(node3);
            

            treeView1.EndUpdate();
            treeView1.ExpandAll();

            SaveTree();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Pred zobrazenim Form1 otevreme COM port
            serialPort1.Open();



        }


        private void button1_Click(object sender, EventArgs e)
        {
            Form2 form = new Form2();
            TreeNode addTree = new TreeNode();
            form.ShowDialog();
            if (form.DialogResult == DialogResult.OK)
            {
                if (treeView1.Nodes.Count > 0)
                {
                    addTree.Text = form.newNodeText;
                    addTree.Tag = form.newNodeText;
                    treeView1.SelectedNode.Nodes.Add(addTree);
                    treeView1.SelectedNode.Expand();
                    treeView1.SelectedNode = addTree;
                    treeView1.Focus();
                }
                else
                {
                    addTree.Text = form.newNodeText;
                    addTree.Tag = "ROOT";
                    treeView1.Nodes.Add(addTree);
                    treeView1.SelectedNode = addTree;
                    treeView1.Focus();
                }
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            if ( treeView1.SelectedNode.Tag.ToString() == "ROOT" )
            {
                button2.Enabled = false;
            }
            else
            {
                button2.Enabled = true;
            }
            String countPositions = "";
            countPositions = treeView1.SelectedNode.Nodes.Count.ToString();
            countPositions += " / ";
            countPositions += treeView1.SelectedNode.GetNodeCount(true).ToString();
            countPositions += "";
            label4.Text = countPositions.ToString();

            dataGridView1.Rows.Clear();

            int index = 0;
            int count = 0;
            for ( int i = 0; i < parts.Count; i++ )
            {
                if (treeView1.SelectedNode.FullPath.Equals(parts[i].TreeFullPath))
                {
                    index = dataGridView1.Rows.Add();
                    dataGridView1.Rows[index].Cells[0].Value = parts[i].PartNumber;
                    dataGridView1.Rows[index].Cells[1].Value = parts[i].Name;
                    dataGridView1.Rows[index].Cells[2].Value = parts[i].Count;
                    count++;
                    //dataGridView1.Rows[index].Selected = true;
                }
            }

            String countParts = "";
            countParts = count.ToString();
            countParts += " / ";
            countParts += parts.Count.ToString();
            countParts += "";
            label3.Text = countParts.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("POZICE:>   ''"+treeView1.SelectedNode.Text+"''\r\n\r\nURČITĚ CHCEŠ  N E N Á V R A T N Ě  SMAZAT TUTO POZICI ?", "S M A Z Á N Í  POZICE!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                treeView1.SelectedNode.Remove();
            }
            treeView1.Focus();
        }



        public Boolean AddPart(String PartNumber)
        {
            return true;
        }


        private void button4_Click(object sender, EventArgs e)
        {
            Form3 form = new Form3();
            form.ShowDialog();
            if (form.DialogResult == DialogResult.OK)
            {
                String tagPos = treeView1.SelectedNode.Tag.ToString();
                String treeFullPath = treeView1.SelectedNode.FullPath.ToString();

                Boolean isPart = false;
                int partIndex = 0;
                for (int i = 0; i < parts.Count; i++)
                {
                    if ( parts[i].PartNumber == form.partNumber )
                    {
                        isPart = true;
                        partIndex = i;
                        break;
                    }
                }

                int addIndex = 0;
                if ( isPart )
                {
                    Console.Beep(500, 100);
                    Console.Beep(500, 100);
                    Console.Beep(500, 100);
                    Console.Beep(900, 300);
                    Console.Beep(300, 100);
                    DialogResult result = MessageBox.Show("TENTO  P A R T N U M B E R  UŽ SE NACHÁZÍ V DATABÁZI: \r\n=============================================\r\n\r\nPOZICE:>  " + parts[partIndex].TreeFullPath+ "\r\nPOČET:>  "+parts[partIndex].Count+ "\r\n\r\n=============================================\r\n\r\nANO = PŘIPSAT POČET TAM\r\n\r\nNE = PŘEPSAT VŠE SEM", "NALEZENA SHODA", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
                    switch( result )
                    {
                        case DialogResult.Yes:
                            {
                                break;
                            }
                        case DialogResult.No:
                            {
                                break;
                            }
                    }
                    treeView1.Focus();
                }
                else
                {
                    PART part = new PART();
                    part.PartNumber = form.partNumber;
                    part.Name = form.partName;
                    part.Count = form.partCount.ToString();
                    part.ParentTag = tagPos;
                    part.TreeFullPath = treeFullPath;
                    parts.Add(part);
                    addIndex = dataGridView1.Rows.Add();
                    dataGridView1.Rows[addIndex].Cells[0].Value = form.partNumber;
                    dataGridView1.Rows[addIndex].Cells[1].Value = form.partName;
                    dataGridView1.Rows[addIndex].Cells[2].Value = form.partCount.ToString();
                    dataGridView1.Rows[addIndex].Selected = true;
                    addIndex++;
                    treeView1.Focus();
                    Console.Beep(1300, 100);
                    Console.Beep(1500, 100);
                }

                String countParts = "";
                countParts = addIndex.ToString();
                countParts += " / ";
                countParts += parts.Count.ToString();
                countParts += "";
                label3.Text = countParts.ToString();

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form4 form = new Form4();
            form.ShowDialog();
        }

        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            String data = serialPort1.ReadLine();
        }
    }
}
