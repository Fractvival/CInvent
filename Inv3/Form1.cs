//using Microsoft.Office.Interop.Excel;
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
using System.Runtime.Serialization.Formatters.Binary;
using Excel = Microsoft.Office.Interop.Excel;

namespace Inv3
{
    public partial class Form1 : Form
    {
        public struct SETTING
        {
            public String NameRoot;
            public String ComName;
            public String ImportFolder;
            public String ExcelSearch;
            public String AutoRePath;
        }

        public struct PART
        {
            public String ParentTag;
            public String TreeFullPath;
            public String Name;
            public String PartNumber;
            public String KMZ;
            public String Count;
            public String OficialCount;

        }

        public struct KAT
        {
            public String KMZ;
            public String Text;
            public String Count;
        }

        public List<KAT> katList = new List<KAT>();
        public List<PART> partList = new List<PART>();
        public SETTING setting = new SETTING() { NameRoot="SKLAD", ComName="COM2", 
            ExcelSearch="*.xlsx", ImportFolder="import", AutoRePath="false" };
        String Katalog = "";

        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

        public static void SaveTree(TreeView tree)
        {
            String filename = System.Windows.Forms.Application.StartupPath.ToString() + @"\\tree.dat";
            if (File.Exists(filename))
                File.Delete(filename);
            using (Stream file = File.Open(filename, FileMode.Create))
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(file, tree.Nodes.Cast<TreeNode>().ToList());
            }
        }

        public static void LoadTree(TreeView tree)
        {
            String filename = System.Windows.Forms.Application.StartupPath.ToString() + @"\\tree.dat";
            if (File.Exists(filename))
            {
                using (Stream file = File.Open(filename, FileMode.Open))
                {
                    BinaryFormatter bf = new BinaryFormatter();
                    object obj = bf.Deserialize(file);

                    TreeNode[] nodeList = (obj as IEnumerable<TreeNode>).ToArray();
                    tree.Nodes.AddRange(nodeList);
                }
            }
        }


        public Form1()
        {
            InitializeComponent();

            Form4 form = new Form4();
            form.Show();

            treeView1.BeginUpdate();
            LoadTree(treeView1);

            if (treeView1.Nodes.Count == 0)
            {
                TreeNode addTree = new TreeNode { Text=setting.NameRoot, Tag="ROOT" };
                treeView1.Nodes.Add(addTree);
                treeView1.SelectedNode = addTree;
                treeView1.Focus();
            }

            treeView1.EndUpdate();
            treeView1.ExpandAll();
            
            serialPort1.Open();

            String pathname = System.Windows.Forms.Application.StartupPath.ToString();
            if (!Directory.Exists(pathname+@"\\"+setting.ImportFolder))
            {
                Directory.CreateDirectory(pathname+@"\\"+setting.ImportFolder);
            }
            pathname += "\\"+setting.ImportFolder;
            String[] names = Directory.GetFiles(pathname, setting.ExcelSearch);
            Katalog = names[0].ToString();

            Excel.Application xlApp;
            Excel.Workbook wb;
            Excel.Worksheet sheet;
            Excel.Range range;
            object misValue = System.Reflection.Missing.Value;
            xlApp = new Excel.Application();
            wb = xlApp.Workbooks.Open(Katalog, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            sheet = (Excel.Worksheet)wb.Worksheets.get_Item(1);
            int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
            range = sheet.Range["B6","C"+lastRow.ToString()];
            int rowCount = range.Rows.Count;
            int colCount = range.Columns.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                katList.Add(new KAT 
                {
                    KMZ = range.Cells[i, 1].Value2.Substring(0,10).ToString(),
                    Text = range.Cells[i,1].Value2.Remove(0,15).ToString(), 
                    Count = range.Cells[i, 2].Value2.ToString() 
                } ) ;
            }
            wb.Close(true, misValue, misValue);
            xlApp.Quit();
            form.Close();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }


        private void button1_Click(object sender, EventArgs e)
        {
            Form2 form = new Form2();
            TreeNode addTree = new TreeNode();
            form.ShowDialog();
            if (form.DialogResult == DialogResult.OK)
            {
                addTree.Text = form.newNodeText;
                addTree.Tag = form.newNodeText;
                treeView1.SelectedNode.Nodes.Add(addTree);
                treeView1.SelectedNode.Expand();
                treeView1.SelectedNode = addTree;
                treeView1.Focus();
                SaveTree(treeView1);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("POZICE:>   ''" + treeView1.SelectedNode.Text + "''\r\n\r\nURČITĚ CHCEŠ  N E N Á V R A T N Ě  SMAZAT TUTO POZICI ?", "S M A Z Á N Í  POZICE!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes)
            {
                treeView1.SelectedNode.Remove();
            }
            treeView1.Focus();
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
            countPositions += "/";
            countPositions += treeView1.SelectedNode.GetNodeCount(true).ToString();
            countPositions += "";
            label4.Text = countPositions.ToString();

            dataGridView1.Rows.Clear();

            int index = 0;
            int count = 0;
            for ( int i = 0; i < partList.Count; i++ )
            {
                if (treeView1.SelectedNode.FullPath.Equals(partList[i].TreeFullPath))
                {
                    index = dataGridView1.Rows.Add();
                    dataGridView1.Rows[index].Cells[0].Value = partList[i].PartNumber;
                    dataGridView1.Rows[index].Cells[1].Value = partList[i].Name;
                    dataGridView1.Rows[index].Cells[2].Value = partList[i].Count;
                    dataGridView1.Rows[index].Cells[3].Value = partList[i].OficialCount;
                    count++;
                    //dataGridView1.Rows[index].Selected = true;
                }
            }

            String countParts = "";
            countParts = count.ToString();
            countParts += "/";
            countParts += partList.Count.ToString();
            countParts += "";
            label3.Text = countParts.ToString();
            label5.Text = treeView1.SelectedNode.FullPath;
        }


        public int IsKMZinPartList( String KMZ )
        {
            for ( int i = 0; i < partList.Count; i++ )
            {
                if (partList[i].KMZ.Equals(KMZ))
                    return i;
            }
            return -1;
        }

        public int IsPartNumberinPartList(String PartNumber)
        {
            for (int i = 0; i < partList.Count; i++)
            {
                if (partList[i].PartNumber.Equals(PartNumber))
                    return i;
            }
            return -1;
        }

        public Boolean AddPart(PART item)
        {
            Boolean isTypeKMZ = false;
            Boolean isTypePartNumber = false;
            String treeSelectPath = treeView1.SelectedNode.FullPath;

            if (item.KMZ.Length > 0)
                isTypeKMZ = true;
            if (item.PartNumber.Length > 0)
                isTypePartNumber = true;

            if ( isTypePartNumber )
            {
                int indexPartList = IsPartNumberinPartList(item.PartNumber);
                if ( indexPartList != -1 )
                {
                    if ( partList[indexPartList].TreeFullPath.Equals(treeSelectPath) )
                    {
                        int iCount = Int32.Parse(partList[indexPartList].Count);
                        iCount++;
                        PART saveItem = new PART();
                        saveItem = partList[indexPartList];
                        saveItem.Count = iCount.ToString();
                        partList.RemoveAt(indexPartList);
                        partList.Add(saveItem);
                    }
                    else
                    { 
                    }
                }
                else
                {

                }
            }
            else if ( isTypeKMZ )
            {

            }
            else
            {

            }

            return true;
        }

        // PRIDAT NOVY DIL
        private void button4_Click(object sender, EventArgs e)
        {
            Form3 form = new Form3();
            form.ShowDialog();
            if (form.DialogResult == DialogResult.OK)
            {
                AddPart(new PART { 
                    KMZ = form.KMZ, 
                    PartNumber = form.partNumber, 
                    Name = form.partName, 
                    });
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
        }

        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            String data = serialPort1.ReadLine();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveTree(treeView1);
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].IsNewRow)
            {
                MessageBox.Show("NEMA DATA");
            }
            else 
            {
                MessageBox.Show("MA DATA");
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void xLSToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
    }
}
