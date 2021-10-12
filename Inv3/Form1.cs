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
using System.Management;
using System.Runtime.Serialization.Formatters.Binary;
using Excel = Microsoft.Office.Interop.Excel;

namespace Inv3
{
    public partial class Form1 : Form
    {
        public struct SETTING
        {
            // Nazev rootu v seznamu pozic (bude vytvoren v nove databazi)
            public String NameRoot;
            // Nazev COM portu skeneru (COM1,COM2,...)
            public String ComName;
            // Nazev slozky obsahujici ofiko sklad
            public String ImportFolder;
            // Koncovka pro hledani dokumentu (tradicne ".xlsx")
            public String ExcelSearch;
            // Automaticke ukladani poctu existujici polozky do sve pozice
            // Pokud tedy pridavany nebo naskenovany dil uz v databazi existuje, ale jeho pozice
            // zrovna neni otevrena, ma se automaticky pripsat novy pocet do jeho pozice
            // true = ano, false = bude se zobrazovat dialog pro zvoleni akce
            public String AutoRePath;
            // Pokud jeste dil v databazi neexistuje, bude prohledan ofiko sklad a pokud se v ofiko
            // skladu nachazi (na zaklade KMZ nebo PARTNUMBER), budou mu pripsany tyto informace 
            // z ofiko skladu
            // true = ano, false = bude se zobrazovat potvrzovaci dialog
            public String AutoAddInfo;
        }

        public struct PART
        {
            public PART(String KMZ = "", String PartNumber = "",
                String Name = "", String Count = "", String OficialCount = "", String ParentTag = "", String TreeFullPath = "")
            {
                this.KMZ = KMZ;
                this.PartNumber = PartNumber;
                this.Name = Name;
                this.Count = Count;
                this.OficialCount = OficialCount;
                this.ParentTag = ParentTag;
                this.TreeFullPath = TreeFullPath;
            }
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

        // Tento seznam je nactem z oficialniho XLS dokumentu o skladu
        public static List<KAT> katList = new List<KAT>();
        // Tento seznam obsahuje vsechny bud rucne pridane nebo naskenovane dily
        public static List<PART> partList = new List<PART>();
        // Zde je obsazeno nastaveni programu
        public static SETTING setting = new SETTING()
        {
            NameRoot = "SKLAD",
            ComName = "COM2",
            ExcelSearch = "*.xlsx",
            ImportFolder = "import",
            AutoRePath = "false",
            AutoAddInfo = "true"
        };
        // Toto je pomocna promenna s plnou cestou a nazvem souboru ofiko skladu (doplneno dale takze nemenit)
        String Katalog = "";

        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();


        public static void SavePart()
        {
            String filename = System.Windows.Forms.Application.StartupPath.ToString() + @"\\part.dat";
            try
            {
                if (File.Exists(filename))
                    File.Delete(filename);
                StreamWriter writer = new StreamWriter(filename);
                for ( int i = 0; i < partList.Count; i++ )
                {
                    writer.WriteLine(partList[i].KMZ);
                    writer.WriteLine(partList[i].PartNumber);
                    writer.WriteLine(partList[i].Name);
                    writer.WriteLine(partList[i].Count);
                    writer.WriteLine(partList[i].OficialCount);
                    writer.WriteLine(partList[i].TreeFullPath);
                    writer.WriteLine(partList[i].ParentTag);
                }
                writer.Close();
            }
            catch(IOException ioEx)
            {
                MessageBox.Show("NASTALY PROBLÉMY PŘI UKLÁDÁNÍ SEZNAMU DÍLŮ !!\r\n\r\n" +
                    ioEx.Message + "\r\n\r\n" +
                    "ZKONTROLUJ ZDA SE exe soubor PROGRAMU NACHÁZÍ\r\n" +
                    "NA MÍSTĚ S POVOLENÝM ZÁPISEM NA DISK", "PROBLÉM V UKLÁDÁNÍ SEZNAMU DÍLŮ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public static void LoadPart()
        {
            String filename = System.Windows.Forms.Application.StartupPath.ToString() + @"\\part.dat";
            if (File.Exists(filename))
            {
                try
                {
                    StreamReader reader = new StreamReader(filename);
                    while ( !reader.EndOfStream )
                    {
                        partList.Add(new PART
                        {
                            KMZ = reader.ReadLine(),
                            PartNumber = reader.ReadLine(),
                            Name = reader.ReadLine(),
                            Count = reader.ReadLine(),
                            OficialCount = reader.ReadLine(),
                            TreeFullPath = reader.ReadLine(),
                            ParentTag = reader.ReadLine()
                        }); ;
                    }
                    reader.Close();
                }
                catch 
                {
                }
            }
        }


        public static void SaveTree(TreeView tree)
        {
            String filename = System.Windows.Forms.Application.StartupPath.ToString() + @"\\tree.dat";
            try
            {
                if (File.Exists(filename))
                    File.Delete(filename);
                using (Stream file = File.Open(filename, FileMode.Create))
                {
                    BinaryFormatter bf = new BinaryFormatter();
                    bf.Serialize(file, tree.Nodes.Cast<TreeNode>().ToList());
                }
            }
            catch(IOException ioEx)
            {
                MessageBox.Show("NASTALY PROBLÉMY PŘI UKLÁDÁNÍ SEZNAMU POZIC\r\n\r\n" +
                    ioEx.Message + "\r\n\r\n" +
                    "ZKONTROLUJ ZDA SE exe soubor PROGRAMU NACHÁZÍ\r\n" +
                    "NA MÍSTĚ S POVOLENÝM ZÁPISEM NA DISK", "PROBLÉM V UKLÁDÁNÍ NASTAVENÍ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public static void LoadTree(TreeView tree)
        {
            String filename = System.Windows.Forms.Application.StartupPath.ToString() + @"\\tree.dat";
            try
            {
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
            catch(IOException ioEx)
            {

            }
        }

        public static void SaveSetting()
        {
            String filename = System.Windows.Forms.Application.StartupPath.ToString() + @"\\setting.dat";
            try
            {
                if (File.Exists(filename))
                    File.Delete(filename);
                StreamWriter writer = new StreamWriter(filename, false);
                writer.WriteLine(setting.NameRoot);
                writer.WriteLine(setting.ImportFolder);
                writer.WriteLine(setting.ExcelSearch);
                writer.WriteLine(setting.ComName);
                writer.WriteLine(setting.AutoRePath);
                writer.WriteLine(setting.AutoAddInfo);
                writer.Close();
            }
            catch (IOException ioEx)
            {
                MessageBox.Show("NASTALY PROBLÉMY PŘI UKLÁDÁNÍ NASTAVENÍ PROGRAMU\r\n\r\n" +
                    ioEx.Message + "\r\n\r\n" +
                    "ZKONTROLUJ ZDA SE exe soubor PROGRAMU NACHÁZÍ\r\n" +
                    "NA MÍSTĚ S POVOLENÝM ZÁPISEM NA DISK", "PROBLÉM V UKLÁDÁNÍ NASTAVENÍ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        public static void LoadSetting()
        {
            String filename = System.Windows.Forms.Application.StartupPath.ToString() + @"\\setting.dat";
            try
            {
                StreamReader reader = new StreamReader(filename);
                setting.NameRoot = reader.ReadLine();
                setting.ImportFolder = reader.ReadLine();
                setting.ExcelSearch = reader.ReadLine();
                setting.ComName = reader.ReadLine();
                setting.AutoRePath = reader.ReadLine();
                setting.AutoAddInfo = reader.ReadLine();
                reader.Close();
            }
            catch (IOException ioEx)
            {
                MessageBox.Show("NASTALY PROBLÉMY PŘI NAČÍTÁNÍ NASTAVENÍ PROGRAMU\r\n\r\n" +
                    ioEx.Message + "\r\n\r\n" +
                    "ZKONTROLUJ ZDA SE exe soubor PROGRAMU NACHÁZÍ\r\n" +
                    "NA MÍSTĚ S POVOLENÝM ZÁPISEM NA DISK", "PROBLÉM V NAČÍTÁNÍ NASTAVENÍ", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void MenuSettingPort_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            Boolean isErr = false;
            try
            {
                if (serialPort1.IsOpen)
                    serialPort1.Close();
                serialPort1.PortName = e.ClickedItem.Text;
                serialPort1.Open();
                //serialPort1.Close();
            }
            catch(IOException io)
            {
                isErr = true;
            }
            if ( isErr )
            {
                toolStripStatusLabel1.Text = ">> CHYBA PORTU !";
                Console.Beep(500, 100);
                Console.Beep(500, 100);
            }
            else
            {
                toolStripStatusLabel1.Text = ">> SKENER NA "+serialPort1.PortName;
                setting.ComName = serialPort1.PortName;
                Console.Beep(1300, 80);
                Console.Beep(1500, 80);
                Console.Beep(1300, 80);
            }
        }

        public void InitPort()
        {
            MenuSettingPort.DropDownItems.Clear();
            Boolean isErr = false;
            for ( int i = 1; i < 25; i++ )
            {
                isErr = false;
                try
                {
                    serialPort1.PortName = "COM" + i.ToString();
                    serialPort1.Open();
                    serialPort1.Close();
                }
                catch(IOException ioEx)
                {
                    isErr = true;
                }
                if (!isErr)
                    MenuSettingPort.DropDownItems.Add(serialPort1.PortName.ToString());
            }
            isErr = false;
            try
            {
                if (serialPort1.IsOpen)
                    serialPort1.Close();
                serialPort1.PortName = setting.ComName;
                serialPort1.Open();
            }
            catch (IOException ioEx)
            {
                isErr = true;
                toolStripStatusLabel1.Text = ">> CHYBA PORTU !" + serialPort1.PortName;
                MessageBox.Show("NELZE OTEVŘÍT COM PORT SKENERU\r\n\r\n" +
                    ioEx.Message + "\r\n\r\n" +
                    "V NASTAVENÍ PROGRAMU ZVOL Z DOSTUPNÝCH COM PORTŮ\r\n" +
                    "", "PROBLÉM COM PORTU SKENERU", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            if (!isErr)
                toolStripStatusLabel1.Text = ">> SKENER NA " + serialPort1.PortName;
        }


        public Form1()
        {
            InitializeComponent();

            LoadSetting();
            InitPort();
            LoadPart();


            treeView1.BeginUpdate();
            LoadTree(treeView1);
            if (treeView1.Nodes.Count == 0)
            {
                TreeNode addTree = new TreeNode { Text = setting.NameRoot, Tag = "ROOT" };
                treeView1.Nodes.Add(addTree);
                treeView1.SelectedNode = addTree;
                treeView1.Focus();
            }
            treeView1.EndUpdate();
            treeView1.ExpandAll();

            String pathname = System.Windows.Forms.Application.StartupPath.ToString();
            if (!Directory.Exists(pathname + @"\\" + setting.ImportFolder))
            {
                try
                {
                    Directory.CreateDirectory(pathname + @"\\" + setting.ImportFolder);
                }
                catch(IOException ioEx)
                {
                    MessageBox.Show("NASTAL PROBLÉM INICIALIZACE IMPORTNÍ SLOŽKY\r\n\r\n" +
                        ioEx.Message + "\r\n\r\n" +
                        "ZKONTROLUJ ZDA SE exe soubor PROGRAMU\r\n"+
                        "NACHÁZÍ NA MÍSTĚ S POVOLENÝM ZÁPISEM NA DISK", "PROBLÉM SLOŽKY PRO IMPORT", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            pathname += "\\" + setting.ImportFolder;
            DirectoryInfo info = new DirectoryInfo(pathname);
            try
            {
                FileInfo[] files = info.GetFiles(setting.ExcelSearch).OrderBy(p => p.CreationTime).ToArray();
                Katalog = files.Last().FullName.ToString(); // names[0].ToString();
            }
            catch(IOException ioEx)
            {
                MessageBox.Show("NASTAL PROBLÉM PŘI HLEDÁNÍ KATALOGU\r\n\r\n"+
                    ioEx.Message+"\r\n\r\n"+
                    "ZKONTROLUJ ZDA JE OBSAŽEN excel SOUBOR VE SLOŽCE import", "PROBLÉM KATALOGU", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            Form4 form = new Form4();
            form.Show();
            try
            {
                Excel.Application xlApp;
                Excel.Workbook wb;
                Excel.Worksheet sheet;
                Excel.Range range;
                object misValue = System.Reflection.Missing.Value;
                xlApp = new Excel.Application();
                wb = xlApp.Workbooks.Open(Katalog, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                sheet = (Excel.Worksheet)wb.Worksheets.get_Item(1);
                int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                range = sheet.Range["B6", "C" + lastRow.ToString()];
                int rowCount = range.Rows.Count;
                int colCount = range.Columns.Count;
                for (int i = 1; i <= rowCount; i++)
                {
                    katList.Add(new KAT
                    {
                        KMZ = range.Cells[i, 1].Value2.Substring(0, 10).ToString(),
                        Text = range.Cells[i, 1].Value2.Remove(0, 15).ToString(),
                        Count = range.Cells[i, 2].Value2.ToString()
                    });
                }
                wb.Close(true, misValue, misValue);
                xlApp.Quit();
            }
            catch(IOException ioEx)
            {
                MessageBox.Show("NASTAL PROBLÉM PŘI NAČÍTÁNÍ Z KATALOGU\r\n\r\n" +
                    ioEx.Message + "\r\n\r\n" +
                    "ZKONTROLUJ ZDA JE excel SOUBOR V ŘÁDNÉM FORMÁTU", "PROBLÉM KATALOGU", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
            if (treeView1.SelectedNode.Tag.ToString() == "ROOT")
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
            //int index = 0;
            int count = 0;
            for (int i = 0; i < partList.Count; i++)
            {
                if (treeView1.SelectedNode.FullPath.Equals(partList[i].TreeFullPath))
                {
                    ChangeItemGrid(new PART
                    {
                        KMZ = partList[i].KMZ,
                        PartNumber = partList[i].PartNumber,
                        Name = partList[i].Name,
                        Count = partList[i].Count,
                        OficialCount = partList[i].OficialCount,
                        TreeFullPath = partList[i].TreeFullPath
                    });
                    count++;
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

        // Prohleda databazi zadanych/nactenych dilu na existenci KMZ
        // Pokud existuje, vrati jeho index v seznamu, jinak -1
        public int IsKMZinPartList(String KMZ)
        {
            for (int i = 0; i < partList.Count; i++)
            {
                try
                {
                    if (partList[i].KMZ.Equals(KMZ))
                        return i;
                }
                catch
                {
                    return -1;
                }
            }
            return -1;
        }

        // Prohleda databazi zadanych/nactenych dilu na existenci PN
        // Pokud existuje, vrati jeho index v seznamu, jinak -1
        public int IsPartNumberinPartList(String PartNumber)
        {
            for (int i = 0; i < partList.Count; i++)
            {
                try
                {
                    if (partList[i].PartNumber.Equals(PartNumber))
                        return i;
                }
                catch
                {
                    return -1;
                }
            }
            return -1;
        }

        public int IsKMZinKatalogList(String KMZ)
        {
            for (int i = 0; i < katList.Count; i++)
            {
                try
                {
                    if (katList[i].KMZ.Equals(KMZ))
                        return i;
                }
                catch
                {
                    return -1;
                }
            }
            return -1;
        }

        public int IsPartNumberinKatalogList(String PartNumber)
        {
            for (int i = 0; i < katList.Count; i++)
            {
                try
                {
                    if (katList[i].Text.Contains(PartNumber))
                        return i;
                }
                catch
                {
                    return -1;
                }
            }
            return -1;
        }


        public Boolean ChangeItemGrid(PART item)
        {
            int countRow = dataGridView1.Rows.Count;
            int indexItem = -1;

            if (item.KMZ == null)
                item.KMZ = "";
            if (item.PartNumber == null)
                item.PartNumber = "";
            if (item.Name == null)
                item.Name = "";
            if (item.Count == null)
                item.Count = "";
            if (item.OficialCount == null)
                item.OficialCount = "";
            if (item.TreeFullPath == null)
                item.TreeFullPath = "";
            if (item.ParentTag == null)
                item.ParentTag = "";

            for (int i = 0; i < countRow; i++)
            {
                if (!dataGridView1.Rows[i].IsNewRow)
                {
                    if ((dataGridView1.Rows[i].Cells[0].Value.ToString().Equals(item.KMZ.ToString())) &&
                        (dataGridView1.Rows[i].Cells[1].Value.ToString().Equals(item.PartNumber.ToString())))
                    {
                        indexItem = i;
                        break;
                    }
                }
            }
            if (indexItem != -1)
            {
                dataGridView1.Rows[indexItem].Cells[0].Value = item.KMZ;
                dataGridView1.Rows[indexItem].Cells[1].Value = item.PartNumber;
                dataGridView1.Rows[indexItem].Cells[2].Value = item.Name;
                dataGridView1.Rows[indexItem].Cells[3].Value = item.Count;
                dataGridView1.Rows[indexItem].Cells[4].Value = item.OficialCount;
                dataGridView1.Rows[indexItem].Cells[5].Value = item.TreeFullPath;
                if (Int32.Parse(item.Count) <= 0)
                    dataGridView1.Rows[indexItem].Cells[3].Style.BackColor = Color.FromArgb(255, 180, 180);
                if (Int32.Parse(item.Count) > 0)
                    dataGridView1.Rows[indexItem].Cells[3].Style.BackColor = Color.FromArgb(180, 255, 180);
                if (Int32.Parse(item.Count) < Int32.Parse(item.OficialCount))
                    dataGridView1.Rows[indexItem].Cells[4].Style.BackColor = Color.FromArgb(255, 100, 100);
                dataGridView1.Rows[indexItem].Selected = true;
            }
            else
            {
                //DataGridViewRow nRow = new DataGridViewRow();
                int newIndexRow = 0;

                dataGridView1.Invoke((MethodInvoker)delegate {
                    newIndexRow = dataGridView1.Rows.Add();
                });

                //int newIndexRow = dataGridView1.Rows.Add();
                dataGridView1.Rows[newIndexRow].Cells[0].Value = item.KMZ;
                dataGridView1.Rows[newIndexRow].Cells[1].Value = item.PartNumber;
                dataGridView1.Rows[newIndexRow].Cells[2].Value = item.Name;
                dataGridView1.Rows[newIndexRow].Cells[3].Value = item.Count;
                dataGridView1.Rows[newIndexRow].Cells[4].Value = item.OficialCount;
                dataGridView1.Rows[newIndexRow].Cells[5].Value = item.TreeFullPath;
                if (Int32.Parse(item.Count) <= 0)
                    dataGridView1.Rows[newIndexRow].Cells[3].Style.BackColor = Color.FromArgb(255, 180, 180);
                if (Int32.Parse(item.Count) > 0)
                    dataGridView1.Rows[newIndexRow].Cells[3].Style.BackColor = Color.FromArgb(180, 255, 180);
                if (Int32.Parse(item.Count) < Int32.Parse(item.OficialCount))
                    dataGridView1.Rows[newIndexRow].Cells[4].Style.BackColor = Color.FromArgb(255, 100, 100);
                dataGridView1.Rows[newIndexRow].Selected = true;
            }
            return true;
        }


        // Vlozi dil do seznamu v aktualne vybrane pozici
        public Boolean AddPart(PART item, int minusplusCount)
        {
            if (item.KMZ == null)
                item.KMZ = "";
            if (item.PartNumber == null)
                item.PartNumber = "";
            if (item.Name == null)
                item.Name = "";
            if (item.Count == null)
                item.Count = "";
            if (item.OficialCount == null)
                item.OficialCount = "";
            if (item.TreeFullPath == null)
                item.TreeFullPath = "";
            if (item.ParentTag == null)
                item.ParentTag = "";

            Boolean isTypeKMZ = false;
            Boolean isTypePartNumber = false;
            String treeSelectPath = @"";

            treeView1.Invoke((MethodInvoker)delegate {
                treeSelectPath = treeView1.SelectedNode.FullPath;
            });

            //String treeSelectPath = treeView1.SelectedNode.FullPath.ToString();

            // Zjistim, jestli bylo zadane i KMZ
            if (item.KMZ.Length > 0)
                isTypeKMZ = true;
            // Zjistim, jestli byl zadan PARTNUMBER
            if (item.PartNumber.Length > 0)
                isTypePartNumber = true;

            // Pokud byl zadan PARTNUMBER, dam mu prednost v pridavani pred KMZ
            // KMZ beru jako vedlejsi informaci kterou vyuziji pouze pokud neni PN
            if (isTypePartNumber)
            {
                // Zjistim, zda-li se uz dil nahodou nenachazi v databazi dilu
                int indexPartList = IsPartNumberinPartList(item.PartNumber);
                // Pokud fce indexPartList NEvraci -1, znamena to ze dil v seznamu dilu uz je
                // a vysledkem teto fce je pozice (jeho index) v seznamu dilu
                if (indexPartList != -1)
                {
                    // Nyni otestuji, zda-li je tento dil obsazen v aktualne otevrene pozici
                    if (partList[indexPartList].TreeFullPath.Equals(treeSelectPath))
                    {
                        // Takze dil je v seznamu a dokonce i v otevrene pozici
                        // Nyni zjistim jeho pocet a pote upravim na zaklade parametru minusplusCount
                        int iCount = Int32.Parse(partList[indexPartList].Count);
                        if (minusplusCount > 0)
                            iCount += minusplusCount;
                        if (minusplusCount < 0)
                            iCount -= minusplusCount;
                        PART saveItem = new PART();
                        saveItem = partList[indexPartList];
                        saveItem.Count = iCount.ToString();
                        // Protoze v kolekci List<> nelze upravovat polozku primo, tak ji nejdrive odstranim
                        // a pote znovu pridam s upravenym poctem
                        partList.RemoveAt(indexPartList);
                        partList.Add(saveItem);
                        ChangeItemGrid(saveItem);
                        Console.Beep(1500, 80);
                    }
                    // Pokud dil NENI v aktualne otevrene pozici, a protoze uz je v seznamu (mame jeho index),
                    // je nutne ho pripsat na spravne misto. Toto lze provest bud automaticky (dle nastaveni)
                    // nebo pokud neni automatika, zeptame se uzivatele skrzeva dialog jak to chce.
                    // Uzivatel bude moct rozhodnout jestli jeho pocet pripsat na misto kde se nachazi, nebo
                    // jestli ho chce z jeho pozice nejdrive presunout do otevrene pozice a tam mu i 
                    // pripsat pocet.
                    else
                    {
                        if (setting.AutoRePath.ToUpper().Equals("TRUE"))
                        {
                            int iCount = Int32.Parse(partList[indexPartList].Count);
                            if (minusplusCount > 0)
                                iCount += minusplusCount;
                            if (minusplusCount < 0)
                                iCount -= minusplusCount;
                            PART saveItem = new PART();
                            saveItem = partList[indexPartList];
                            saveItem.Count = iCount.ToString();
                            // Protoze v kolekci List<> nelze upravovat polozku primo, tak ji nejdrive odstranim
                            // a pote znovu pridam s upravenym poctem
                            partList.RemoveAt(indexPartList);
                            partList.Add(saveItem);
                            Console.Beep(1500, 80);
                            Console.Beep(1500, 80);
                        }
                        else
                        {
                            DialogResult result = MessageBox.Show("TATO POLOZKA UZ V DATABAZI EXISTUJE !\r\n\r\nJEJI POZICE:> " + partList[indexPartList].TreeFullPath +
                                "\r\n\r\nKLIKNI NA ''A N O'' POKUD JI MAM ZAPSAT NA SVE MISTO\r\n" +
                                "KLIKNI NA ''N E'' POKUD JI MAM PREPSAT SEM\r\n\r\n" +
                                "KLIKNI NA ''S T O R N O'' PRO ZADNOU AKCI", "EXISTUJICI POLOZKA", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
                            switch (result)
                            {
                                case DialogResult.Yes:
                                    {
                                        int iCount = Int32.Parse(partList[indexPartList].Count);
                                        if (minusplusCount > 0)
                                            iCount += minusplusCount;
                                        if (minusplusCount < 0)
                                            iCount -= minusplusCount;
                                        PART saveItem = new PART();
                                        saveItem = partList[indexPartList];
                                        saveItem.Count = iCount.ToString();
                                        // Protoze v kolekci List<> nelze upravovat polozku primo, tak ji nejdrive odstranim
                                        // a pote znovu pridam s upravenym poctem
                                        partList.RemoveAt(indexPartList);
                                        partList.Add(saveItem);
                                        Console.Beep(1500, 80);
                                        break;
                                    }
                                case DialogResult.No:
                                    {
                                        int iCount = Int32.Parse(partList[indexPartList].Count);
                                        if (minusplusCount > 0)
                                            iCount += minusplusCount;
                                        if (minusplusCount < 0)
                                            iCount -= minusplusCount;
                                        PART saveItem = new PART();
                                        saveItem = partList[indexPartList];
                                        saveItem.TreeFullPath = treeSelectPath;
                                        saveItem.Count = iCount.ToString();
                                        // Protoze v kolekci List<> nelze upravovat polozku primo, tak ji nejdrive odstranim
                                        // a pote znovu pridam s upravenym poctem
                                        partList.RemoveAt(indexPartList);
                                        partList.Add(saveItem);
                                        ChangeItemGrid(saveItem);
                                        Console.Beep(1500, 80);
                                        Console.Beep(1500, 80);
                                        break;
                                    }
                                case DialogResult.Cancel:
                                    {
                                        break;
                                    }
                            }
                        }
                    }
                }
                // POKUD se dil v seznamu dilu jeste nenachazi, zavedeme jej tedy jako novy do 
                // aktualne zvolene pozice
                else
                {
                    // V zakladu muzeme rovnou pro novy dil pripsat jeho PN jelikoz
                    // se nachazime ve splnene podmince pro vyplnene PN z Form
                    // Zrovna tak uz muzeme pridat cestu aktualne zvolene pozice z Tree
                    PART katItem = new PART() { PartNumber = item.PartNumber, TreeFullPath = treeSelectPath };
                    // Zkusime, jestli se na zaklade zadaneho PN dil nachazi v katalogu
                    int indexKat = IsPartNumberinKatalogList(item.PartNumber);
                    // Pokud nenachazi...
                    if (indexKat == -1)
                    {
                        // ..a pokud bylo zadano alespon KMZ..
                        if (isTypeKMZ)
                        {
                            // ..tak muzeme uz zapsat i KMZ..
                            katItem.KMZ = item.KMZ;
                            // ..a jeste se pokusit prohledat katalog na shodne KMZ
                            indexKat = IsKMZinKatalogList(item.KMZ);
                        }
                    }
                    // Pokud se v katalogu naslo PN nebo aspon KMZ, doplnenime ostatni informace
                    // pro novy dil z katalogu
                    if (indexKat != -1)
                    {
                        katItem.KMZ = katList[indexKat].KMZ;
                        katItem.OficialCount = katList[indexKat].Count;
                        katItem.Count = "0";
                        katItem.Name = katList[indexKat].Text;
                    }
                    // pokud vsak novy dil neexistuje ani v ofiko katalogu, zapiseme jeho pocty na nuly
                    else
                    {
                        katItem.OficialCount = "0";
                        katItem.Count = "0";
                    }
                    // AT uz hledani v katalogu dopadlo jakkoliv, musime jeste upravit nove dany pocet
                    int iCount = Int32.Parse(katItem.Count); ;
                    if (minusplusCount > 0)
                        iCount += minusplusCount;
                    if (minusplusCount < 0)
                        iCount -= minusplusCount;
                    katItem.Count = iCount.ToString();
                    partList.Add(katItem);
                    ChangeItemGrid(katItem);
                    Console.Beep(1500, 80);
                }
            }
            // TAKZE pokud neni PN a je zadany KMZ, budu pridavat na zaklade KMZ
            else if (isTypeKMZ)
            {
                // V zakladu muzeme rovnou pro novy dil pripsat jeho KMZ jelikoz
                // se nachazime ve splnene podmince pro vyplnene KMZ z Form, avsak mimo PN
                // Zrovna tak uz muzeme pridat cestu aktualne zvolene pozice z Tree
                PART katItem = new PART() { KMZ = item.KMZ, TreeFullPath = treeSelectPath };
                // Zkusime, jestli se na zaklade zadaneho KMZ dil nachazi v katalogu
                int indexKat = IsKMZinKatalogList(item.KMZ);
                // Pokud se v katalogu naslo KMZ, doplnenime ostatni informace
                // pro novy dil z katalogu
                if (indexKat != -1)
                {
                    katItem.OficialCount = katList[indexKat].Count;
                    katItem.Count = "0";
                    katItem.Name = katList[indexKat].Text;
                }
                // pokud vsak novy dil neexistuje ani v ofiko katalogu, zapiseme jeho pocty na nuly
                else
                {
                    katItem.OficialCount = "0";
                    katItem.Count = "0";
                }
                // AT uz hledani v katalogu dopadlo jakkoliv, musime jeste upravit nove dany pocet
                int iCount = Int32.Parse(katItem.Count); ;
                if (minusplusCount > 0)
                    iCount += minusplusCount;
                if (minusplusCount < 0)
                    iCount -= minusplusCount;
                katItem.Count = iCount.ToString();
                partList.Add(katItem);
                ChangeItemGrid(katItem);
                Console.Beep(1500, 80);
            }
            // Neni nic ?
            // To je problem, a polozka nemuze byt vlozena do seznamu
            else
            {
                Console.Beep(500, 100);
                Console.Beep(500, 100);
                return false;
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
                AddPart(new PART
                {
                    KMZ = form.KMZ,
                    PartNumber = form.partNumber,
                    Name = form.partName,
                    Count = "",
                    OficialCount = "",
                    TreeFullPath = "",
                    ParentTag = ""
                }, form.partCount);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
        }

        private void serialPort1_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            String data = serialPort1.ReadLine();
            AddPart(new PART
            {
                KMZ = "",
                PartNumber = data.ToString(),
                Name = "",
                Count = "",
                OficialCount = "",
                TreeFullPath = "",
                ParentTag = ""
            }, 1); ;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveTree(treeView1);
            SavePart();
            SaveSetting();
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
            if (dataGridView1.SelectedRows.Count != 0)
            {
                if (!this.dataGridView1.SelectedRows[0].IsNewRow)
                {
                    button5.Enabled = true;
                    button6.Enabled = true;
                }
                else
                {
                    button5.Enabled = false;
                    button6.Enabled = false;
                }
            }
            else
            {
                button5.Enabled = false;
                button6.Enabled = false;
            }
        }

        private void xLSToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 0)
            {
                if (!this.dataGridView1.SelectedRows[0].IsNewRow)
                {
                    DialogResult result = MessageBox.Show("ODSTRANIT TUTO POLOŽKU ?\r\n\r\n" +
                        this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString() + "\r\n" +
                        this.dataGridView1.SelectedRows[0].Cells[1].Value.ToString() + "\r\n" +
                        this.dataGridView1.SelectedRows[0].Cells[2].Value.ToString() + "\r\n\r\n", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        Boolean isOK = false;
                        if (this.dataGridView1.SelectedRows[0].Cells[1].Value.ToString().Length > 0)
                        {
                            for (int i = 0; i < partList.Count; i++)
                            {
                                if (partList[i].PartNumber.Equals(this.dataGridView1.SelectedRows[0].Cells[1].Value.ToString()))
                                {
                                    try
                                    {
                                        partList.RemoveAt(i);
                                        isOK = true;
                                        break;
                                    }
                                    catch
                                    {
                                        isOK = false;
                                        break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString().Length > 0)
                            {
                                for (int i = 0; i < partList.Count; i++)
                                {
                                    if (partList[i].KMZ.Equals(this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString()))
                                    {
                                        try
                                        {
                                            partList.RemoveAt(i);
                                            isOK = true;
                                            break;
                                        }
                                        catch
                                        {
                                            isOK = false;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        if (isOK)
                        {
                            this.dataGridView1.Rows.Remove(this.dataGridView1.SelectedRows[0]);
                        }
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count != 0)
            {
                if (!this.dataGridView1.SelectedRows[0].IsNewRow)
                {
                    int indexPart = -1;
                    if (this.dataGridView1.SelectedRows[0].Cells[1].Value.ToString().Length > 0)
                    {
                        for (int i = 0; i < partList.Count; i++)
                        {
                            if (partList[i].PartNumber.Equals(this.dataGridView1.SelectedRows[0].Cells[1].Value.ToString()))
                            {
                                try
                                {
                                    indexPart = i;
                                    break;
                                }
                                catch
                                {
                                    break;
                                }
                            }
                        }
                    }
                    else
                    {
                        if (this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString().Length > 0)
                        {
                            for (int i = 0; i < partList.Count; i++)
                            {
                                if (partList[i].KMZ.Equals(this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString()))
                                {
                                    try
                                    {
                                        indexPart = i;
                                        break;
                                    }
                                    catch
                                    {
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    if (indexPart != -1)
                    {
                        Form3 form = new Form3();
                        form.partCount = Int32.Parse(partList[indexPart].Count.ToString());
                        form.partName = partList[indexPart].Name.ToString();
                        form.partNumber = partList[indexPart].PartNumber.ToString();
                        form.KMZ = partList[indexPart].KMZ.ToString();
                        DialogResult result = form.ShowDialog();
                        if (result == DialogResult.OK)
                        {
                            PART saveItem = new PART();
                            saveItem = partList[indexPart];
                            saveItem.KMZ = form.KMZ;
                            saveItem.PartNumber = form.partNumber;
                            saveItem.Name = form.partName;
                            saveItem.Count = form.partCount.ToString();
                            this.dataGridView1.Rows.Remove(this.dataGridView1.SelectedRows[0]);
                            partList.RemoveAt(indexPart);
                            partList.Add(saveItem);
                            ChangeItemGrid(saveItem);
                        }
                    }
                }
            }
        }

    }
}
