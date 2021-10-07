
namespace Inv3
{
    partial class Form1
    {
        /// <summary>
        /// Vyžaduje se proměnná návrháře.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Uvolněte všechny používané prostředky.
        /// </summary>
        /// <param name="disposing">hodnota true, když by se měl spravovaný prostředek odstranit; jinak false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Kód generovaný Návrhářem Windows Form

        /// <summary>
        /// Metoda vyžadovaná pro podporu Návrháře - neupravovat
        /// obsah této metody v editoru kódu.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.PartNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PartName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Count = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CountKat = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.serialPort1 = new System.IO.Ports.SerialPort(this.components);
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.nASTAVENIToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pORTSKENERUToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM1ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM2ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM3ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM4ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM5ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM6ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM7ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM8ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM9ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM10ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM11ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM12ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM13ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM14ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cOM15ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.kATALOGToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.xLSXToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.xLSToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.xLSToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.zVOLITEXTRAToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.treeView1.Font = new System.Drawing.Font("Consolas", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.treeView1.FullRowSelect = true;
            this.treeView1.HotTracking = true;
            this.treeView1.Location = new System.Drawing.Point(13, 27);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(365, 420);
            this.treeView1.TabIndex = 0;
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.groupBox1.Location = new System.Drawing.Point(385, 27);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(543, 100);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "INFORMACE";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Consolas", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label4.Location = new System.Drawing.Point(192, 61);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(20, 22);
            this.label4.TabIndex = 3;
            this.label4.Text = "0";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Consolas", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label3.Location = new System.Drawing.Point(192, 28);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(20, 22);
            this.label3.TabIndex = 2;
            this.label3.Text = "0";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Consolas", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label2.Location = new System.Drawing.Point(26, 61);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(140, 22);
            this.label2.TabIndex = 1;
            this.label2.Text = "CELKEM POZIC:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Consolas", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label1.Location = new System.Drawing.Point(26, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(160, 22);
            this.label1.TabIndex = 0;
            this.label1.Text = "CELKEM POLOŽEK:";
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button1.Font = new System.Drawing.Font("Consolas", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.button1.Location = new System.Drawing.Point(13, 453);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(167, 46);
            this.button1.TabIndex = 3;
            this.button1.Text = "PŘIDAT";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button2.Font = new System.Drawing.Font("Consolas", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.button2.Location = new System.Drawing.Point(257, 453);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(121, 46);
            this.button2.TabIndex = 4;
            this.button2.Text = "SMAZAT";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView1.ClipboardCopyMode = System.Windows.Forms.DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.PartNumber,
            this.PartName,
            this.Count,
            this.CountKat});
            this.dataGridView1.Location = new System.Drawing.Point(385, 191);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridView1.Size = new System.Drawing.Size(543, 256);
            this.dataGridView1.TabIndex = 5;
            this.dataGridView1.SelectionChanged += new System.EventHandler(this.dataGridView1_SelectionChanged);
            this.dataGridView1.DoubleClick += new System.EventHandler(this.dataGridView1_DoubleClick);
            // 
            // PartNumber
            // 
            this.PartNumber.HeaderText = "ČÍSLO DÍLU";
            this.PartNumber.Name = "PartNumber";
            this.PartNumber.ReadOnly = true;
            // 
            // PartName
            // 
            this.PartName.HeaderText = "NÁZEV";
            this.PartName.Name = "PartName";
            this.PartName.ReadOnly = true;
            // 
            // Count
            // 
            this.Count.HeaderText = "POČET";
            this.Count.Name = "Count";
            this.Count.ReadOnly = true;
            // 
            // CountKat
            // 
            this.CountKat.HeaderText = "KATALOGOVÝ POČET";
            this.CountKat.Name = "CountKat";
            this.CountKat.ReadOnly = true;
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Enabled = false;
            this.button3.Font = new System.Drawing.Font("Consolas", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.button3.Location = new System.Drawing.Point(385, 453);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(167, 46);
            this.button3.TabIndex = 6;
            this.button3.Text = "EXPORT";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button4.Font = new System.Drawing.Font("Consolas", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.button4.Location = new System.Drawing.Point(558, 452);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(167, 46);
            this.button4.TabIndex = 7;
            this.button4.Text = "PŘIDAT RUČNĚ";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button5.Font = new System.Drawing.Font("Consolas", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.button5.Location = new System.Drawing.Point(807, 453);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(121, 46);
            this.button5.TabIndex = 8;
            this.button5.Text = "SMAZAT";
            this.button5.UseVisualStyleBackColor = true;
            // 
            // serialPort1
            // 
            this.serialPort1.PortName = "COM2";
            this.serialPort1.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(this.serialPort1_DataReceived);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.groupBox2.Location = new System.Drawing.Point(385, 124);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(543, 48);
            this.groupBox2.TabIndex = 9;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "POZICE";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Consolas", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            this.label5.Location = new System.Drawing.Point(86, 19);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(36, 19);
            this.label5.TabIndex = 0;
            this.label5.Text = "N/A";
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nASTAVENIToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(940, 24);
            this.menuStrip1.TabIndex = 10;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // nASTAVENIToolStripMenuItem
            // 
            this.nASTAVENIToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.pORTSKENERUToolStripMenuItem,
            this.kATALOGToolStripMenuItem});
            this.nASTAVENIToolStripMenuItem.Name = "nASTAVENIToolStripMenuItem";
            this.nASTAVENIToolStripMenuItem.Size = new System.Drawing.Size(82, 20);
            this.nASTAVENIToolStripMenuItem.Text = "NASTAVENI";
            // 
            // pORTSKENERUToolStripMenuItem
            // 
            this.pORTSKENERUToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cOM1ToolStripMenuItem,
            this.cOM2ToolStripMenuItem,
            this.cOM3ToolStripMenuItem,
            this.cOM4ToolStripMenuItem,
            this.cOM5ToolStripMenuItem,
            this.cOM6ToolStripMenuItem,
            this.cOM7ToolStripMenuItem,
            this.cOM8ToolStripMenuItem,
            this.cOM9ToolStripMenuItem,
            this.cOM10ToolStripMenuItem,
            this.cOM11ToolStripMenuItem,
            this.cOM12ToolStripMenuItem,
            this.cOM13ToolStripMenuItem,
            this.cOM14ToolStripMenuItem,
            this.cOM15ToolStripMenuItem});
            this.pORTSKENERUToolStripMenuItem.Name = "pORTSKENERUToolStripMenuItem";
            this.pORTSKENERUToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.pORTSKENERUToolStripMenuItem.Text = "PORT SKENERU";
            // 
            // cOM1ToolStripMenuItem
            // 
            this.cOM1ToolStripMenuItem.Checked = true;
            this.cOM1ToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cOM1ToolStripMenuItem.Name = "cOM1ToolStripMenuItem";
            this.cOM1ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM1ToolStripMenuItem.Text = "COM1";
            // 
            // cOM2ToolStripMenuItem
            // 
            this.cOM2ToolStripMenuItem.Name = "cOM2ToolStripMenuItem";
            this.cOM2ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM2ToolStripMenuItem.Text = "COM2";
            // 
            // cOM3ToolStripMenuItem
            // 
            this.cOM3ToolStripMenuItem.Name = "cOM3ToolStripMenuItem";
            this.cOM3ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM3ToolStripMenuItem.Text = "COM3";
            // 
            // cOM4ToolStripMenuItem
            // 
            this.cOM4ToolStripMenuItem.Name = "cOM4ToolStripMenuItem";
            this.cOM4ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM4ToolStripMenuItem.Text = "COM4";
            // 
            // cOM5ToolStripMenuItem
            // 
            this.cOM5ToolStripMenuItem.Name = "cOM5ToolStripMenuItem";
            this.cOM5ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM5ToolStripMenuItem.Text = "COM5";
            // 
            // cOM6ToolStripMenuItem
            // 
            this.cOM6ToolStripMenuItem.Name = "cOM6ToolStripMenuItem";
            this.cOM6ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM6ToolStripMenuItem.Text = "COM6";
            // 
            // cOM7ToolStripMenuItem
            // 
            this.cOM7ToolStripMenuItem.Name = "cOM7ToolStripMenuItem";
            this.cOM7ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM7ToolStripMenuItem.Text = "COM7";
            // 
            // cOM8ToolStripMenuItem
            // 
            this.cOM8ToolStripMenuItem.Name = "cOM8ToolStripMenuItem";
            this.cOM8ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM8ToolStripMenuItem.Text = "COM8";
            // 
            // cOM9ToolStripMenuItem
            // 
            this.cOM9ToolStripMenuItem.Name = "cOM9ToolStripMenuItem";
            this.cOM9ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM9ToolStripMenuItem.Text = "COM9";
            // 
            // cOM10ToolStripMenuItem
            // 
            this.cOM10ToolStripMenuItem.Name = "cOM10ToolStripMenuItem";
            this.cOM10ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM10ToolStripMenuItem.Text = "COM10";
            // 
            // cOM11ToolStripMenuItem
            // 
            this.cOM11ToolStripMenuItem.Name = "cOM11ToolStripMenuItem";
            this.cOM11ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM11ToolStripMenuItem.Text = "COM11";
            // 
            // cOM12ToolStripMenuItem
            // 
            this.cOM12ToolStripMenuItem.Name = "cOM12ToolStripMenuItem";
            this.cOM12ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM12ToolStripMenuItem.Text = "COM12";
            // 
            // cOM13ToolStripMenuItem
            // 
            this.cOM13ToolStripMenuItem.Name = "cOM13ToolStripMenuItem";
            this.cOM13ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM13ToolStripMenuItem.Text = "COM13";
            // 
            // cOM14ToolStripMenuItem
            // 
            this.cOM14ToolStripMenuItem.Name = "cOM14ToolStripMenuItem";
            this.cOM14ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM14ToolStripMenuItem.Text = "COM14";
            // 
            // cOM15ToolStripMenuItem
            // 
            this.cOM15ToolStripMenuItem.Name = "cOM15ToolStripMenuItem";
            this.cOM15ToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.cOM15ToolStripMenuItem.Text = "COM15";
            // 
            // statusStrip1
            // 
            this.statusStrip1.Location = new System.Drawing.Point(0, 509);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(940, 22);
            this.statusStrip1.TabIndex = 11;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // kATALOGToolStripMenuItem
            // 
            this.kATALOGToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.xLSXToolStripMenuItem,
            this.xLSToolStripMenuItem,
            this.xLSToolStripMenuItem1,
            this.zVOLITEXTRAToolStripMenuItem});
            this.kATALOGToolStripMenuItem.Name = "kATALOGToolStripMenuItem";
            this.kATALOGToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.kATALOGToolStripMenuItem.Text = "KATALOG";
            // 
            // xLSXToolStripMenuItem
            // 
            this.xLSXToolStripMenuItem.Checked = true;
            this.xLSXToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.xLSXToolStripMenuItem.Name = "xLSXToolStripMenuItem";
            this.xLSXToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.xLSXToolStripMenuItem.Text = "*.XLSX + *.XLS";
            // 
            // xLSToolStripMenuItem
            // 
            this.xLSToolStripMenuItem.Name = "xLSToolStripMenuItem";
            this.xLSToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.xLSToolStripMenuItem.Text = "*.XLSX";
            this.xLSToolStripMenuItem.Click += new System.EventHandler(this.xLSToolStripMenuItem_Click);
            // 
            // xLSToolStripMenuItem1
            // 
            this.xLSToolStripMenuItem1.Name = "xLSToolStripMenuItem1";
            this.xLSToolStripMenuItem1.Size = new System.Drawing.Size(180, 22);
            this.xLSToolStripMenuItem1.Text = "*.XLS";
            // 
            // zVOLITEXTRAToolStripMenuItem
            // 
            this.zVOLITEXTRAToolStripMenuItem.Name = "zVOLITEXTRAToolStripMenuItem";
            this.zVOLITEXTRAToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.zVOLITEXTRAToolStripMenuItem.Text = "ZVOLIT EXTRA";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(940, 531);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.treeView1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CInvent";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.DataGridViewTextBoxColumn PartNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn PartName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Count;
        private System.Windows.Forms.DataGridViewTextBoxColumn CountKat;
        private System.IO.Ports.SerialPort serialPort1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem nASTAVENIToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pORTSKENERUToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM1ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM2ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM3ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM4ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM5ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM6ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM7ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM8ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM9ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM10ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM11ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM12ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM13ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM14ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cOM15ToolStripMenuItem;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripMenuItem kATALOGToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem xLSXToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem xLSToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem xLSToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem zVOLITEXTRAToolStripMenuItem;
    }
}

