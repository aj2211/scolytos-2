namespace BeetleBase
{
    partial class Form5
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("[Capture Type]");
            System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("[Country]");
            System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("[Experiment]");
            System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("[Fungus]");
            System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("[Province]");
            System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("[Collector/Museum]");
            System.Windows.Forms.TreeNode treeNode7 = new System.Windows.Forms.TreeNode("Dropdowns", new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2,
            treeNode3,
            treeNode4,
            treeNode5,
            treeNode6});
            System.Windows.Forms.TreeNode treeNode8 = new System.Windows.Forms.TreeNode("[Species_Table]");
            System.Windows.Forms.TreeNode treeNode9 = new System.Windows.Forms.TreeNode("Species Data", new System.Windows.Forms.TreeNode[] {
            treeNode8});
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form5));
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // treeView1
            // 
            this.treeView1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.treeView1.HideSelection = false;
            this.treeView1.Indent = 10;
            this.treeView1.ItemHeight = 32;
            this.treeView1.Location = new System.Drawing.Point(12, 12);
            this.treeView1.Margin = new System.Windows.Forms.Padding(0, 0, 3, 3);
            this.treeView1.Name = "treeView1";
            treeNode1.Name = "CaptureType";
            treeNode1.Text = "[Capture Type]";
            treeNode2.Name = "Country";
            treeNode2.Text = "[Country]";
            treeNode3.Name = "Experiment";
            treeNode3.Text = "[Experiment]";
            treeNode4.Name = "Fungus";
            treeNode4.Text = "[Fungus]";
            treeNode5.Name = "Province";
            treeNode5.Text = "[Province]";
            treeNode6.Name = "collectorMuseum";
            treeNode6.Text = "[Collector/Museum]";
            treeNode7.Name = "DropdownNode";
            treeNode7.Text = "Dropdowns";
            treeNode8.Name = "SpeciesTable";
            treeNode8.Text = "[Species_Table]";
            treeNode9.Name = "SpeciesTableRoot";
            treeNode9.Text = "Species Data";
            this.treeView1.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode7,
            treeNode9});
            this.treeView1.ShowPlusMinus = false;
            this.treeView1.Size = new System.Drawing.Size(106, 332);
            this.treeView1.TabIndex = 0;
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            this.treeView1.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeView1_NodeMouseClick);
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(132, 12);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(189, 299);
            this.dataGridView1.TabIndex = 1;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.UserAddedRow += new System.Windows.Forms.DataGridViewRowEventHandler(this.dataGridView1_UserAddedRow);
            this.dataGridView1.UserDeletingRow += new System.Windows.Forms.DataGridViewRowCancelEventHandler(this.dataGridView1_UserDeletingRow);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(132, 321);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(189, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Save Changes";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form5
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(333, 350);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.treeView1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form5";
            this.Text = "Table Editor";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form5_FormClosing);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
    }
}