namespace Scolytos2
{
    partial class speciesLookUpForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(speciesLookUpForm));
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.prevbutton = new System.Windows.Forms.Button();
            this.nextbutton = new System.Windows.Forms.Button();
            this.picturelabel = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Location = new System.Drawing.Point(12, 195);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(168, 21);
            this.comboBox1.TabIndex = 0;
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectionChangeCommitted);
            this.comboBox1.SelectionChangeCommitted += new System.EventHandler(this.comboBox1_SelectionChangeCommitted);
            this.comboBox1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.comboBox1_KeyUp);
            this.comboBox1.Leave += new System.EventHandler(this.comboBox1_SelectionChangeCommitted);
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Location = new System.Drawing.Point(12, 238);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(168, 21);
            this.comboBox2.TabIndex = 1;
            this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectionChangeCommitted);
            this.comboBox2.SelectionChangeCommitted += new System.EventHandler(this.comboBox2_SelectionChangeCommitted);
            this.comboBox2.KeyUp += new System.Windows.Forms.KeyEventHandler(this.comboBox2_KeyUp);
            this.comboBox2.Leave += new System.EventHandler(this.comboBox2_SelectionChangeCommitted);
            // 
            // comboBox3
            // 
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Location = new System.Drawing.Point(12, 279);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(167, 21);
            this.comboBox3.TabIndex = 2;
            this.comboBox3.SelectedIndexChanged += new System.EventHandler(this.comboBox3_SelectionChangeCommitted);
            this.comboBox3.SelectionChangeCommitted += new System.EventHandler(this.comboBox3_SelectionChangeCommitted);
            this.comboBox3.KeyUp += new System.Windows.Forms.KeyEventHandler(this.comboBox3_KeyUp);
            this.comboBox3.Leave += new System.EventHandler(this.comboBox3_SelectionChangeCommitted);
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(198, 14);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(493, 352);
            this.dataGridView1.TabIndex = 4;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 179);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(31, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Tribe";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 222);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(38, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Genus";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 263);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Species";
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(12, 33);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(167, 106);
            this.pictureBox1.TabIndex = 9;
            this.pictureBox1.TabStop = false;
            // 
            // prevbutton
            // 
            this.prevbutton.Location = new System.Drawing.Point(12, 149);
            this.prevbutton.Name = "prevbutton";
            this.prevbutton.Size = new System.Drawing.Size(75, 23);
            this.prevbutton.TabIndex = 10;
            this.prevbutton.Text = "Prev";
            this.prevbutton.UseVisualStyleBackColor = true;
            // 
            // nextbutton
            // 
            this.nextbutton.Location = new System.Drawing.Point(105, 149);
            this.nextbutton.Name = "nextbutton";
            this.nextbutton.Size = new System.Drawing.Size(75, 23);
            this.nextbutton.TabIndex = 11;
            this.nextbutton.Text = "Next";
            this.nextbutton.UseVisualStyleBackColor = true;
            // 
            // picturelabel
            // 
            this.picturelabel.AutoSize = true;
            this.picturelabel.Location = new System.Drawing.Point(80, 12);
            this.picturelabel.Name = "picturelabel";
            this.picturelabel.Size = new System.Drawing.Size(34, 13);
            this.picturelabel.TabIndex = 12;
            this.picturelabel.Text = "0 of 0";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 310);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(168, 23);
            this.button1.TabIndex = 13;
            this.button1.Text = "Reset Dropdowns";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button2_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(12, 343);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(168, 23);
            this.button2.TabIndex = 14;
            this.button2.Text = "Explore Species in Vial";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // speciesLookUpForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(703, 378);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.picturelabel);
            this.Controls.Add(this.nextbutton);
            this.Controls.Add(this.prevbutton);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.comboBox3);
            this.Controls.Add(this.comboBox2);
            this.Controls.Add(this.comboBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "speciesLookUpForm";
            this.Text = "Species Lookup";
            this.Load += new System.EventHandler(this.speciesLookUpForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button prevbutton;
        private System.Windows.Forms.Button nextbutton;
        private System.Windows.Forms.Label picturelabel;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}