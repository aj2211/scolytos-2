namespace BeetleBase
{
    partial class Form2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.identifiercombo = new System.Windows.Forms.ComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.pinnedbox = new System.Windows.Forms.CheckBox();
            this.comboBox3 = new System.Windows.Forms.ComboBox();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.button3 = new System.Windows.Forms.Button();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox14 = new System.Windows.Forms.TextBox();
            this.textBox13 = new System.Windows.Forms.TextBox();
            this.getSpCodeButton = new System.Windows.Forms.Button();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.textBox9 = new System.Windows.Forms.TextBox();
            this.textBox8 = new System.Windows.Forms.TextBox();
            this.textBox7 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.button6 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.richTextBox2 = new System.Windows.Forms.RichTextBox();
            this.richTextBox3 = new System.Windows.Forms.RichTextBox();
            this.richTextBox4 = new System.Windows.Forms.RichTextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.showVials = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(12, 25);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(75, 20);
            this.textBox1.TabIndex = 0;
            this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
            this.textBox1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBox1_KeyPress);
            this.textBox1.KeyUp += new System.Windows.Forms.KeyEventHandler((a, b) => { this.textBox1_KeyUp(a, b, false); });
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(20, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Search Vial";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(11, 219);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(952, 129);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            this.dataGridView1.Click += new System.EventHandler(this.dataGridView1_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(11, 89);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(76, 44);
            this.button1.TabIndex = 0;
            this.button1.Text = "Edit Species in Vial";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(359, 144);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(102, 38);
            this.button2.TabIndex = 16;
            this.button2.Text = "Save Edited Species In Vial";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.richTextBox1);
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.identifiercombo);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.pinnedbox);
            this.groupBox1.Controls.Add(this.comboBox3);
            this.groupBox1.Controls.Add(this.comboBox2);
            this.groupBox1.Controls.Add(this.comboBox1);
            this.groupBox1.Controls.Add(this.button3);
            this.groupBox1.Controls.Add(this.button7);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.textBox14);
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.textBox13);
            this.groupBox1.Controls.Add(this.getSpCodeButton);
            this.groupBox1.Controls.Add(this.textBox11);
            this.groupBox1.Controls.Add(this.textBox9);
            this.groupBox1.Controls.Add(this.textBox8);
            this.groupBox1.Controls.Add(this.textBox7);
            this.groupBox1.Controls.Add(this.textBox5);
            this.groupBox1.Controls.Add(this.textBox4);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Location = new System.Drawing.Point(98, 25);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(572, 188);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Edit info";
            // 
            // richTextBox1
            // 
            this.richTextBox1.Enabled = false;
            this.richTextBox1.Location = new System.Drawing.Point(117, 20);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(103, 46);
            this.richTextBox1.TabIndex = 29;
            this.richTextBox1.Text = "";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(109, 143);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(142, 13);
            this.label15.TabIndex = 28;
            this.label15.Text = "Identifier (Collector/Museum)";
            // 
            // identifiercombo
            // 
            this.identifiercombo.FormattingEnabled = true;
            this.identifiercombo.Location = new System.Drawing.Point(112, 160);
            this.identifiercombo.Name = "identifiercombo";
            this.identifiercombo.Size = new System.Drawing.Size(100, 21);
            this.identifiercombo.TabIndex = 15;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(5, 143);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(40, 13);
            this.label14.TabIndex = 27;
            this.label14.Text = "Pinned";
            // 
            // pinnedbox
            // 
            this.pinnedbox.AutoSize = true;
            this.pinnedbox.Location = new System.Drawing.Point(50, 161);
            this.pinnedbox.Name = "pinnedbox";
            this.pinnedbox.Size = new System.Drawing.Size(15, 14);
            this.pinnedbox.TabIndex = 14;
            this.pinnedbox.UseVisualStyleBackColor = true;
            // 
            // comboBox3
            // 
            this.comboBox3.FormattingEnabled = true;
            this.comboBox3.Items.AddRange(new object[] {
            "31",
            "30",
            "29",
            "28",
            "27",
            "26",
            "25",
            "24",
            "23",
            "22",
            "21",
            "20",
            "19",
            "18",
            "17",
            "16",
            "15",
            "14",
            "13",
            "12",
            "11",
            "10",
            "09",
            "08",
            "07",
            "06",
            "05",
            "04",
            "03",
            "02",
            "01"});
            this.comboBox3.Location = new System.Drawing.Point(208, 118);
            this.comboBox3.Name = "comboBox3";
            this.comboBox3.Size = new System.Drawing.Size(40, 21);
            this.comboBox3.TabIndex = 10;
            // 
            // comboBox2
            // 
            this.comboBox2.FormattingEnabled = true;
            this.comboBox2.Items.AddRange(new object[] {
            "12",
            "11",
            "10",
            "09",
            "08",
            "07",
            "06",
            "05",
            "04",
            "03",
            "02",
            "01"});
            this.comboBox2.Location = new System.Drawing.Point(166, 118);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(40, 21);
            this.comboBox2.TabIndex = 9;
            // 
            // comboBox1
            // 
            this.comboBox1.FormattingEnabled = true;
            this.comboBox1.Items.AddRange(new object[] {
            "2030",
            "2029",
            "2028",
            "2027",
            "2026",
            "2025",
            "2024",
            "2023",
            "2022",
            "2021",
            "2020",
            "2019",
            "2018",
            "2017",
            "2016",
            "2015",
            "2014",
            "2013",
            "2012",
            "2011",
            "2010",
            "2009",
            "2008",
            "2007",
            "2006",
            "2005",
            "2004",
            "2003",
            "2002",
            "2001",
            "2000",
            "1999",
            "1998",
            "1997",
            "1996",
            "1995",
            "1994",
            "1993",
            "1992",
            "1991"});
            this.comboBox1.Location = new System.Drawing.Point(114, 118);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(50, 21);
            this.comboBox1.TabIndex = 8;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(466, 143);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(85, 38);
            this.button3.TabIndex = 0;
            this.button3.Text = "Cancel/Exit  (No Save)";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(463, 101);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(57, 13);
            this.label13.TabIndex = 25;
            this.label13.Text = "From Plate";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(340, 15);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(35, 13);
            this.label12.TabIndex = 24;
            this.label12.Text = "Count";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(358, 101);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(83, 13);
            this.label11.TabIndex = 23;
            this.label11.Text = "Number Loaned";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(250, 100);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(59, 13);
            this.label10.TabIndex = 22;
            this.label10.Text = "Loaned To";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(112, 100);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(77, 13);
            this.label9.TabIndex = 21;
            this.label9.Text = "Date Returned";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(4, 100);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(83, 13);
            this.label8.TabIndex = 20;
            this.label8.Text = "Borrowed Count";
            this.label8.Click += new System.EventHandler(this.label8_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(4, 59);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(71, 13);
            this.label7.TabIndex = 19;
            this.label7.Text = "Species Note";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(462, 16);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(30, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "Male";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(234, 15);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "Vial Number";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(5, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Species Code";
            // 
            // textBox14
            // 
            this.textBox14.Location = new System.Drawing.Point(360, 119);
            this.textBox14.Name = "textBox14";
            this.textBox14.Size = new System.Drawing.Size(100, 20);
            this.textBox14.TabIndex = 12;
            // 
            // textBox13
            // 
            this.textBox13.Location = new System.Drawing.Point(466, 119);
            this.textBox13.Name = "textBox13";
            this.textBox13.Size = new System.Drawing.Size(100, 20);
            this.textBox13.TabIndex = 13;
            // 
            // getSpCodeButton
            // 
            this.getSpCodeButton.Location = new System.Drawing.Point(5, 33);
            this.getSpCodeButton.Name = "getSpCodeButton";
            this.getSpCodeButton.Size = new System.Drawing.Size(60, 22);
            this.getSpCodeButton.TabIndex = 1;
            this.getSpCodeButton.Text = "Select";
            this.getSpCodeButton.UseVisualStyleBackColor = true;
            this.getSpCodeButton.Click += new System.EventHandler(this.getSpCodeButton_Click);
            // 
            // textBox11
            // 
            this.textBox11.Location = new System.Drawing.Point(465, 34);
            this.textBox11.Name = "textBox11";
            this.textBox11.Size = new System.Drawing.Size(100, 20);
            this.textBox11.TabIndex = 5;
            // 
            // textBox9
            // 
            this.textBox9.Location = new System.Drawing.Point(343, 34);
            this.textBox9.Name = "textBox9";
            this.textBox9.Size = new System.Drawing.Size(100, 20);
            this.textBox9.TabIndex = 3;
            // 
            // textBox8
            // 
            this.textBox8.Enabled = false;
            this.textBox8.Location = new System.Drawing.Point(237, 34);
            this.textBox8.Name = "textBox8";
            this.textBox8.Size = new System.Drawing.Size(100, 20);
            this.textBox8.TabIndex = 2;
            // 
            // textBox7
            // 
            this.textBox7.Location = new System.Drawing.Point(254, 119);
            this.textBox7.Name = "textBox7";
            this.textBox7.Size = new System.Drawing.Size(100, 20);
            this.textBox7.TabIndex = 11;
            // 
            // textBox5
            // 
            this.textBox5.Location = new System.Drawing.Point(7, 119);
            this.textBox5.Name = "textBox5";
            this.textBox5.Size = new System.Drawing.Size(100, 20);
            this.textBox5.TabIndex = 7;
            // 
            // textBox4
            // 
            this.textBox4.Location = new System.Drawing.Point(7, 77);
            this.textBox4.Name = "textBox4";
            this.textBox4.Size = new System.Drawing.Size(558, 20);
            this.textBox4.TabIndex = 6;
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(68, 34);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(39, 20);
            this.textBox2.TabIndex = 1;
            this.textBox2.TabStop = false;
            // 
            // button6
            // 
            this.button6.Location = new System.Drawing.Point(12, 50);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(75, 35);
            this.button6.TabIndex = 0;
            this.button6.Text = "Add Species To Vial";
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(257, 145);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(96, 36);
            this.button7.TabIndex = 17;
            this.button7.Text = "Save New Species In Vial";
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button8
            // 
            this.button8.Location = new System.Drawing.Point(12, 137);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(75, 47);
            this.button8.TabIndex = 100;
            this.button8.Text = "Delete Species In Vial";
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.pictureBox1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.pictureBox1.Location = new System.Drawing.Point(676, 31);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(160, 182);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 6;
            this.pictureBox1.TabStop = false;
            // 
            // richTextBox2
            // 
            this.richTextBox2.Enabled = false;
            this.richTextBox2.Location = new System.Drawing.Point(842, 47);
            this.richTextBox2.Name = "richTextBox2";
            this.richTextBox2.Size = new System.Drawing.Size(121, 42);
            this.richTextBox2.TabIndex = 7;
            this.richTextBox2.Text = "";
            // 
            // richTextBox3
            // 
            this.richTextBox3.Enabled = false;
            this.richTextBox3.Location = new System.Drawing.Point(842, 109);
            this.richTextBox3.Name = "richTextBox3";
            this.richTextBox3.Size = new System.Drawing.Size(121, 42);
            this.richTextBox3.TabIndex = 8;
            this.richTextBox3.Text = "";
            this.richTextBox3.TextChanged += new System.EventHandler(this.richTextBox3_TextChanged);
            // 
            // richTextBox4
            // 
            this.richTextBox4.Enabled = false;
            this.richTextBox4.Location = new System.Drawing.Point(843, 170);
            this.richTextBox4.Name = "richTextBox4";
            this.richTextBox4.Size = new System.Drawing.Size(121, 42);
            this.richTextBox4.TabIndex = 9;
            this.richTextBox4.Text = "";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(841, 30);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(54, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "Vial Origin";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(841, 91);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(63, 13);
            this.label16.TabIndex = 11;
            this.label16.Text = "Vial Country";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(841, 154);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(66, 13);
            this.label17.TabIndex = 12;
            this.label17.Text = "Host or Trap";
            // 
            // showVials
            // 
            this.showVials.Location = new System.Drawing.Point(12, 190);
            this.showVials.Name = "showVials";
            this.showVials.Size = new System.Drawing.Size(75, 23);
            this.showVials.TabIndex = 0;
            this.showVials.TabStop = false;
            this.showVials.Text = "Show Vials";
            this.showVials.UseVisualStyleBackColor = true;
            this.showVials.Click += new System.EventHandler(this.showVials_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(976, 360);
            this.Controls.Add(this.showVials);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.richTextBox4);
            this.Controls.Add(this.richTextBox3);
            this.Controls.Add(this.richTextBox2);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.button8);
            this.Controls.Add(this.button6);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form2";
            this.Text = "Specimens in Vials Manager";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form2_FormClosing);
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBox11;
        private System.Windows.Forms.TextBox textBox9;
        private System.Windows.Forms.TextBox textBox8;
        private System.Windows.Forms.TextBox textBox7;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox14;
        private System.Windows.Forms.TextBox textBox13;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.ComboBox comboBox3;
        private System.Windows.Forms.ComboBox comboBox2;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.CheckBox pinnedbox;
        private System.Windows.Forms.Label label15;
        public System.Windows.Forms.ComboBox identifiercombo;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.RichTextBox richTextBox2;
        private System.Windows.Forms.RichTextBox richTextBox3;
        private System.Windows.Forms.RichTextBox richTextBox4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Button showVials;
        private System.Windows.Forms.Button getSpCodeButton;
    }
}