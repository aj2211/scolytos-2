using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace BeetleBase
{
    public partial class Form4 : Form
    {
        public void form4editenable(bool a)
        {
            this.checkBox1.Enabled = a;
            this.comboBox1.Enabled = a;
            this.comboBox2.Enabled = a;
            this.comboBox3.Enabled = a;
            this.comboBox4.Enabled = a;
            this.comboBox5.Enabled = a;
            this.comboBox6.Enabled = a;
            this.fieldVialTextBox4.Enabled = a;
            this.hostTrapTextBox5.Enabled = a;
            this.localityTextBox8.Enabled = a;
            this.textBox9.Enabled = a;
            this.comboBox7.Enabled = false;
            this.comboBox8.Enabled = false;
            this.comboBox9.Enabled = false;
            this.textBox13.Enabled = a;
            this.comboBox10.Enabled = a;
            this.comboBox11.Enabled = a;
            this.comboBox12.Enabled = a;
            this.button1.Enabled = !a;
            this.button2.Enabled = a;
            this.button3.Enabled = a;
            this.dataGridView1.Enabled = !a;
            this.textBox1.Enabled = !a;
            this.button7.Enabled = !a;
            if (!a)
            {
                this.checkBox1.Checked = false;
            }
        }

        public void initializeComponent()
        {
            InitializeComponent();
            this.button1.Focus();
            form4editenable(false);
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new System.Drawing.Point(0, 0);
            this.button5.Enabled = false;
            this.dataGridView1.DataSource = this.thefile.main2.Tables[0];        
            this.dataGridView1.Columns[0].Width = 40;
            this.dataGridView1.Columns[1].Width = 80;
            this.dataGridView1.Columns[2].Width = 40;
            this.dataGridView1.Columns[3].Width = 100;
            this.dataGridView1.Columns[4].Width = 108;
            this.dataGridView1.Columns[5].Width = 70;
            this.dataGridView1.Columns[6].Width = 70;
            this.dataGridView1.Columns[7].Width = 70;
            this.dataGridView1.Columns[8].Width = 70;
            this.dataGridView1.Columns[9].Width = 70;
            this.dataGridView1.Columns[10].Width = 50;
            this.dataGridView1.Columns[11].Width = 150;            
            this.dataGridView1.Columns[12].Width = 70;
            this.dataGridView1.Columns[13].Width = 70;
            this.dataGridView1.Columns[14].Width = 70;
            DataSet dropdowns = new DataSet();
            OleDbCommand first = new OleDbCommand("SELECT * FROM [COLLECTIONS- Drop Down Capture Type]", this.thefile.dbo);
            OleDbDataAdapter dropdown = new OleDbDataAdapter(first);
            dropdown.Fill(dropdowns, "capturetype");
            dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Collector/Museum]";
            dropdown.Fill(dropdowns, "collectormuseum");
            dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Country]";
            dropdown.Fill(dropdowns, "country");
            dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Experiment]";
            dropdown.Fill(dropdowns, "experiment");
            dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Fungus]";
            dropdown.Fill(dropdowns, "fungus");
            dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Province]";
            dropdown.Fill(dropdowns, "province");
            int capturetypecount = dropdowns.Tables["capturetype"].Rows.Count;
            int collectormuseumcount = dropdowns.Tables["collectormuseum"].Rows.Count;
            int countrycount = dropdowns.Tables["country"].Rows.Count;
            int experimentcount = dropdowns.Tables["experiment"].Rows.Count;
            int funguscount = dropdowns.Tables["fungus"].Rows.Count;
            int provincecount = dropdowns.Tables["province"].Rows.Count;
            for (int i = 0; i < capturetypecount; i++)
            {
                comboBox2.Items.Add(dropdowns.Tables["capturetype"].Rows[i][0]);
            }
            for (int i = 0; i < countrycount; i++)
            {
                comboBox5.Items.Add(dropdowns.Tables["country"].Rows[i][0]);
            }
            for (int i = 0; i < collectormuseumcount; i++)
            {
                comboBox6.Items.Add(dropdowns.Tables["collectormuseum"].Rows[i][0]);
            }
            for (int i = 0; i < experimentcount; i++)
            {
                comboBox1.Items.Add(dropdowns.Tables["experiment"].Rows[i][0]);
            }
            for (int i = 0; i < funguscount; i++)
            {
                comboBox3.Items.Add(dropdowns.Tables["fungus"].Rows[i][0]);
            }
            for (int i = 0; i < provincecount; i++)
            {
                comboBox4.Items.Add(dropdowns.Tables["province"].Rows[i][0]);
            }
            dropdown.Dispose();
        }

        public Form4(mutual mutual, DB thefile)
        {

            this.mutual = mutual;
            this.thefile = thefile;
        }

        public mutual mutual;
        public DB thefile;
        public Form2 aa;
        public void textBox1_KeyUp(object sender, KeyEventArgs e, bool loop)
        {
            if (itsUnderControl)
            {
                itsUnderControl = false;
                return;
            }
            if (e != null && e.KeyCode.ToString() != "Return")
            {
                return;
            }
            // this is the default value if noting is entered
            string cmd = "SELECT COLLECTIONS.vial, COLLECTIONS.experiment, COLLECTIONS.field_vial, COLLECTIONS.host_or_trap, COLLECTIONS.[capture->storage], COLLECTIONS.fungus, COLLECTIONS.Country, COLLECTIONS.province, COLLECTIONS.county, COLLECTIONS.locality, COLLECTIONS.date, COLLECTIONS.VIAL_note, COLLECTIONS.[collector/museum], COLLECTIONS.[pair/family], COLLECTIONS.date_collected FROM [COLLECTIONS] WHERE vial Is Not Null";
            if (textBox1.Text.Trim() != "")
            {
                if (textBox1.Text.Contains("-"))
                {
                    string[] searchrange = textBox1.Text.Split('-');
                    var minvial = searchrange[0];
                    var maxvial = searchrange[1];
                    cmd = "SELECT COLLECTIONS.vial, COLLECTIONS.experiment, COLLECTIONS.field_vial, COLLECTIONS.host_or_trap, COLLECTIONS.[capture->storage], COLLECTIONS.fungus, COLLECTIONS.Country, COLLECTIONS.province, COLLECTIONS.county, COLLECTIONS.locality, COLLECTIONS.date, COLLECTIONS.VIAL_note, COLLECTIONS.[collector/museum], COLLECTIONS.[pair/family], COLLECTIONS.date_collected FROM [COLLECTIONS] WHERE vial BETWEEN " + minvial + " and " +maxvial;                
                }
                else
                {
                    cmd = "SELECT COLLECTIONS.vial, COLLECTIONS.experiment, COLLECTIONS.field_vial, COLLECTIONS.host_or_trap, COLLECTIONS.[capture->storage], COLLECTIONS.fungus, COLLECTIONS.Country, COLLECTIONS.province, COLLECTIONS.county, COLLECTIONS.locality, COLLECTIONS.date, COLLECTIONS.VIAL_note, COLLECTIONS.[collector/museum], COLLECTIONS.[pair/family], COLLECTIONS.date_collected FROM [COLLECTIONS] WHERE vial =" + textBox1.Text;
                }
                aa.textBox1.Text = textBox1.Text;
                if (!loop)
                {
                    aa.textBox1_KeyUp(null, null, true);
                }
            }
            OleDbCommand vialsearch = new OleDbCommand(cmd, this.thefile.dbo);
            OleDbDataAdapter vialadapter = new OleDbDataAdapter(vialsearch);
            DataSet vials = new DataSet();
            vialadapter.Fill(vials, "INIT");
            vials.Tables["INIT"].Columns["date"].DataType = System.Type.GetType("System.DateTime");
            this.dataGridView1.DataSource = vials.Tables["INIT"];
            vialadapter.Dispose();
        }

        public void textBox1_KeyUp2(object sender, KeyEventArgs e)
        {
            string cmd = "SELECT COLLECTIONS.vial, COLLECTIONS.experiment, COLLECTIONS.field_vial, COLLECTIONS.host_or_trap, COLLECTIONS.[capture->storage], COLLECTIONS.fungus, COLLECTIONS.Country, COLLECTIONS.province, COLLECTIONS.county, COLLECTIONS.locality, COLLECTIONS.date, COLLECTIONS.VIAL_note, COLLECTIONS.[collector/museum], COLLECTIONS.[pair/family], COLLECTIONS.date_collected FROM [COLLECTIONS] WHERE vial =" + textBox1.Text + "";
            OleDbCommand vialsearch = new OleDbCommand(cmd, this.thefile.dbo);
            OleDbDataAdapter vialadapter = new OleDbDataAdapter(vialsearch);
            DataSet vials = new DataSet();
            vialadapter.Fill(vials);
            this.dataGridView1.DataSource = vials.Tables[0];
            vialadapter.Dispose();
        }

        // Button 1 is to edit the vial.

        public void button1_Click(object sender, EventArgs e)
        {
            this.editting = true;
            this.button4.Enabled = false;
            button8newsimilarvial.Enabled = false;
            DataGridViewSelectedRowCollection editted;
            DataGridViewCellCollection col;
            if (this.dataGridView1.SelectedCells.Count == 0 && this.dataGridView1.SelectedRows.Count == 0)
            {
                return;
            }
            if (this.dataGridView1.SelectedCells.Count == 1 && this.dataGridView1.SelectedRows.Count < 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else if (this.dataGridView1.SelectedCells.Count > 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else
            {
                editted = this.dataGridView1.SelectedRows;
            }
            form4editenable(true);
            col = editted[0].Cells;

            char[] delimiterChars = { '/' };
            string[] words1 = col[10].Value.ToString().Split(delimiterChars);
            if (words1.Length > 1)
            {
                if (Int32.Parse(words1[0]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
                if (Int32.Parse(words1[1]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
            }
            string[] words2 = col[13].Value.ToString().Split(delimiterChars);
            if (words2.Length > 1)
            {
                if (Int32.Parse(words2[0]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
                if (Int32.Parse(words2[1]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
            }
            groupBox1.Text = "Edit Vial " + col[0].Value.ToString();
            this.currentvial = col[0].Value.ToString();
            comboBox1.Text = col[1].Value.ToString();
            fieldVialTextBox4.Text = col[2].Value.ToString();
            hostTrapTextBox5.Text = col[3].Value.ToString();
            comboBox2.Text = col[4].Value.ToString();
            comboBox3.Text = col[5].Value.ToString();
            localityTextBox8.Text = col[9].Value.ToString();
            textBox9.Text = col[8].Value.ToString();
            comboBox4.Text = col[7].Value.ToString();
            comboBox5.Text = col[6].Value.ToString();
            if (col[14].Value.ToString() == "True")
            {
                checkBox1.Checked = true;
            }
            else
            {
                checkBox1.Checked = false;
            }
            //            textBox12.Text = col[10].Value.ToString();
            if (words1.Length > 1)
            {
//                comboBox7.Text = words1[2];
//                comboBox8.Text = words1[0];
//                comboBox9.Text = words1[1];
            }
            comboBox7.Text = DateTime.Now.Year.ToString();
            comboBox8.Text = (DateTime.Now.Month > 9) ? DateTime.Now.Month.ToString() : ("0" + DateTime.Now.Month.ToString());
            comboBox9.Text = (DateTime.Now.Day > 9) ? DateTime.Now.Day.ToString() : ("0" + DateTime.Now.Day.ToString());
            textBox13.Text = col[11].Value.ToString();
            //            textBox15.Text = col[12].Value.ToString();
            if (words2.Length > 1)
            {
                comboBox10.Text = words2[2];
                comboBox11.Text = words2[0];
                comboBox12.Text = words2[1];
            }
            // Added AJJ 2017-06-21
            comboBox6.Text = col[12].Value.ToString();
            if (col[14].Value.ToString().Length == 10)
            {
                comboBox10.Text = col[14].Value.ToString().Substring(0,4);
                comboBox11.Text = col[14].Value.ToString().Substring(5,2);
                comboBox12.Text = col[14].Value.ToString().Substring(8,2);
            }
   
        }

        // Not sure what this is for?
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public string currentvial;
        public bool editting = false;
        public bool itsUnderControl;
        //  Note, this is the start of the update (ie click save edited vial)
        //  ...
        private void button2_Click(object sender, EventArgs e)
        {
            this.editting = false;
            if (this.thefile.dbo.State != ConnectionState.Open)
            {
                this.thefile.dbo.Open();
            }
            string updatemaster = "UPDATE [COLLECTIONS] SET ";
            // updatemaster += " [vial] = (MAX([vial]) + 1), ";
            updatemaster += " [experiment] = '" + comboBox1.Text + "'";
            //if (fieldVialTextBox4.Text == "")
            //{
            //    //                updatemaster += ", [Count] = null";
            //}
            //else
            //{
            //    updatemaster += ", [field_vial] = " + fieldVialTextBox4.Text;
            //}
            updatemaster += ", [field_vial] = '" + fieldVialTextBox4.Text + "'";
            updatemaster += ", [host_or_trap] = '" + hostTrapTextBox5.Text + "'";
            updatemaster += ", [capture->storage] = '" + comboBox2.Text + "'";
            updatemaster += ", [fungus] = '" + comboBox3.Text + "'";
            updatemaster += ", [locality] = '" + localityTextBox8.Text + "'";
            updatemaster += ", [county] = '" + textBox9.Text + "'";
            updatemaster += ", [province] = '" + comboBox4.Text + "'";
            updatemaster += ", [country] = '" + comboBox5.Text + "'";
            if (checkBox1.Checked)
            {
                updatemaster += ", [pair/family] = True";
            }
            else
            {
                updatemaster += ", [pair/family] = False";
            }
            if
                (
                comboBox7.Text != ""
                && comboBox7.Text != " "
                && comboBox7.Text != "  "
                && comboBox8.Text != ""
                && comboBox8.Text != " "
                && comboBox8.Text != "  "
                && comboBox9.Text != ""
                && comboBox9.Text != " "
                && comboBox9.Text != "  "
                )
            {
                updatemaster += ", [date] = '" + comboBox8.Text + "/" + comboBox9.Text + "/" + comboBox7.Text + "'";
            }
            updatemaster += ", [VIAL_note] = '" + textBox13.Text + "'";
            updatemaster += ", [collector/museum] = '" + comboBox6.Text + "'";
            if
                (
                comboBox10.Text == ""
                )
            {
                comboBox10.Text = "????";
            }
            if
                (
                comboBox11.Text == ""
                )
            {
                comboBox11.Text = "??";
            }  
            if
                (
                comboBox12.Text == ""
                )
            {
                comboBox12.Text = "??";
            }            
            if 
                (
                comboBox11.Text != "" 
                && comboBox11.Text != " " 
                && comboBox11.Text != "  "
                && comboBox12.Text != ""
                && comboBox12.Text != " "
                && comboBox12.Text != "  "
                && comboBox10.Text != ""
                && comboBox10.Text != " "
                && comboBox10.Text != "  "
                )
            {
                updatemaster += ", [date_collected] = '" + comboBox10.Text + "-" + comboBox11.Text + "-" + comboBox12.Text + "'";
            }
            updatemaster += " WHERE vial = " + this.currentvial;
            try
            {
                OleDbCommand up = new OleDbCommand(updatemaster, this.thefile.dbo);
                OleDbDataAdapter upd = new OleDbDataAdapter();
                upd.UpdateCommand = up;
                upd.UpdateCommand.ExecuteNonQuery();
            }
            catch (OleDbException error)
            {
                MessageBox.Show("Unable to write. Check to make sure information is valid!");
                MessageBox.Show(error.ToString());
                MessageBox.Show(updatemaster);
                return;
            }
            textBox1_KeyUp(null, null, false);
            button3_Click(null, null);
            form4editenable(false);
            showSpecies.Enabled = true;
        }
        // Button 3 is exit no save
        private void button3_Click(object sender, EventArgs e)
        {
            this.editting = false;
            this.comboBox1.Text = "";
            this.fieldVialTextBox4.Text = "";
            this.hostTrapTextBox5.Text = "";
            this.comboBox2.Text = "";
            this.comboBox3.Text = "";
            this.comboBox4.Text = "";
            this.comboBox5.Text = "";
            this.comboBox6.Text = "";
            this.localityTextBox8.Text = "";
            this.textBox9.Text = "";
            this.comboBox7.Text = "";
            this.comboBox8.Text = "";
            this.comboBox9.Text = "";
            this.textBox13.Text = "";
            this.comboBox10.Text = "";
            this.comboBox11.Text = "";
            this.comboBox12.Text = "";
            this.button5.Enabled = false;
            this.button8newsimilarvial.Enabled = true;
            this.button4.Enabled = true;
            this.showSpecies.Enabled = true;
            form4editenable(false);
            groupBox1.Text = "Vial Info";
        }

        public void button4newvial_Click(object sender, EventArgs e)
        {
            this.textBox1.Text = "  ";
            this.comboBox1.Text = "";
            this.fieldVialTextBox4.Text = "";
            this.hostTrapTextBox5.Text = "";
            this.comboBox2.Text = "";
            this.comboBox3.Text = "no";
            this.comboBox4.Text = "";
            this.comboBox5.Text = "";
            this.comboBox6.Text = "";
            this.localityTextBox8.Text = "";
            this.textBox9.Text = "";
            this.comboBox7.Text = "";
            this.comboBox8.Text = "";
            this.comboBox9.Text = "";
            this.textBox13.Text = "";
            this.comboBox10.Text = "";
            this.comboBox11.Text = "";
            this.comboBox12.Text = "";
            form4editenable(true);
            button2.Enabled = false;
            button3.Enabled = true;
            dataGridView1.ClearSelection();
            button4.Enabled = false;
            button5.Enabled = true;
            button8newsimilarvial.Enabled = false;
            this.showSpecies.Enabled = false;
            string year = DateTime.Today.Year.ToString();
            string month = DateTime.Today.Month.ToString("D2");
            string day = DateTime.Today.Day.ToString("D2");
            comboBox7.Text = year;
            comboBox8.Text = month;
            comboBox9.Text = day;
        }
       public void button8_Click(object sender, EventArgs e)
        {
            this.textBox1.Text = "   ";
            button8newsimilarvial.Enabled = false;
            this.showSpecies.Enabled = false;
            DataGridViewSelectedRowCollection editted;
            DataGridViewCellCollection col;
            if (this.dataGridView1.SelectedCells.Count == 0 && this.dataGridView1.SelectedRows.Count == 0)
            {
                return;
            }
            if (this.dataGridView1.SelectedCells.Count == 1 && this.dataGridView1.SelectedRows.Count < 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else if (this.dataGridView1.SelectedCells.Count > 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else
            {
                editted = this.dataGridView1.SelectedRows;
            }
            form4editenable(true);
            col = editted[0].Cells;

            char[] delimiterChars = { '/' };
            string[] words1 = col[10].Value.ToString().Split(delimiterChars);
            if (words1.Length > 1)
            {
                if (Int32.Parse(words1[0]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
                if (Int32.Parse(words1[1]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
            }
            string[] words2 = col[13].Value.ToString().Split(delimiterChars);
            if (words2.Length > 1)
            {
                if (Int32.Parse(words2[0]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
                if (Int32.Parse(words2[1]) < 10)
                {
                    words1[0] = "0" + words1[0];
                }
            }
            groupBox1.Text = "New vial ";
            this.currentvial = "";
            comboBox1.Text = col[1].Value.ToString();
            fieldVialTextBox4.Text = col[2].Value.ToString();
            hostTrapTextBox5.Text = col[3].Value.ToString();
            comboBox2.Text = col[4].Value.ToString();
            comboBox3.Text = col[5].Value.ToString();
            localityTextBox8.Text = col[9].Value.ToString();
            textBox9.Text = col[8].Value.ToString();
            comboBox4.Text = col[7].Value.ToString();
            comboBox5.Text = col[6].Value.ToString();
            if (col[14].Value.ToString() == "True")
            {
                checkBox1.Checked = true;
            }
            else
            {
                checkBox1.Checked = false;
            }
            //            textBox12.Text = col[10].Value.ToString();
            if (words1.Length > 1)
            {
//                comboBox7.Text = words1[2];
//                comboBox8.Text = words1[0];
//                comboBox9.Text = words1[1];
            }
            comboBox7.Text = DateTime.Now.Year.ToString();
            comboBox8.Text = (DateTime.Now.Month > 9) ? DateTime.Now.Month.ToString() : ("0" + DateTime.Now.Month.ToString());
            comboBox9.Text = (DateTime.Now.Day > 9) ? DateTime.Now.Day.ToString() : ("0" + DateTime.Now.Day.ToString());
            textBox13.Text = col[11].Value.ToString();
            //            textBox15.Text = col[12].Value.ToString();
            if (words2.Length > 1)
            {
                comboBox10.Text = words2[2];
                comboBox11.Text = words2[0];
                comboBox12.Text = words2[1];
            }
            // Added AJJ 2017-06-21
            comboBox6.Text = col[12].Value.ToString();
            if (col[14].Value.ToString().Length == 10)
            {
                comboBox10.Text = col[14].Value.ToString().Substring(0,4);
                comboBox11.Text = col[14].Value.ToString().Substring(5,2);
                comboBox12.Text = col[14].Value.ToString().Substring(8,2);
            }
            
            button2.Enabled = false;
            button3.Enabled = true;
            button4.Enabled = false;
            button5.Enabled = true;  
   
        }

        public void button5_Click(object sender, EventArgs e)
        {
            if (this.thefile.dbo.State != ConnectionState.Open)
            {
                this.thefile.dbo.Open();
            }
            string selectBefore = "SELECT (MAX([vial]) + 1) FROM [COLLECTIONS];";
            OleDbCommand getNextVial = new OleDbCommand(selectBefore, this.thefile.dbo);
            DataSet nextHighest = new DataSet();
            OleDbDataAdapter getNextHighest = new OleDbDataAdapter(getNextVial);
            getNextHighest.Fill(nextHighest, "next");
            var next = Int32.Parse(nextHighest.Tables["next"].Rows[0][0].ToString());
            string insertmaster = "INSERT INTO [COLLECTIONS] ([vial], [experiment],";
            //            if (fieldVialTextBox4.Text != "" && fieldVialTextBox4.Text != " " && fieldVialTextBox4.Text != "  ")
            //            {
            //                insertmaster += " [field_vial],";
            //            }
            insertmaster += " [field_vial], [host_or_trap], [capture->storage], [fungus], [locality], [county], [province], [country], [date], [VIAL_note], [collector/museum], [date_collected], [pair/family]) VALUES (";
            insertmaster += next + ", '" + comboBox1.Text + "', ";
            //            if (fieldVialTextBox4.Text != "" && fieldVialTextBox4.Text != " " && fieldVialTextBox4.Text != "  ")
            //            {
            //                var num = -1;
            //                Int32.TryParse(fieldVialTextBox4.Text, out num);
            //                MessageBox.Show(num.ToString());
            //                if (num != -1 && num != 0)
            //                {
            //                    insertmaster += fieldVialTextBox4.Text + ", ";
            //                }
            //                else
            //                {
            //                    insertmaster += "'" + fieldVialTextBox4.Text + "', ";
            //                }
            //            }
            insertmaster += "'" + fieldVialTextBox4.Text + "', ";
            insertmaster += "'" + hostTrapTextBox5.Text + "', ";
            insertmaster += "'" + comboBox2.Text + "', ";
            insertmaster += "'" + comboBox3.Text + "', ";
            insertmaster += "'" + localityTextBox8.Text + "', ";
            insertmaster += "'" + textBox9.Text + "', ";
            insertmaster += "'" + comboBox4.Text + "', ";
            insertmaster += "'" + comboBox5.Text + "', ";
            //            insertmaster += "'" + textBox12.Text + "', ";
            insertmaster += "'" + comboBox8.Text + "/" + comboBox9.Text + "/" + comboBox7.Text + "', ";
            insertmaster += "'" + textBox13.Text + "', ";
            insertmaster += "'" + comboBox6.Text + "', ";
            //            insertmaster += "'" + textBox15.Text + "'); ";
            if
                (
                comboBox10.Text == ""
                )
            {
                comboBox10.Text = "????";
            }
            if
                (
                comboBox11.Text == ""
                )
            {
                comboBox11.Text = "??";
            }  
            if
                (
                comboBox12.Text == ""
                )
            {
                comboBox12.Text = "??";
            }  
            if (
                comboBox10.Text.Trim() == "" &&
                comboBox11.Text.Trim() == "" &&
                comboBox12.Text.Trim() == ""
            ) {
                insertmaster += " NULL, ";
            } else {
                insertmaster += "'" + comboBox10.Text + "-" + comboBox11.Text + "-" + comboBox12.Text + "', ";
            }
            insertmaster += (checkBox1.Checked) ? "True" : "False";
            insertmaster += ")";
            try
            {
                if (this.thefile.dbo.State != ConnectionState.Open)
                {
                    this.thefile.dbo.Open();
                }
                OleDbCommand thebiginsert = new OleDbCommand(insertmaster, this.thefile.dbo);
                OleDbDataAdapter inserter = new OleDbDataAdapter();
                inserter.InsertCommand = thebiginsert;
                inserter.InsertCommand.ExecuteNonQuery();
                if (this.thefile.dbo.State != ConnectionState.Open)
                {
                    this.thefile.dbo.Open();
                }
                string newest = "SELECT TOP 1 vial FROM [COLLECTIONS] ORDER BY vial DESC";
                OleDbCommand selectnew = new OleDbCommand(newest, this.thefile.dbo);
                OleDbDataAdapter sn = new OleDbDataAdapter(selectnew);
                DataSet latest = new DataSet();
                sn.Fill(latest);
                textBox1.Text = latest.Tables[0].Rows[0][0].ToString();
                button3_Click(null, null);
                form4editenable(false);
                button4.Enabled = true;
                button5.Enabled = false;         
                button2.Enabled = false;
                textBox1_KeyUp(null, null, false);
                MessageBox.Show("New vial created: " + next + "\n\nPlease label the new vial now...","New vial");
            }
            catch (OleDbException err)
            {
                MessageBox.Show("Error inserting vial information. Please make sure all required fields are filled. Experiment, Capture->storage and collecetor must be present in the menus.\n\nSome fields may have been changed following this error- please check before resubmitting");
                MessageBox.Show(err.ToString());
//                MessageBox.Show("Error inserting Data! Did you check to make sure everything is filled out, valid, and in the right format? E.g. Is the date mm/dd/yyyy? Is the fungus a member of it's respective dropdown menu?");
//                MessageBox.Show(insertmaster);
                return;
            }
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form editDropDowns = new Form5(this.thefile, this, aa);
            editDropDowns.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to delete vial?", "Delete Vial",
    MessageBoxButtons.YesNo, MessageBoxIcon.Question,
    MessageBoxDefaultButton.Button2) != DialogResult.Yes)
            {
                return;
            }
            DataGridViewSelectedRowCollection editted;
            DataGridViewCellCollection col;
            if (this.dataGridView1.SelectedCells.Count == 0 && this.dataGridView1.SelectedRows.Count == 0)
            {
                return;
            }
            if (this.dataGridView1.SelectedCells.Count == 1 && this.dataGridView1.SelectedRows.Count < 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else if (this.dataGridView1.SelectedCells.Count > 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else
            {
                editted = this.dataGridView1.SelectedRows;
            }
            col = editted[0].Cells;
            if (col.Count > 0)
            {
                if (this.thefile.dbo.State != ConnectionState.Open)
                {
                    this.thefile.dbo.Open();
                }
                var todelete = col[0].Value.ToString();
                string deletemaster = "DELETE FROM [COLLECTIONS] WHERE VIAL = " + todelete;
                string deletemaster2 = "DELETE FROM [SPECIES_IN_COLLECTIONS] WHERE VIAL = " + todelete;
                OleDbCommand thebigdelete2 = new OleDbCommand(deletemaster2, this.thefile.dbo);
                OleDbDataAdapter deleter2 = new OleDbDataAdapter();
                deleter2.DeleteCommand = thebigdelete2;
                deleter2.DeleteCommand.ExecuteNonQuery();
                OleDbCommand thebigdelete = new OleDbCommand(deletemaster, this.thefile.dbo);
                OleDbDataAdapter deleter = new OleDbDataAdapter();
                deleter.DeleteCommand = thebigdelete;
                deleter.DeleteCommand.ExecuteNonQuery();
                textBox1_KeyUp(null, null, false);
            }
        }


        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (itsUnderControl) { return; }
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '-')
            {
                e.Handled = true;
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Back)
            {
                this.itsUnderControl = true;
                return;
            }
            if ((!e.Control && (e.KeyCode != Keys.A || e.KeyCode != Keys.C || e.KeyCode != Keys.X || e.KeyCode != Keys.V)))
            {
                e.Handled = true;
            }
            else
            {
                this.itsUnderControl = true;
            }
        }

        private void Form4_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.aa.IsDisposed)
            {
                Application.Exit();
            }
        }

        private void showSpecies_Click(object sender, EventArgs e)
        {
            if (this.aa.Visible)
            {
                this.aa.Focus();
            }
            if (this.aa.IsDisposed)
            {
                this.aa = new Form2(this.thefile, this.mutual);
                this.aa.vial = this;
                this.aa.initializeComponent();
                this.aa.Show();
            }
        }

        public Scolytos2.speciesLookUpForm speciesLookUpForm;

        private void speciesLookUp_Click(object sender, EventArgs e)
        {
            if (this.speciesLookUpForm == null || this.speciesLookUpForm.IsDisposed)
            {
                this.speciesLookUpForm = new Scolytos2.speciesLookUpForm(this, this.mutual, this.thefile);
                this.speciesLookUpForm.Show();
            }
            if (this.speciesLookUpForm != null && this.speciesLookUpForm.Visible)
            {
                this.speciesLookUpForm.Focus();
            }
       }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }
    }
}
