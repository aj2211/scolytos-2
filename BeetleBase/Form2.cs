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
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;

namespace BeetleBase
{
    public partial class Form2 : Form
    {
        public string dialogresult1;
        public mutual mutual;
        public string currentvial;
        public Form4 vial;
        public Form6 subform;
        public Thread thread;
        public DataSet dataset;
        public string currentrecord;
        public FileSystemWatcher watcher;
        public bool editting = false;
        public bool itsUnderControl;

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int AllocConsole();

        public void form2editenable(bool set)
        {
            this.textBox2.Enabled = set;
//            this.textBox3.Enabled = set;
            this.textBox4.Enabled = set;
            this.textBox5.Enabled = set;
            this.comboBox1.Enabled = set;
            this.comboBox2.Enabled = set;
            this.comboBox3.Enabled = set;
            //            this.textBox6.Enabled = set;
            this.textBox7.Enabled = set;
            this.textBox8.Enabled = set;
            this.textBox9.Enabled = set;
            this.textBox11.Enabled = set;
            this.textBox13.Enabled = set;
            this.textBox14.Enabled = set;
            this.button1.Enabled = !set;
            this.button2.Enabled = set;
            this.button3.Enabled = set;
            this.dataGridView1.Enabled = !set;
            this.textBox1.Enabled = !set;
            this.pinnedbox.Enabled = set;
            this.identifiercombo.Enabled = set;
            this.button7.Enabled = set;
        }

        public void initializeComponent()
        {
            InitializeComponent();
            this.button1.Focus();
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new System.Drawing.Point(0, 350);
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.DataSource = this.thefile.main.Tables[0];
            this.dataGridView1.Columns[13].Visible = false;
            try
            {
                string getidentifiers = "SELECT * FROM [COLLECTIONS- Drop Down Collector/Museum]";
                OleDbCommand getident = new OleDbCommand(getidentifiers, this.thefile.dbo);
                OleDbDataAdapter idadapt = new OleDbDataAdapter(getident);
                DataSet iddropdown = new DataSet();
                idadapt.Fill(iddropdown, "ID");
                int identifierscount = iddropdown.Tables["ID"].Rows.Count;
                for (int i = 0; i < identifierscount; i++)
                {
                    string jj = iddropdown.Tables["ID"].Rows[i][0].ToString();
                    identifiercombo.Items.Add(jj);
                }
            }
            catch (OleDbException err)
            {
                MessageBox.Show(err.ToString());
            }
            form2editenable(false);
        }
        public Form2(DB thefile, mutual mutual)
        {
            //            InitializeComponent();

            //           this.vial = form4;
            this.thefile = thefile;
            this.mutual = mutual;

            this.watcher = new FileSystemWatcher();
            this.watcher.SynchronizingObject = this;
            this.watcher.Path = Path.GetDirectoryName(thefile.watch);
            this.watcher.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
| NotifyFilters.FileName | NotifyFilters.DirectoryName;
            this.watcher.Filter = Path.GetFileName(thefile.watch);
            this.watcher.Changed += new FileSystemEventHandler(OnChange);
            this.watcher.EnableRaisingEvents = true;

        }

        public DB thefile;

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
            richTextBox2.Clear();
            richTextBox3.Clear();
            richTextBox4.Clear();
            string cmd = "SELECT b.[record], a.[vial], (c.[SpCode] & ' - ' & c.[Genus] & ' ' & c.[Species]) as [Species In Vial], b.[count], b.[male], b.[pair/family], b.[collector/museum], b.[SPECIES_note], b.[borrowed_count], b.[returned_date], b.[loaned_to], b.[loaned_number], b.[from plate], b.[SpCode], b.[PINNED] FROM (([COLLECTIONS] a LEFT OUTER JOIN [SPECIES_IN_COLLECTIONS] b ON a.[vial] = b.[vial]) LEFT OUTER JOIN [Species_table] c ON b.[SpCode] = c.[SpCode])";
            if (textBox1.Text.Trim() != "")
            {
                cmd += "WHERE a.[vial] = " + textBox1.Text;
                this.vial.textBox1.Text = textBox1.Text;
                if (!loop)
                {
                    this.vial.textBox1_KeyUp(null, null, false);
                }
            }
            OleDbCommand vialsearch = new OleDbCommand(cmd, this.thefile.dbo);
            OleDbDataAdapter vialadapter = new OleDbDataAdapter(vialsearch);
            DataSet vials = new DataSet();
            vialadapter.Fill(vials);
            this.dataGridView1.DataSource = vials.Tables[0];
            this.dataGridView1.Columns[13].Visible = false;
            vialadapter.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.editting = true;
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

            if (col[0].Value.ToString().Trim() == "")
            {
                MessageBox.Show("Use 'Add Species to Vial'");
                return;
            }

            form2editenable(true);
            button6.Enabled = false;
            button8.Enabled = false;
            this.currentvial = col[1].Value.ToString();
            this.currentrecord = col[0].Value.ToString();
            textBox2.Text = col[13].Value.ToString();
            textBox8.Text = col[1].Value.ToString();
            textBox9.Text = col[3].Value.ToString();

            textBox11.Text = col[4].Value.ToString();
//            textBox3.Text = col[6].Value.ToString();
            textBox4.Text = col[7].Value.ToString();
//            e.ToString();
            textBox5.Text = col[8].Value.ToString();

            char[] delimiterChars = { '/' };
            string[] words1 = col[10].Value.ToString().Split(delimiterChars);
            if (words1.Length > 1)
            {
                comboBox1.Text = words1[2];
                comboBox2.Text = words1[0];
                comboBox3.Text = words1[1];
            }
            textBox7.Text = col[10].Value.ToString();
            textBox14.Text = col[11].Value.ToString();
            textBox13.Text = col[12].Value.ToString();
            if (col[14].Value.ToString() == "True")
            {
                pinnedbox.Checked = true;
            }
            else
            {
                pinnedbox.Checked = false;
            }
            identifiercombo.Text = col[6].Value.ToString();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            Form popup = new Form3(this.thefile, this.mutual);
            popup.ShowDialog();
            if (this.mutual.result1 == null)
            {
                return;
            }
            this.textBox2.Text = this.mutual.result1;
            this.Cursor = Cursors.Default;
            DataSet speciesDisplay = new DataSet();
            string speciesDisplayCommand = "SELECT ([Genus] & ' ' & [species]) as [speciesDisplay] FROM [Species_table] WHERE [SpCode] = ";
            speciesDisplayCommand += this.mutual.result1;
            OleDbCommand speciesIdentifier = new OleDbCommand(speciesDisplayCommand, this.thefile.dbo);
            OleDbDataAdapter speciesId = new OleDbDataAdapter(speciesIdentifier);
            speciesId.Fill(speciesDisplay, "speciesId");
            try
            {
                this.richTextBox1.Text = speciesDisplay.Tables[0].Rows[0][0].ToString();
            }
            catch (NullReferenceException err)
            {
                MessageBox.Show(err.ToString());
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.editting = false;
            textBox2.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox11.Text = "";
//            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            textBox7.Text = "";
            textBox14.Text = "";
            textBox13.Text = "";
            identifiercombo.Text = "";
            pinnedbox.Checked = false;
            form2editenable(false);
            button6.Enabled = true;
            button8.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.editting = false;
            if (this.thefile.dbo.State != ConnectionState.Open)
            {
                this.thefile.dbo.Open();
            }
            string updatemaster = "UPDATE [SPECIES_IN_COLLECTIONS] SET ";
            if (textBox8.Text.Trim() == "")
            {

            }
            else
            {
                updatemaster += " [vial] = " + textBox8.Text;
            }
            if (textBox2.Text == "")
            {
                //                updatemaster += ", [SpCode] = null";
            }
            else
            {
                updatemaster += ", [SpCode] = " + textBox2.Text;
            }
            if (textBox9.Text == "")
            {
                //                updatemaster += ", [Count] = null";
            }
            else
            {
                updatemaster += ", [Count] = " + textBox9.Text;
            }
            if (textBox11.Text == "")
            {
                //                updatemaster += ", [male] = null";
            }
            else
            {
                updatemaster += ", [male] = " + textBox11.Text;
            }
//            updatemaster += ", [collector/museum] = '" + textBox3.Text + "'";
            updatemaster += ", [SPECIES_note] = '" + textBox4.Text + "'";
            if (textBox5.Text == "")
            {
                //                updatemaster += ", [borrowed_count] = null";
            }
            else
            {
                updatemaster += ", [borrowed_count] = " + textBox5.Text;
            }
            if
            (
                comboBox1.Text.Trim() != ""
            && comboBox2.Text.Trim() != ""
            && comboBox3.Text.Trim() != ""
            )
            {
                updatemaster += ", [returned_date] = '" + comboBox2.Text + "/" + comboBox3.Text + "/" + comboBox1.Text + "'";
                updatemaster += ", [loaned_to] = '" + textBox7.Text + "'";
            }
            if (textBox14.Text == "")
            {
                //                updatemaster += ", [loaned_number] = null";
            }
            else
            {
                updatemaster += ", [loaned_number] = " + textBox14.Text;
            }
            if (textBox13.Text == "")
            {

            }
            else
            {
                updatemaster += ", [from_plate] = " + textBox13.Text;
            }
            if (pinnedbox.Checked == true)
            {
                updatemaster += ", [PINNED] = " + "True";
            }
            else
            {
                updatemaster += ", [PINNED] = " + "False";
            }
            try
            {
                updatemaster += " WHERE record = " + currentrecord;

                string updatemaster2 = "UPDATE [SPECIES_IN_COLLECTIONS] SET [identifier] = '" + identifiercombo.Text + "' WHERE [record] = " + currentrecord;
                try
                {
                    OleDbCommand up = new OleDbCommand(updatemaster, this.thefile.dbo);
                    OleDbDataAdapter upd = new OleDbDataAdapter();
                    upd.UpdateCommand = up;
                    upd.UpdateCommand.ExecuteNonQuery();
                }
                catch (OleDbException)
                {
                    //                MessageBox.Show("Unable to write Species to Vial. Check to make sure information is valid!");
//                    MessageBox.Show(updatemaster);
                }
            }
            catch (OleDbException err) { MessageBox.Show(err.ToString()); return; }
            textBox1_KeyUp(null, null, false);
            button3_Click(null, null);
            form2editenable(false);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (!this.vial.Visible)
            {
                this.vial = new Form4(this.mutual, this.thefile);
                this.vial.aa = this;
                this.vial.Show();
            }
            else
            {
                this.vial.Focus();
            }
            this.vial.button4_Click(null, null);

        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (!this.vial.Visible)
            {
                this.vial = new Form4(this.mutual, this.thefile);
                this.vial.aa = this;
                this.vial.Show();
            }
            else
            {
                this.vial.Focus();
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
            this.vial.Controls["textbox1"].Text = col[1].Value.ToString();
            this.vial.textBox1_KeyUp2(null, null);
            this.vial.button1_Click(null, null);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            DataGridViewSelectedRowCollection editted;
            DataGridViewCellCollection col;
            if (this.dataGridView1.SelectedCells.Count == 0 && this.dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("Select Vial!");
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
            this.currentvial = col[1].Value.ToString();
            this.textBox8.Text = this.currentvial;
            this.currentrecord = col[0].Value.ToString();
            this.textBox8.ReadOnly = false;
            form2editenable(true);
            button8.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = true;
            //            button4.Enabled = false;
            //           button5.Enabled = false;
            button6.Enabled = false;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (this.thefile.dbo.State != ConnectionState.Open)
            {
                this.thefile.dbo.Open();
            }
            string insertmaster = "INSERT INTO [SPECIES_IN_COLLECTIONS] (";
            string insertmaster2 = ") VALUES (";
            insertmaster += "[vial]";
            insertmaster2 += textBox8.Text;
            insertmaster += ", [SpCode]";
            insertmaster2 += ", " + textBox2.Text;
            if (textBox9.Text != "" && textBox9.Text != " " && textBox9.Text != "  ")
            {
                insertmaster += ", [count]";
                insertmaster2 += ", " + textBox9.Text;
            }
            if (textBox11.Text != "" && textBox11.Text != " " && textBox11.Text != "  ")
            {
                insertmaster += ", [male]";
                insertmaster2 += ", " + textBox11.Text;
            }
            insertmaster += ", [SPECIES_note]";
            insertmaster2 += ", '" + textBox4.Text + "'";
            if (textBox5.Text != "" && textBox5.Text != " " && textBox5.Text != "  ")
            {
                insertmaster += ", [borrowed_count]";
                insertmaster2 += ", " + textBox5.Text;
            }
            if
                (
                comboBox1.Text != ""
                && comboBox1.Text != " "
                && comboBox1.Text != "  "
                && comboBox2.Text != ""
                && comboBox2.Text != " "
                && comboBox2.Text != "  "
                && comboBox3.Text != ""
                && comboBox3.Text != " "
                && comboBox3.Text != "  "
                )
            {
                insertmaster += ", [returned_date]";
                insertmaster2 += ", '" + comboBox2.Text + "/" + comboBox3.Text + "/" + comboBox1.Text + "'";
            }
            insertmaster += ", [loaned_to]";
            insertmaster2 += ", '" + textBox7.Text + "'";
            if (textBox14.Text != "" && textBox14.Text != " " && textBox14.Text != "  ")
            {
                insertmaster += ", [loaned_number]";
                insertmaster2 += ", " + textBox14.Text;
            }
            if (textBox13.Text != "" && textBox13.Text != " " && textBox13.Text != "  ")
            {
                insertmaster += ", [from plate]";
                insertmaster2 += ", " + textBox13.Text;
            }
            if (identifiercombo.Text.Trim() != "")
            {
                insertmaster += ", [identifier]";
                insertmaster2 += ", '" + identifiercombo.Text + "'";
            }
            if (pinnedbox.Checked)
            {
                insertmaster += ", [PINNED]";
                insertmaster2 += ", True";
            }
            else
            {
                insertmaster += ", [PINNED]";
                insertmaster2 += ", False";
            }
            insertmaster += insertmaster2 + ");";
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
//                string finalinsertcommand = "INSERT INTO [Identifiers] ([record], [identifier]) VALUES (" + lastid + ", '" + identifiercombo.Text + "')";
//                OleDbCommand finalinsert = new OleDbCommand(finalinsertcommand, this.thefile.dbo);
//                inserter.InsertCommand = finalinsert;
//                inserter.InsertCommand.ExecuteNonQuery();
                textBox1.Text = this.currentvial;
                this.vial.textBox1.Text = this.currentvial;
                button3_Click(null, null);
                form2editenable(false);
                //                button4.Enabled = true;
                //                button5.Enabled = true;
                button6.Enabled = true;
                button7.Enabled = false;
                this.textBox1_KeyUp(null, null, false);
            }
            catch (OleDbException err)
            {
                MessageBox.Show(err.ToString());
                //                MessageBox.Show("Error inserting Data! Did you check to make sure everything is filled out, valid, and in the right format? E.g. Is 'male' correctly formatted as a number? Does the editted vial exist?");
            }
            //            AllocConsole();
            //            Console.WriteLine(insertmaster);
            this.textBox8.ReadOnly = false;
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to delete species in vial?", "Delete Species In Vial",
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
                try
                {
                    var todelete = col[0].Value.ToString();
                    string deletemaster = "DELETE FROM [SPECIES_IN_COLLECTIONS] WHERE [record] = " + todelete;
                    OleDbCommand thebigdelete = new OleDbCommand(deletemaster, this.thefile.dbo);
                    OleDbDataAdapter deleter = new OleDbDataAdapter();
                    deleter.DeleteCommand = thebigdelete;
                    deleter.DeleteCommand.ExecuteNonQuery();
                    textBox1_KeyUp(null, null, false);
                }
                catch (OleDbException err)
                {
                    MessageBox.Show(err.ToString());
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (this.subform == null)
            {
                this.subform = new Form6();
                FormExtensions.RunInNewThread(subform, false, this.thread);
            }
            else if (!this.subform.Visible)
            {
                this.subform = new Form6();
                FormExtensions.RunInNewThread(subform, false, this.thread);
            }
        }

        private void dataGridView1_Click(object sender, EventArgs e)
        {
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
                if (editted[0].Cells[13].Value.ToString().Trim() != "")
                {
                    /*
                    try
                    {
                    */
                    string cmd = "SELECT [ImagePath] FROM [Images] WHERE [SpCode] = " + editted[0].Cells[13].Value.ToString() + " ORDER BY [SpCode] ASC";
//                    MessageBox.Show(cmd);
                    OleDbCommand test = new OleDbCommand(cmd, this.thefile.dbo);
                    OleDbDataAdapter begin = new OleDbDataAdapter(test);
                    this.dataset = new DataSet();
                    begin.Fill(this.dataset);
                    begin.Dispose();
                    if (this.dataset.Tables[0].Rows.Count > 0)
                    {
                        string a = this.dataset.Tables[0].Rows[0][0].ToString();
                        string str = this.thefile.root + @"\" + a;
                        //                        File.Open(@"C:\temp\Scol023.jpg", FileMode.Open);
                        // MessageBox.Show(a);
                        if (File.Exists(@str))
                        {
                            this.pictureBox1.Image = Image.FromFile(@str);
                            //                            MessageBox.Show("a");
                        }
                        else
                        {
                            this.pictureBox1.Image = null;
                        }
                        //                        MessageBox.Show(@str);
                        //                        }
                        /*
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Error!");
                        }
                        */
                    }
                    else
                    {
                        this.pictureBox1.Image = null;
                    }
                }
                else
                {
                    this.pictureBox1.Image = null;
                }
            }
            if (col[1].Value.ToString() != null)
            {
                DataSet vialInfo = new DataSet();
                string vialInfoQuery = "SELECT [locality], [county], [province], [Country], [host_or_trap] FROM [COLLECTIONS] WHERE [vial] = " + col[1].Value.ToString();
                OleDbCommand vialInfoQueryCommand = new OleDbCommand(vialInfoQuery, this.thefile.dbo);
                OleDbDataAdapter vialInfoAdapter = new OleDbDataAdapter(vialInfoQueryCommand);
                vialInfoAdapter.Fill(vialInfo);
                vialInfoAdapter.Dispose();
                DataRow vialInfoResults = vialInfo.Tables[0].Rows[0];
                richTextBox2.Clear();
                richTextBox3.Clear();
                richTextBox4.Clear();
                if (vialInfoResults[0].ToString().Trim() != "") { richTextBox2.Text = vialInfoResults[0].ToString(); }
                if (vialInfoResults[1].ToString().Trim() != "" && vialInfoResults[0].ToString().Trim() != "")
                {
                    richTextBox2.Text += ", " + vialInfoResults[1].ToString();
                }
                else if (vialInfoResults[1].ToString().Trim() != "")
                {
                    richTextBox2.Text = vialInfoResults[1].ToString();
                }
                if ((vialInfoResults[2].ToString().Trim() != "" || vialInfoResults[0].ToString().Trim() != "") && vialInfoResults[1].ToString().Trim() != "")
                {
                    richTextBox2.Text += ", " + vialInfoResults[2].ToString();
                }
                else if (vialInfoResults[2].ToString().Trim() != "")
                {
                    richTextBox2.Text = vialInfoResults[2].ToString();
                }
                if (vialInfoResults[3].ToString().Trim() != "") { richTextBox3.Text = vialInfoResults[3].ToString(); }
                if (vialInfoResults[4].ToString().Trim() != "") { richTextBox4.Text = vialInfoResults[4].ToString(); }
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        public formholder formhold;

        public void OnChange(object source, FileSystemEventArgs e)
        {
            if (this.textBox1.Text.ToString().Trim() != "")
            {
                this.textBox1_KeyUp(null, null, false);
            }
            if (this.editting)
            {
                this.button1_Click(null, null);
            }
            if (this.vial.textBox1.Text.ToString().Trim() != "")
            {
                this.vial.textBox1_KeyUp(null, null, true);
            }
            if (this.vial.editting)
            {
                this.vial.button1_Click(null, null);
            }
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (itsUnderControl) { return; }
            if (!char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {

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

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.vial.IsDisposed)
            {
                Application.Exit();
            }
        }

        private void showVials_Click(object sender, EventArgs e)
        {
            if (this.vial.Visible)
            {
                this.vial.Focus();
            }
            if (this.vial.IsDisposed)
            {
                this.vial = new BeetleBase.Form4(this.mutual, this.thefile);
                this.vial.aa = this;
                this.vial.initializeComponent();
                this.vial.Show();
            }
        }
    }

    internal static class FormExtensions
    {
        private static void ApplicationRunProc(object state)
        {
            Application.Run(state as Form);
        }

        public static void RunInNewThread(this Form form, bool isBackground, Thread thread)
        {
            if (form == null)
                throw new ArgumentNullException("form");
            if (form.IsHandleCreated)
                throw new InvalidOperationException("Form is already running.");
            thread = new Thread(ApplicationRunProc);
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = isBackground;
            thread.Start(form);
        }

    }

}
