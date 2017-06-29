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

namespace BeetleBase
{
    public partial class Form5 : Form
    {
        public Form5(DB thefile, Form4 a, Form2 aa)
        {
            InitializeComponent();
            this.treeView1.ExpandAll();
            this.thefile = thefile;
            this.ds = new DataSet();
            this.dsalt = new DataSet();
            this.a = a;
            this.aa = aa;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        internal static extern int AllocConsole();

        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            try
            {
                //            MessageBox.Show(e.Node.Name);
                this.ds.Clear();
                this.ds = new DataSet();
                dataGridView1.DataSource = null;
                dataGridView1.Columns.Clear();
                dataGridView1.Rows.Clear();
                dataGridView1.Refresh();
                dataGridView1.AllowUserToDeleteRows = true;
                this.currentname = e.Node.Name;
                if (e.Node.Name == "CaptureType")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Capture Type]";
                    this.currenttable = "[COLLECTIONS- Drop Down Capture Type]";
                    this.cn = "capture type";

                }
                else if (e.Node.Name == "Country")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Country]";
                    this.currenttable = "[COLLECTIONS- Drop Down Country]";
                    this.cn = "country";
                }
                else if (e.Node.Name == "Experiment")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Experiment]";
                    this.currenttable = "[COLLECTIONS- Drop Down Experiment]";
                    this.cn = "experiment";
                }
                else if (e.Node.Name == "Fungus")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Fungus]";
                    this.currenttable = "[COLLECTIONS- Drop Down Fungus]";
                    this.cn = "fungus";
                }
                else if (e.Node.Name == "Province")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Province]";
                    this.currenttable = "[COLLECTIONS- Drop Down Province]";
                    this.cn = "province";
                }
                else if (e.Node.Name == "Identifiers")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Identifiers]";
                    this.currenttable = "[COLLECTIONS- Drop Down Identifiers]";
                    this.cn = "identifier";
                }
                else if (e.Node.Name == "SpeciesTable")
                {
                    this.cmd = "SELECT [SpCode], [Tribe], [Genus], [species] FROM [Species_table]";
                    this.currenttable = "[Species_table]";
                    this.cn = "speciestable";
                }
                else if (e.Node.Name == "collectorMuseum")
                {
                    this.cmd = "SELECT * FROM [COLLECTIONS- Drop Down Collector/Museum]";
                    this.currenttable = "[COLLECTIONS- Drop Down Collector/Museum]";
                    this.cn = "collectormuseum";
                }
                else
                {
                    return;
                }
                OleDbCommand vialsearch = new OleDbCommand(this.cmd, this.thefile.dbo);
                this.vialadapter = new OleDbDataAdapter(vialsearch);
                //                DataSet vials = new DataSet();
                vialadapter.Fill(ds, "TABLE");
//                if (this.cn == "speciestable")
//                {
//                    DataGridViewTextBoxColumn key = new DataGridViewTextBoxColumn();
//                    key.Name = "key";
//                    DataGridViewTextBoxColumn spcode = new DataGridViewTextBoxColumn();
//                    spcode.Name = "SpCode";
//                    DataGridViewComboBoxColumn trb = new DataGridViewComboBoxColumn();
//                    trb.Name = "Tribe";
//                    DataGridViewComboBoxColumn genus = new DataGridViewComboBoxColumn();
//                    genus.Name = "Genus";
//                    DataGridViewComboBoxColumn spc = new DataGridViewComboBoxColumn();
//                    spc.Name = "Species";
//                    dataGridView1.Columns.AddRange(key, spcode, trb, genus, spc);
//                    int ii = 0;
//                    foreach (DataRow i in ds.Tables[0].Rows)
//                    {
//                        int a = dataGridView1.Rows.Add(i[0], i[1]);
//                        ((DataGridViewComboBoxCell)dataGridView1.Rows[a].Cells[2]).Items.Add(i[2]);
//                        ii++;
//                    }
//                }
//                else
//                {
                    this.dataGridView1.DataSource = ds.Tables[0];
                //                }
                if (this.cn == "speciestable")
                {
                    dataGridView1.Columns[0].Visible = false;
                    this.dataGridView1.AllowUserToDeleteRows = true;
                }
                else
                {
                    this.dataGridView1.AllowUserToDeleteRows = false;
                }
                this.additionalRows = new Dictionary<int, int>();
            }
            catch (OleDbException err)
            {
                MessageBox.Show(err.ToString());
            }
        }
       
        public DB thefile;
        public Form2 aa;
        public string cmd;
        public DataSet ds;
        public DataSet dsalt;
        public OleDbDataAdapter vialadapter;
        public string currenttable;
        public string currentname;
        public string cn;
        public Form4 a;
        public Dictionary<int, int> additionalRows = new Dictionary<int, int>();
        public Dictionary<int, string> additionalRowsStr = new Dictionary<int, string>();
        public List<int> subtractionalRows = new List<int>();

        private void Form5_FormClosing(object sender, FormClosingEventArgs e)
        {
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
            dropdown.SelectCommand.CommandText = "SELECT * FROM [COLLECTIONS- Drop Down Collector/Museum]";
            dropdown.Fill(dropdowns, "collectormuseum");
            a.comboBox1.Items.Clear();
            a.comboBox2.Items.Clear();
            a.comboBox3.Items.Clear();
            a.comboBox4.Items.Clear();
            a.comboBox5.Items.Clear();
            a.comboBox6.Items.Clear();
            aa.identifiercombo.Items.Clear();
            int capturetypecount = dropdowns.Tables["capturetype"].Rows.Count;
            int collectormuseumcount = dropdowns.Tables["collectormuseum"].Rows.Count;
            int countrycount = dropdowns.Tables["country"].Rows.Count;
            int experimentcount = dropdowns.Tables["experiment"].Rows.Count;
            int funguscount = dropdowns.Tables["fungus"].Rows.Count;
            int provincecount = dropdowns.Tables["province"].Rows.Count;
            for (int i = 0; i < capturetypecount; i++)
            {
                a.comboBox2.Items.Add(dropdowns.Tables["capturetype"].Rows[i][0]);
            }
            for (int i = 0; i < countrycount; i++)
            {
                a.comboBox5.Items.Add(dropdowns.Tables["country"].Rows[i][0]);
            }
            for (int i = 0; i < collectormuseumcount; i++)
            {
                a.comboBox6.Items.Add(dropdowns.Tables["collectormuseum"].Rows[i][0]);
                aa.identifiercombo.Items.Add(dropdowns.Tables["collectormuseum"].Rows[i][0]);
            }
            for (int i = 0; i < experimentcount; i++)
            {
                a.comboBox1.Items.Add(dropdowns.Tables["experiment"].Rows[i][0]);
            }
            for (int i = 0; i < funguscount; i++)
            {
                a.comboBox3.Items.Add(dropdowns.Tables["fungus"].Rows[i][0]);
            }
            for (int i = 0; i < provincecount; i++)
            {
                a.comboBox4.Items.Add(dropdowns.Tables["province"].Rows[i][0]);
            }
            dropdown.Dispose();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex > 0 && e.ColumnIndex > 0)
                {
                    dataGridView1.Rows[e.RowIndex].ReadOnly = true;
                    if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Trim() == "")
                    {
                        dataGridView1.Rows[e.RowIndex].ReadOnly = false;
                    }
                }
            }
            catch (IndexOutOfRangeException a)
            {
                //                MessageBox.Show("a");
            }
        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            if (this.thefile.dbo.State != ConnectionState.Open) this.thefile.dbo.Open(); 
            if ((MessageBox.Show("Are you sure you want to remove this item?", "Delete species",
    MessageBoxButtons.YesNo, MessageBoxIcon.Question,
    MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.Yes))
            {
                if (this.cn == "speciestable")
                {
                    var del = dataGridView1.Rows[e.Row.Index].Cells[0].Value.ToString().Trim();
                    string deleter;
                    string predeleter;
                    if (del == "")
                    {
                        deleter = "DELETE FROM [Species_table] WHERE [SpCode] Is Null";
                        predeleter = "DELETE FROM [SPECIES_IN_COLLECTIONS] WHERE [SpCode] Is Null";
                    }
                    else
                    {
                        deleter = "DELETE FROM [Species_table] WHERE [SpCode] = " + del;
                        predeleter = "DELETE FROM [SPECIES_IN_COLLECTIONS] WHERE [SpCode] = " + del;
                    }
                    OleDbCommand prefirst = new OleDbCommand(predeleter, this.thefile.dbo);
                    OleDbCommand first = new OleDbCommand(deleter, this.thefile.dbo);
                    prefirst.ExecuteNonQuery();
                    first.ExecuteNonQuery();
                }
                else if (this.cn == "capture type")
                {
                    var del = dataGridView1.Rows[e.Row.Index].Cells[0].Value.ToString().Trim();
                    string deleter;
                    if (del == "")
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [capture type] Is Null";
                    }
                    else
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [capture type] = '" + del + "'";
                    }
                    OleDbCommand first = new OleDbCommand(deleter, this.thefile.dbo);
                    first.ExecuteNonQuery();
                }
                else if (this.cn == "country")
                {
                    var del = dataGridView1.Rows[e.Row.Index].Cells[0].Value.ToString().Trim();
                    string deleter;
                    if (del == "")
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [country] Is Null";
                    }
                    else
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [country] = '" + del + "'";
                    }
                    OleDbCommand first = new OleDbCommand(deleter, this.thefile.dbo);
                    first.ExecuteNonQuery();
                }
                else if (this.cn == "experiment")
                {
                    var del = dataGridView1.Rows[e.Row.Index].Cells[0].Value.ToString().Trim();
                    string deleter;
                    if (del == "")
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [experiment] Is Null";
                    }
                    else
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [experiment] = '" + del + "'";
                    }
                    OleDbCommand first = new OleDbCommand(deleter, this.thefile.dbo);
                    first.ExecuteNonQuery();
                }
                else if (this.cn == "fungus")
                {
                    var del = dataGridView1.Rows[e.Row.Index].Cells[0].Value.ToString().Trim();
                    string deleter;
                    if (del == "")
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [fungus] Is Null";
                    }
                    else
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [fungus] = '" + del + "'";
                    }
                    OleDbCommand first = new OleDbCommand(deleter, this.thefile.dbo);
                    first.ExecuteNonQuery();
                }
                else if (this.cn == "province")
                {
                    var del = dataGridView1.Rows[e.Row.Index].Cells[0].Value.ToString().Trim();
                    string deleter;
                    if (del == "")
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [province] Is Null";
                    }
                    else
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [province] = '" + del + "'";
                    }
                    OleDbCommand first = new OleDbCommand(deleter, this.thefile.dbo);
                    first.ExecuteNonQuery();
                }
                else if (this.cn == "collectormuseum")
                {
                    var del = dataGridView1.Rows[e.Row.Index].Cells[0].Value.ToString().Trim();
                    string deleter;
                    if (del == "")
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [collector/museum] Is Null";
                    }
                    else
                    {
                        deleter = "DELETE FROM " + this.currenttable + " WHERE [collector/museum] = '" + del + "'";
                    }
                    OleDbCommand first = new OleDbCommand(deleter, this.thefile.dbo);
                    first.ExecuteNonQuery();
                }
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            if (this.cn == "speciestable")
            {
                DataSet top = new DataSet();
                OleDbCommand first = new OleDbCommand("SELECT TOP 1 [SpCode] FROM [Species_table] ORDER BY [SpCode] DESC", this.thefile.dbo);
                OleDbDataAdapter topdown = new OleDbDataAdapter(first);
                topdown.Fill(top, "FIRST");
                this.additionalRows.Add((e.Row.Index - 1), (Int32.Parse(top.Tables["FIRST"].Rows[0][0].ToString()) + 1));
                dataGridView1.Rows[(e.Row.Index - 1)].Cells[0].Value = (Int32.Parse(top.Tables["FIRST"].Rows[0][0].ToString()) + 1);
            }
            else
            {
                this.additionalRows.Add((e.Row.Index - 1), 0);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
//            string display = "";
//            foreach (var a in this.additionalRows)
//           {
//                display += a.Key.ToString() + " => " + a.Value.ToString() + ", ";
//            }
//            MessageBox.Show(display);
            foreach (var a in this.additionalRows)
            {
                if (this.thefile.dbo.State != ConnectionState.Open) this.thefile.dbo.Open();
                try
                {
                    string inserter = "";
                    if (this.cn == "speciestable")
                    {
                        inserter = "INSERT INTO [Species_table] ([Tribe], [Genus], [species]) VALUES ('"
                        + dataGridView1.Rows[a.Key].Cells[1].Value.ToString() + "', '"
                        + dataGridView1.Rows[a.Key].Cells[2].Value.ToString() + "', '"
                        + dataGridView1.Rows[a.Key].Cells[3].Value.ToString() + "')";
                    }
                    else if (this.cn == "capture type")
                    {
                        inserter = "INSERT INTO "
                        + this.currenttable
                        + " ([capture type]) VALUES ('"
                        + dataGridView1.Rows[a.Key].Cells[0].Value.ToString()
                        + "')";
                    }
                    else if (this.cn == "country")
                    {
                        inserter = "INSERT INTO "
                        + this.currenttable
                        + " ([country]) VALUES ('"
                        + dataGridView1.Rows[a.Key].Cells[0].Value.ToString()
                        + "')";
                    }
                    else if (this.cn == "experiment")
                    {
                        inserter = "INSERT INTO "
                        + this.currenttable
                        + " ([experiment]) VALUES ('"
                        + dataGridView1.Rows[a.Key].Cells[0].Value.ToString()
                        + "')";
                    }
                    else if (this.cn == "fungus")
                    {
                        inserter = "INSERT INTO "
                        + this.currenttable
                        + " ([fungus]) VALUES ('"
                        + dataGridView1.Rows[a.Key].Cells[0].Value.ToString()
                        + "')";
                    }
                    else if (this.cn == "province")
                    {
                        inserter = "INSERT INTO "
                        + this.currenttable
                        + " ([province]) VALUES ('"
                        + dataGridView1.Rows[a.Key].Cells[0].Value.ToString()
                        + "')";
                    }
                    else if (this.cn == "collectormuseum")
                    {
                        inserter = "INSERT INTO "
                        + this.currenttable
                        + " ([collector/museum]) VALUES ('"
                        + dataGridView1.Rows[a.Key].Cells[0].Value.ToString()
                        + "')";
                    }
  //                  OleDbCommand precmd = new OleDbCommand("INSERT INTO [SPECIES_IN_COLLECTIONS] ([SpCode]) ", this.thefile.dbo);
                    OleDbCommand cmd = new OleDbCommand(inserter, this.thefile.dbo);
                    cmd.ExecuteNonQuery();
                }
                catch (OleDbException err)
                {

                }
            }
       }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }
    }
}
