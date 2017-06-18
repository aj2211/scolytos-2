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
using System.IO;

namespace Scolytos2
{
    public partial class speciesLookUpForm : Form
    {
        public BeetleBase.Form4 vial;
        public BeetleBase.mutual mutual;
        public BeetleBase.DB thefile;
        public DataSet pictureset;
        public int index = 1;
        public int indexmax = 1;
        public bool refresh = true;

        public speciesLookUpForm(BeetleBase.Form4 vial, BeetleBase.mutual mutual, BeetleBase.DB thefile)
        {
            InitializeComponent();
            this.vial = vial;
            this.mutual = mutual;
            this.nextbutton.Enabled = false;
            this.prevbutton.Enabled = false;

            this.thefile = thefile;
            OleDbCommand tribesearch = new OleDbCommand(@" SELECT DISTINCT Tribe FROM [Species_table];", this.thefile.dbo);
        OleDbDataAdapter tribeadapter = new OleDbDataAdapter(tribesearch);
        DataSet tribeset = new DataSet();
        tribeadapter.Fill(tribeset, "TRIBE");
            tribeadapter.SelectCommand.CommandText = @"SELECT DISTINCT [Genus] FROM [Species_table];";
            tribeadapter.Fill(tribeset, "GENUS");
            tribeadapter.SelectCommand.CommandText = @"SELECT DISTINCT [species] FROM [Species_table];";
            tribeadapter.Fill(tribeset, "SPECIES");
            tribeadapter.Dispose();
            //            MessageBox.Show(tribeset.Tables["SPECIES"].Columns[0].ColumnName);
            int triberowcount = tribeset.Tables["TRIBE"].Rows.Count;
        int genusrowcount = tribeset.Tables["GENUS"].Rows.Count;
        int speciesrowcount = tribeset.Tables["SPECIES"].Rows.Count;
            for (int tribei = 0; tribei<triberowcount; tribei++)
            {
                this.comboBox1.Items.Add(tribeset.Tables["TRIBE"].Rows[tribei][0]);
            }
            for (int genusi = 0; genusi<genusrowcount; genusi++)
            {
                this.comboBox2.Items.Add(tribeset.Tables["GENUS"].Rows[genusi][0]);
             }
            for (int speciesi = 0; speciesi<speciesrowcount; speciesi++)
            {
                this.comboBox3.Items.Add(tribeset.Tables["SPECIES"].Rows[speciesi][0]);
            }

        }

private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
{
    if (!this.refresh)
    {
        return;
    }
    if (this.comboBox1.Text != "" && this.comboBox1.Text != " " && this.comboBox1.Text != "  ")
    {
        this.Cursor = Cursors.WaitCursor;
        DataSet genusset = new DataSet();
        DataSet speciesset = new DataSet();
        string genusselection = @"SELECT DISTINCT Genus FROM [Species_table] WHERE [Tribe] ='" + comboBox1.Text + "';";
        OleDbCommand genussearch = new OleDbCommand(genusselection, this.thefile.dbo);
        OleDbDataAdapter speciesadapter = new OleDbDataAdapter(genussearch);
        speciesadapter.Fill(genusset, "GENUS");
        string speciesselection = @"SELECT DISTINCT Species FROM [Species_table] WHERE [Tribe] ='" + comboBox1.Text + "';";
        speciesadapter.SelectCommand.CommandText = speciesselection;
        speciesadapter.Fill(speciesset, "SPECIES");
        comboBox2.Items.Clear();
        comboBox3.Items.Clear();
        int genusrowcount = genusset.Tables["GENUS"].Rows.Count;
        for (int genusi = 0; genusi < genusrowcount; genusi++)
        {
            comboBox2.Items.Add(genusset.Tables["GENUS"].Rows[genusi][0]);
        }
        int speciesrowcount = speciesset.Tables["SPECIES"].Rows.Count;
        for (int speciesi = 0; speciesi < speciesrowcount; speciesi++)
        {
            comboBox3.Items.Add(speciesset.Tables["SPECIES"].Rows[speciesi][0]);
        }
        speciesadapter.Dispose();
        updateGridViewsForm3(null, null);
        this.Cursor = Cursors.Default;
    }
}

private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
{
    if (!this.refresh)
    {
        return;
    }
    if (this.comboBox2.Text != "" && this.comboBox2.Text != " " && this.comboBox2.Text != "  ")
    {
        this.Cursor = Cursors.WaitCursor;
        DataSet tribeset = new DataSet();
        DataSet speciesset = new DataSet();
        string tribeselection = @"SELECT DISTINCT Tribe FROM [Species_table] WHERE [Genus] ='" + comboBox2.Text + "';";
        OleDbCommand tribesearch = new OleDbCommand(tribeselection, this.thefile.dbo);
        OleDbDataAdapter speciesadapter = new OleDbDataAdapter(tribesearch);
        speciesadapter.Fill(tribeset, "TRIBE");
        string speciesselection = @"SELECT DISTINCT Species FROM [Species_table] WHERE [Genus] ='" + comboBox2.Text + "';";
        speciesadapter.SelectCommand.CommandText = speciesselection;
        speciesadapter.Fill(speciesset, "SPECIES");
        comboBox1.Items.Clear();
        comboBox3.Items.Clear();
        int genusrowcount = tribeset.Tables["TRIBE"].Rows.Count;
        for (int genusi = 0; genusi < genusrowcount; genusi++)
        {
            comboBox1.Items.Add(tribeset.Tables["TRIBE"].Rows[genusi][0]);
        }
        int speciesrowcount = speciesset.Tables["SPECIES"].Rows.Count;
        for (int speciesi = 0; speciesi < speciesrowcount; speciesi++)
        {
            comboBox3.Items.Add(speciesset.Tables["SPECIES"].Rows[speciesi][0]);
        }
        speciesadapter.Dispose();
        updateGridViewsForm3(null, null);
        this.Cursor = Cursors.Default;
    }
}

private void comboBox3_SelectionChangeCommitted(object sender, EventArgs e)
{
    if (!this.refresh)
    {
        return;
    }
    if (this.comboBox3.Text != "" && this.comboBox3.Text != " " && this.comboBox3.Text != "  ")
    {
        this.Cursor = Cursors.WaitCursor;
        DataSet tribeset = new DataSet();
        DataSet genusset = new DataSet();
        string tribeselection = @"SELECT DISTINCT Tribe FROM [Species_table] WHERE Species ='" + comboBox3.Text + "';";
        OleDbCommand tribesearch = new OleDbCommand(tribeselection, this.thefile.dbo);
        OleDbDataAdapter tribeadapter = new OleDbDataAdapter(tribesearch);
        tribeadapter.Fill(tribeset, "TRIBE");
        string genusselection = @"SELECT DISTINCT Genus FROM [Species_table] WHERE Species ='" + comboBox3.Text + "'";
        if (this.comboBox1.Text != "" && this.comboBox1.Text != " " && this.comboBox1.Text != "  ")
        {
            genusselection += " AND Tribe='" + comboBox1.Text + "'";
        }
        tribeadapter.SelectCommand.CommandText = genusselection;
        tribeadapter.Fill(genusset, "GENUS");
        comboBox1.Items.Clear();
        comboBox2.Items.Clear();
        int triberowcount = tribeset.Tables["TRIBE"].Rows.Count;
        for (int genusi = 0; genusi < triberowcount; genusi++)
        {
            comboBox1.Items.Add(tribeset.Tables["TRIBE"].Rows[genusi][0]);
        }
        int speciesrowcount = genusset.Tables["GENUS"].Rows.Count;
        for (int speciesi = 0; speciesi < speciesrowcount; speciesi++)
        {
            comboBox2.Items.Add(genusset.Tables["GENUS"].Rows[speciesi][0]);
        }
        tribeadapter.Dispose();
        updateGridViewsForm3(null, null);
        this.Cursor = Cursors.Default;
    }
}

private void comboBox3_Leave(object sender, EventArgs e)
{
    //            MessageBox.Show("left!");
}

private void updateGridViewsForm3(object sender, EventArgs e)
{
    if (!this.refresh)
    {
        return;
    }
    int genusselected = 0;
    int tribeselected = 0;
    string masterselect = "";
    if (comboBox1.Text != "" && comboBox1.Text != " " && comboBox1.Text != "  ")
    {
        masterselect += " WHERE c.[Tribe] = '" + comboBox1.Text + "'";
        tribeselected = 1;
    }
    if (comboBox2.Text != "" && comboBox2.Text != " " && comboBox2.Text != "  ")
    {
        if (tribeselected == 0)
        {
            masterselect += " WHERE";
        }
        else
        {
            masterselect += " AND";
        }
        masterselect += " c.[Genus] = '" + comboBox2.Text + "'";
        genusselected = 1;
    }
    if (comboBox3.Text != "" && comboBox3.Text != " " && comboBox3.Text != "  ")
    {
        if (tribeselected == 0 && genusselected == 0)
        {
            masterselect += " WHERE";
        }
        else
        {
            masterselect += " AND";
        }
        masterselect += " c.[species] ='" + comboBox3.Text + "'";
    }
    //            string speciescodeselect = "SELECT SpCode, species FROM [Species_table]" + masterselect;
    string relationselect = "SELECT b.[SpCode], a.[vial], b.[record], (c.[SpCode] & ' - ' & c.[Genus] & ' ' & c.[Species]) as [Species In Vial], b.[count], b.[male], b.[pair/family], b.[collector/museum], b.[SPECIES_note], b.[borrowed_count], b.[returned_date], b.[loaned_to], b.[loaned_number], b.[from plate], b.[PINNED], b.[identifier] FROM (([COLLECTIONS] a LEFT OUTER JOIN [SPECIES_IN_COLLECTIONS] b ON a.[vial] = b.[vial]) LEFT OUTER JOIN [Species_table] c ON b.[SpCode] = c.[SpCode]) " + masterselect;
    //            DataSet SpCode = new DataSet();
    DataSet Relation = new DataSet();
    OleDbCommand fetch = new OleDbCommand(relationselect, this.thefile.dbo);
    OleDbDataAdapter fetchadapter = new OleDbDataAdapter(fetch);
    fetchadapter.Fill(Relation, "RELATIONS");
            this.dataGridView1.DataSource = null;
    this.dataGridView1.DataSource = Relation.Tables["RELATIONS"];
            this.dataGridView1.Refresh();
    //            this.dataGridView2.DataSource = Relation.Tables["RELATIONS"];
}

private void dataGridView1_MouseUp(object sender, MouseEventArgs e)
{
    this.refresh = false;
    DataGridViewSelectedRowCollection editted;
    DataGridViewCellCollection col;
    if (this.dataGridView1.SelectedCells.Count == 1 && this.dataGridView1.SelectedRows.Count < 1)
    {
        int row = this.dataGridView1.SelectedCells[0].RowIndex;
        this.dataGridView1.ClearSelection();
        this.dataGridView1.Rows[row].Selected = true;
        editted = this.dataGridView1.SelectedRows;
        col = editted[0].Cells;
    }
    else if (this.dataGridView1.SelectedCells.Count > 1)
    {
        int row = this.dataGridView1.SelectedCells[0].RowIndex;
        this.dataGridView1.ClearSelection();
        this.dataGridView1.Rows[row].Selected = true;
        editted = this.dataGridView1.SelectedRows;
        col = editted[0].Cells;
    }
    else
    {
        editted = this.dataGridView1.SelectedRows;
    }
    if (editted.Count > 0)
    {
        col = editted[0].Cells;
                //            MessageBox.Show(col[0].Value.ToString());

                try
                {
                    string picturecommand = "SELECT b.[ImagePath] FROM([Species_table] a RIGHT JOIN[Images] b ON a.[SpCode] = b.[SpCode]) WHERE a.[SpCode] = " + col[0].Value.ToString();
                    this.pictureset = new DataSet();
                    OleDbCommand picturefetch = new OleDbCommand(picturecommand, this.thefile.dbo);
                    OleDbDataAdapter fetchadapter = new OleDbDataAdapter(picturefetch);
                    fetchadapter.Fill(this.pictureset, "PICTURESET");
                    this.index = 1;
                    this.indexmax = this.pictureset.Tables["PICTURESET"].Rows.Count;
                    this.picturelabel.Text = this.index + " of " + this.indexmax;
                    if (indexmax < 2)
                    {
                        this.prevbutton.Enabled = false;
                        this.nextbutton.Enabled = false;
                    }
                    else
                    {
                        this.prevbutton.Enabled = false;
                        this.nextbutton.Enabled = true;
                    }
                    if (this.pictureset.Tables["PICTURESET"].Rows.Count > 0)
                    {
                        string a = this.pictureset.Tables["PICTURESET"].Rows[0][0].ToString();
                        string str = this.thefile.root + @"\" + a;
                        if (File.Exists(@str))
                        {
                            this.pictureBox1.Image = Image.FromFile(@str);
                            this.picturelabel.Text = this.index + " of " + this.indexmax;
                        }
                        else
                        {
                            this.pictureBox1.Image = null;
                            this.picturelabel.Text = this.index + " of " + this.indexmax;
                        }
                    }
                }
                catch (NullReferenceException) { MessageBox.Show("e"); }
                catch (OleDbException f) { MessageBox.Show(f.ToString()); }
        try
        {
            this.comboBox1.Text = col[3].Value.ToString();
        }
        catch (NullReferenceException g) { MessageBox.Show("g"); }
        finally { }
        try
        {
            this.comboBox2.Text = col[2].Value.ToString();
        }
        catch (NullReferenceException h) { MessageBox.Show("h"); }
        finally { }
        try
        {
            this.comboBox3.Text = col[1].Value.ToString();
        }
        catch (NullReferenceException i) { MessageBox.Show("i"); }
        finally { }
        try
        {
            this.mutual.result1 = col[0].Value.ToString();
        }
        catch (NullReferenceException j) { MessageBox.Show("null"); }
        finally { }
    }
    else
    {
        this.picturelabel.Text = null;
        this.pictureBox1.Image = null;
    }
    this.refresh = true;
}


private void button1_Click(object sender, EventArgs e)
{
    dataGridView1_MouseUp(null, null);
    this.Close();
}

private void Form3_FormClosed(object sender, FormClosedEventArgs e)
{
    //            dataGridView1_MouseUp(null, null);
}

private void button2_Click(object sender, EventArgs e)
{
    this.Cursor = Cursors.WaitCursor;
    this.comboBox1.Text = "";
    this.comboBox2.Text = "";
    this.comboBox3.Text = "";
    this.comboBox1.Items.Clear();
    this.comboBox2.Items.Clear();
    this.comboBox3.Items.Clear();
    OleDbCommand tribesearch = new OleDbCommand(@" SELECT DISTINCT Tribe FROM [Species_table];", this.thefile.dbo);
    OleDbDataAdapter tribeadapter = new OleDbDataAdapter(tribesearch);
    DataSet tribeset = new DataSet();
    tribeadapter.Fill(tribeset, "TRIBE");
    tribeadapter.SelectCommand.CommandText = @"SELECT DISTINCT [Genus] FROM [Species_table];";
    tribeadapter.Fill(tribeset, "GENUS");
    tribeadapter.SelectCommand.CommandText = @"SELECT DISTINCT [species] FROM [Species_table];";
    tribeadapter.Fill(tribeset, "SPECIES");
    //            MessageBox.Show(tribeset.Tables["SPECIES"].Columns[0].ColumnName);
    int triberowcount = tribeset.Tables["TRIBE"].Rows.Count;
    int genusrowcount = tribeset.Tables["GENUS"].Rows.Count;
    int speciesrowcount = tribeset.Tables["SPECIES"].Rows.Count;
    this.comboBox1.BeginUpdate();
    object[] triber = new object[triberowcount];
    object[] speciesr = new object[speciesrowcount];
    object[] genusr = new object[genusrowcount];
    for (int tribei = 0; tribei < triberowcount; tribei++)
    {
        triber[tribei] = tribeset.Tables["TRIBE"].Rows[tribei][0];
    }
    this.comboBox1.Items.AddRange(triber);
    for (int genusi = 0; genusi < genusrowcount; genusi++)
    {
        genusr[genusi] = tribeset.Tables["GENUS"].Rows[genusi][0];
    }
    this.comboBox2.Items.AddRange(genusr);
    for (int speciesi = 0; speciesi < speciesrowcount; speciesi++)
    {
        speciesr[speciesi] = tribeset.Tables["SPECIES"].Rows[speciesi][0];
    }
    this.comboBox3.Items.AddRange(speciesr);
    this.comboBox1.EndUpdate();
    tribeadapter.Dispose();
    this.Cursor = Cursors.Default;

}

private void prevbutton_Click(object sender, EventArgs e)
{
    this.index--;
    if (this.index < 1)
    {
        this.index = 1;
    }
    if (this.index < 2)
    {
        this.prevbutton.Enabled = false;
    }
    if (this.index < this.indexmax)
    {
        this.nextbutton.Enabled = true;
    }
    if (index >= 1)
    {
        string a = this.pictureset.Tables["PICTURESET"].Rows[(index - 1)][0].ToString();
        string str = this.thefile.root + @"\" + a;
        if (File.Exists(@str))
        {
            this.pictureBox1.Image = Image.FromFile(@str);
        }
        else
        {
            this.pictureBox1.Image = null;
        }
    }
    this.picturelabel.Text = this.index + " of " + this.indexmax;
}

private void nextbutton_Click(object sender, EventArgs e)
{
    index++;
    if (this.index > this.indexmax)
    {
        this.index = this.indexmax;
    }
    if (this.index == this.indexmax || this.index > this.indexmax)
    {
        this.nextbutton.Enabled = false;
        this.prevbutton.Enabled = true;
    }
    if (this.index > 1)
    {
        this.prevbutton.Enabled = true;
    }
    if (this.index <= this.indexmax)
    {
        string a = this.pictureset.Tables["PICTURESET"].Rows[(index - 1)][0].ToString();
        string str = this.thefile.root + @"\" + a;
        if (File.Exists(@str))
        {
            this.pictureBox1.Image = Image.FromFile(@str);
        }
        else
        {
            this.pictureBox1.Image = null;
        }
    }
    this.picturelabel.Text = this.index + " of " + this.indexmax;
}

private void Form3_Load(object sender, EventArgs e)
{

}

private void comboBox1_KeyUp(object sender, KeyEventArgs e)
{
    if (e.KeyCode == Keys.Enter)
    {
        comboBox1_SelectionChangeCommitted(null, null);
    }
}

private void comboBox2_KeyUp(object sender, KeyEventArgs e)
{
    if (e.KeyCode == Keys.Enter)
    {
        comboBox2_SelectionChangeCommitted(null, null);
    }
}

private void comboBox3_KeyUp(object sender, KeyEventArgs e)
{
    if (e.KeyCode == Keys.Enter)
    {
        comboBox3_SelectionChangeCommitted(null, null);
    }
}
private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void speciesLookUpForm_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            this.refresh = false;
            DataGridViewSelectedRowCollection editted;
            DataGridViewCellCollection col;
            if (this.dataGridView1.SelectedCells.Count == 1 && this.dataGridView1.SelectedRows.Count < 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                this.dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else if (this.dataGridView1.SelectedCells.Count > 1)
            {
                int row = this.dataGridView1.SelectedCells[0].RowIndex;
                this.dataGridView1.ClearSelection();
                this.dataGridView1.Rows[row].Selected = true;
                editted = this.dataGridView1.SelectedRows;
                col = editted[0].Cells;
            }
            else
            {
                editted = this.dataGridView1.SelectedRows;
            }

            if (editted.Count > 0)
            {
                col = editted[0].Cells;
                string tribedungo = "SELECT b.[SpCode], a.[vial], b.[record], (c.[SpCode] & ' - ' & c.[Genus] & ' ' & c.[Species]) as [Species In Vial], b.[count], b.[male], b.[pair/family], b.[collector/museum], b.[SPECIES_note], b.[borrowed_count], b.[returned_date], b.[loaned_to], b.[loaned_number], b.[from plate], b.[PINNED], b.[identifier] FROM (([COLLECTIONS] a LEFT OUTER JOIN [SPECIES_IN_COLLECTIONS] b ON a.[vial] = b.[vial]) LEFT OUTER JOIN [Species_table] c ON b.[SpCode] = c.[SpCode]) WHERE a.[vial] = " + ((col[1].Value.ToString() != null) ? col[1].Value.ToString() : "NULL");
                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    this.comboBox1.Text = "";
                    this.comboBox2.Text = "";
                    this.comboBox3.Text = "";
                    OleDbCommand tribesearch = new OleDbCommand(tribedungo, this.thefile.dbo);
                    OleDbDataAdapter tribeadapter = new OleDbDataAdapter(tribesearch);
                    DataSet tribeset = new DataSet();
                    tribeadapter.Fill(tribeset, "RELATIONS");
                    this.dataGridView1.DataSource = null;
                    this.dataGridView1.DataSource = tribeset.Tables["RELATIONS"];
                    this.dataGridView1.Refresh();
                    tribeadapter.Dispose();
                    this.Cursor = Cursors.Default;
                }
                catch (OleDbException error)
                {
                    MessageBox.Show(error.ToString());
                    MessageBox.Show(tribedungo);
                    this.Cursor = Cursors.Default;
                    this.refresh = true;
                }
            }
            this.refresh = true;

        }
    }
}
