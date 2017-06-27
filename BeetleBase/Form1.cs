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

namespace BeetleBase
{
    public partial class Form1 : Form
    {
        public string startupsave;
        public string startupsave2;
        public Previous startup;
        public DB thefile;
        public Form1(string startupsave, Previous startup, DB thefile, string startupsave2)
        {
            InitializeComponent();
            this.startupsave = startupsave;
            this.startupsave2 = startupsave2;
            this.textBox1.Text = startupsave;
            this.textBox2.Text = startupsave2;
            this.startup = startup;
            this.thefile = thefile;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult find = this.startup.BrowseDB.ShowDialog();
            if (find == DialogResult.OK)
            {
                this.textBox1.Text = this.startup.BrowseDB.FileName;
            }
            else
            {
                MessageBox.Show("Invalid selection!");
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.thefile.exitloop = false;
            this.Cursor = Cursors.WaitCursor;
            this.Text = "Scolytos 2 (Loading...)";
//            string preconnect = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=";
            string preconnect = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=";
            string connect = preconnect + this.textBox1.Text;
            this.thefile.root = this.textBox2.Text;
            this.thefile.watch = this.textBox1.Text;
            this.thefile.dbo = new OleDbConnection(connect);
            try
            {
                OleDbCommand test = new OleDbCommand("SELECT TOP 10 b.[record], a.[vial], (c.[SpCode] & ' - ' & c.[Genus] & ' ' & c.[Species]) as [Species In Vial], b.[count], b.[male], b.[pair/family], b.[collector/museum], b.[SPECIES_note], b.[borrowed_count], b.[returned_date], b.[loaned_to], b.[loaned_number], b.[from plate], b.[SpCode], b.[PINNED], b.[identifier] FROM (([COLLECTIONS] a LEFT OUTER JOIN [SPECIES_IN_COLLECTIONS] b ON a.[vial] = b.[vial]) LEFT OUTER JOIN [Species_table] c ON b.[SpCode] = c.[SpCode])", this.thefile.dbo);

                OleDbDataAdapter begin = new OleDbDataAdapter(test);
                begin.Fill(this.thefile.main);
                begin.Dispose();
                OleDbCommand test2 = new OleDbCommand("SELECT TOP 10 COLLECTIONS.vial, COLLECTIONS.experiment, COLLECTIONS.field_vial, COLLECTIONS.host_or_trap, COLLECTIONS.[capture->storage], COLLECTIONS.fungus, COLLECTIONS.Country, COLLECTIONS.province, COLLECTIONS.county, COLLECTIONS.locality, COLLECTIONS.date, COLLECTIONS.VIAL_note, COLLECTIONS.[collector/museum], COLLECTIONS.[pair/family], COLLECTIONS.date_collected FROM [COLLECTIONS]", this.thefile.dbo);
                OleDbDataAdapter begin2 = new OleDbDataAdapter(test2);
                begin2.Fill(this.thefile.main2);
                begin2.Dispose();
                thefile.goahead = true;
            }
            catch (Exception error)
            {
                MessageBox.Show(error.ToString());
                MessageBox.Show("Cannot create connection to file used!");
//                button1_Click(null, null);
                return;
            }
            finally
            {
//                thefile.goahead = true;
                try
                {
                    string path = Directory.GetCurrentDirectory() + @"\donotdelete.txt";
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine(textBox1.Text);
                    }
                    string path2 = Directory.GetCurrentDirectory() + @"\donotdelete2.txt";
                    using (StreamWriter sw2 = File.AppendText(path2))
                    {
                        sw2.WriteLine(textBox2.Text);
                    }
                }
                catch (IOException err)
                {
                    MessageBox.Show(err.ToString());
                }
                this.thefile.OK = 1;
                this.Close();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DialogResult find = this.startup.BrowseFolder.ShowDialog();
            if (find == DialogResult.OK)
            {
                this.textBox2.Text = this.startup.BrowseFolder.SelectedPath.ToString();
            }
            else
            {
                MessageBox.Show("Invalid selection!");
                return;
            }
        }
    }
}
