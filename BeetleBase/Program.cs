using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using System.Data;
using System.Runtime.InteropServices;
using Scolytos2;

namespace BeetleBase
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Previous startup = new Previous();
            mutual mutual = new mutual();
            string startupsave = startup.checkPreviousDB();
            string startupsave2 = startup.checkPreviousDB2();
            DB thefile = new DB();
            subMain(startupsave, startup, thefile, startupsave2, mutual);
        }

        static void subMain(string startupsave, Previous startup, DB thefile, string startupsave2, mutual mutual)
        {
            thefile.exitloop = true;
            
            Form a = new Form1(startupsave, startup, thefile, startupsave2);
            a.StartPosition = FormStartPosition.CenterScreen;
            Application.Run(a);
            if (thefile.OK == 1 && thefile.goahead)
            {
                //                formholder formhold = new BeetleBase.formholder();
                //                formhold.form2 = new Form2(thefile, mutual, formhold);
                //this.vial = new Form4(mutual, thefile, formhold.form2);
                //formhold.form2.StartPosition = FormStartPosition.Manual;
                //formhold.form2.Location = new System.Drawing.Point(0, 320);
                //formhold.form2.formhold = formhold;
                //                Application.Run(formhold.form2);
                //                Application.Run(this.vial);
//                Application.Run(new Scolytos2.Form8(thefile, mutual));
                Scolytos2.Form8 mainform = new Scolytos2.Form8(thefile, mutual);
                Application.Run();

            }
            else if (thefile.exitloop)
            {
                Application.Exit();
            }
            else
            {
                subMain(startupsave, startup, thefile, startupsave2, mutual);
            }
        }
    }

    public class Previous
    {
        public OpenFileDialog BrowseDB = new OpenFileDialog();
        public FolderBrowserDialog BrowseFolder = new FolderBrowserDialog();
        public int exists = 0;
        public int exists2 = 0;
        public Form b;
        public string checkPreviousDB()
        {
            string here = Directory.GetCurrentDirectory();
            string previouspath = here + @"\donotdelete.txt";
            if (File.Exists(previouspath))
            {
                this.exists = 1;
                //                string[] prev = File.ReadAllLines(previouspath);
                string prev = File.ReadLines(previouspath).Last();
                if (prev.Count() > 0)
                {
                    return prev;
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
            }
        }
        public string checkPreviousDB2()
        {
            string here2 = Directory.GetCurrentDirectory();
            string previouspath2 = here2 + @"\donotdelete2.txt";
            if (File.Exists(previouspath2))
            {
                this.exists = 1;
                string prev2 = File.ReadLines(previouspath2).Last();
                if (prev2.Count() > 0)
                {
                    return prev2;
                }
                else
                {
                    return "";
                }
            }
            else
            {
                return "";
            }
        }
    }

    public class DB
    {
        public OleDbConnection dbo;
        public DataSet main = new DataSet();
        public DataSet main2 = new DataSet();
        public int OK = 0;
        public string root;
        public bool goahead = false;
        public string watch;
        public bool exitloop = true;
    }

    public class mutual
    {
        public string result1;
    }

    public class formholder
    {
        public Form2 form2;
        public Form4 form4;
        public Form showboth()
        {
            this.form2.Show();
            this.form4.Show();
            return new Form();
        }
    }
}
