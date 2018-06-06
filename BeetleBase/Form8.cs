using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using BeetleBase;

namespace Scolytos2
{
    public partial class Form8 : Form
    {
        public Form2 form2;
        public Form4 form4;
        public Form8(BeetleBase.DB thefile, BeetleBase.mutual mutual)
        {
            InitializeComponent();
            //            this.SendToBack();
            this.Hide();
            this.form2 = new BeetleBase.Form2(thefile, mutual);
            this.form4 = new BeetleBase.Form4(mutual, thefile);
            this.form4.aa = this.form2;
            this.form2.vial = this.form4;
            this.form2.initializeComponent();
            this.form4.initializeComponent();
            this.form2.Show();
            this.form4.Show();
        }
    }
}
