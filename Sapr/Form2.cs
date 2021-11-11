using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Sapr
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        String str = "Pasha Zuev";
        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }


        private void расчитатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            double q = Convert.ToDouble(Ro.Text) + Convert.ToDouble(Ra.Text);
            double h = Convert.ToDouble(ho.Text) + Convert.ToDouble(hz.Text) + Convert.ToDouble(ha.Text) + Convert.ToDouble(hc.Text);
            //double l = (Convert.ToDouble(c.Text) + Convert.ToDouble(d.Text)) * (h + 2 + 2) / (2 + Convert.ToDouble(hc.Text)) + Convert.ToDouble(a.Text);
            double l = h - 10;
            Form3 newForm = new Form3(q, h, l);
            newForm.Show();
        }
    }
}
