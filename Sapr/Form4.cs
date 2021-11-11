using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sapr
{
    public partial class Form4 : Form
    {
        
        public Form4()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            String[] columnsName = new String[9];
            String[] linkName = new String[3];

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Study\САПР\columns.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            Columns[] columns = new Columns[9];

            for (int i = 0; i < columns.Length; i++)
            {
                columns[i] = new Columns("asdasd", "asd", "asd", "sadsa");
            }


            for (int i = 1; i <= 9; i++)
            {
                columns[i - 1].name = Convert.ToString(xlRange.Cells[i, 1].Value2);
                columns[i - 1].section = Convert.ToString(xlRange.Cells[i, 2].Value2);
                columns[i - 1].weight = Convert.ToString(xlRange.Cells[i, 3].Value2);
                columns[i - 1].volume = Convert.ToString(xlRange.Cells[i, 4].Value2);
                for (int j = 1; j <= 4; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");



                    //add useful things here!   
                }
            }

            for (int i = 0; i < columns.Length; i++)
            {
                Console.Write(columns[i].name + " " + columns[i].section + " " + columns[i].weight + " " + columns[i].volume + "\n\n");
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad


            //close and release
            xlWorkbook.Close();

            //quit and release
            xlApp.Quit();

            for (int i = 0; i <= 8; i++)
            {
                columnsName[i] = columns[i].name;
            }

            comboBox1.Items.AddRange(columnsName);

            Dictionary<string, double> connections = new Dictionary<string, double>();
            connections.Add("В виде стержня l=5980", 0.42);
            connections.Add("В виде креста", 1.00);
            connections.Add("В виде фермы", 0.7);

            int pavel = 0;
            foreach (KeyValuePair<string, double> kvp in connections)
            {
                linkName[pavel] = kvp.Key;
                pavel++;
            }
            comboBox2.Items.AddRange(linkName);
        }

       

        public class Columns
        {
            public String name;
            public String section;
            public String weight;
            public String volume;

            public Columns()
            {

            }

            public Columns(String name, String section, String weight, String volume)
            {
                this.name = name;
                this.section = section;
                this.weight = weight;
                this.volume = volume;
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void расчитатьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Form5 newForm = new Form5(weight.Text, length.Text, prolet.Text, 1, 1);
            Form5 newForm = new Form5(width.Text, length.Text, prolet.Text, comboBox1.SelectedIndex, comboBox2.SelectedIndex);
            newForm.Show();
        }
    }
}
