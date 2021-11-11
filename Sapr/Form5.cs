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
    public partial class Form5 : Form
    {


        double P;
        double width, length, prolet;
        int columnIndex, svazIndex;
        public Form5(String width, String length, String prolet, int columnIndex, int svazIndex)
        {
            InitializeComponent();

            this.width = Convert.ToDouble(width);
            this.length = Convert.ToDouble(length);
            this.prolet = Convert.ToDouble(prolet);
            this.columnIndex = columnIndex;
            this.svazIndex = svazIndex;
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


        private void Form5_Load(object sender, EventArgs e)
        {
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

            Dictionary<string, double> connections = new Dictionary<string, double>();
            connections.Add("В виде стержня l=5980", 0.42);
            connections.Add("В виде креста", 1.00);
            connections.Add("В виде фермы", 0.7);

            int people = 5;
            P = 2 * width + 2 * length;
            int countColumn = Convert.ToInt32(P / (prolet + Convert.ToDouble(columns[columnIndex].section) / 1000));
            Console.WriteLine(countColumn);
            double weightColumn = Convert.ToDouble(columns[columnIndex].weight);
            double peoplehours;
            double montHours;
            double machineHours;
            if (weightColumn > 10)
            {
                montHours = 9;
                machineHours = 1.8;
            }
            else if (weightColumn > 8)
            {
                montHours = 7;
                machineHours = 1.4;
            }
            else if (weightColumn > 6)
            {
                montHours = 6;
                machineHours = 1.2;
            }
            else if (weightColumn > 4)
            {
                montHours = 5.5;
                machineHours = 1.1;
            }
            else if (weightColumn > 3)
            {
                montHours = 4.3;
                machineHours = 0.86;
            }
            else if (weightColumn > 2)
            {
                montHours = 3.7;
                machineHours = 0.74;
            }
            else if (weightColumn > 1)
            {
                montHours = 3.1;
                machineHours = 0.61;
            }
            else
            {
                montHours = 2.2;
                machineHours = 0.55;
            }
            int smena = 20;
            Console.WriteLine(montHours + machineHours);
            double fullHours = countColumn * (montHours + machineHours);
            double HoursPerMont = countColumn * montHours / (people - 1);
            double HoursPerMach = countColumn * machineHours;
            double Money = HoursPerMont / 160 * 62221 * 4 + HoursPerMach / 160 * 49564;
            double HoursPerTonn = fullHours / (countColumn * Convert.ToDouble(columns[columnIndex].weight));
            double tonnperpeople = 8 / (HoursPerTonn / people);
            double MoneyPerTonn = Money / (countColumn * Convert.ToDouble(columns[columnIndex].weight));
            text1.Text = Convert.ToString(Math.Round(fullHours, 3));
            text2.Text = Convert.ToString(Math.Round(HoursPerTonn, 3));
            text3.Text = Convert.ToString(Math.Round(tonnperpeople, 3));
            text4.Text = Convert.ToString(Math.Round(Money, 3));
            text5.Text = Convert.ToString(Math.Round(MoneyPerTonn, 3));
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click_1(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
