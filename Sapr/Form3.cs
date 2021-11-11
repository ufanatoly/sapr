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

    public partial class Form3 : Form
    {
        public double q, h, l;
        public void searchKran(double q, double h, double l)
        {
            //q - требуемая грузоподъемность
            //h - требуемая высота подъема крюка
            //l - требуемый  вылет крюка
            //поиск крана по параметрам
            //заполнение данных в label

            for (int i = 0; i <= 2; i++)
            {
                if (q < cranes[i].maxweight)
                {
                    if (h < cranes[i].heightrise)
                    {
                        if (l < cranes[i].maxhook)
                        {
                            label13.Text = cranes[i].name;
                            label7.Text = Convert.ToString(cranes[i].minweight);
                            label8.Text = Convert.ToString(cranes[i].maxweight);
                            label9.Text = Convert.ToString(cranes[i].minhook);
                            label10.Text = Convert.ToString(cranes[i].maxhook);
                            label12.Text = Convert.ToString(cranes[i].heightrise);
                            break;
                        }
                    }
                }

            }
        }
        public Form3(double q, double h, double l)
        {
            InitializeComponent();
            this.q = q;
            this.h = h;
            this.l = l;
            
        }

        Crane[] cranes = new Crane[6];
        
    public class Crane
        {
            public String name;
            public double minweight;
            public double maxweight;
            public double minhook;
            public double maxhook;
            public double heightrise;

        }

        private void Form3_Load(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"D:\Study\САПР\ufanatoly.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;



            for (int i = 1; i <= 3; i++)
            {
                cranes[i - 1] = new Crane();
                for (int j = 1; j <= 6; j++)
                {
                    //new line

                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        if (j == 1) cranes[i - 1].name = xlRange.Cells[i, j].Value2.ToString();
                        if (j == 2) cranes[i - 1].minweight = Convert.ToDouble(xlRange.Cells[i, j].Value2);
                        if (j == 3) cranes[i - 1].maxweight = Convert.ToDouble(xlRange.Cells[i, j].Value2);
                        if (j == 4) cranes[i - 1].minhook = Convert.ToDouble(xlRange.Cells[i, j].Value2);
                        if (j == 5) cranes[i - 1].maxhook = Convert.ToDouble(xlRange.Cells[i, j].Value2);
                        if (j == 6) cranes[i - 1].heightrise = Convert.ToDouble(xlRange.Cells[i, j].Value2);
                    }


                    //add useful things here!   
                }
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

            searchKran(q, h, l);
        }
    }
    }

