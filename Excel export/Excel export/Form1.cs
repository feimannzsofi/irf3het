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
using System.Reflection;

namespace Excel_export
{
    public partial class Form1 : Form
    {
        private int _millian = (int)Math.Pow(10, 6);

        RealEstateEntities context = new RealEstateEntities();
        List<Flat> lakasok;

        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;



        public Form1()
        {
            InitializeComponent();
            LoadData();
            CreateExcel();

            dataGridView1.DataSource = lakasok;
        }

        private void LoadData()
        {
            lakasok = context.Flats.ToList();
        }

        public void CreateExcel()
        {

            try
            {

                xlApp = new Excel.Application();


                xlWB = xlApp.Workbooks.Add(Missing.Value);


                xlSheet = xlWB.ActiveSheet;

                CreateTable();

                xlApp.Visible = true;
                xlApp.UserControl = true;

            }
            catch (Exception ex)
            {
                string errMsg = string.Format("Error: {0}\nLine: {1}", ex.Message, ex.Source);
                MessageBox.Show(errMsg, "Error");


                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }

        }

        private void CreateTable()
        {

            string[] headers = new string[] 
            {
                     "Kód",
                     "Eladó",
                     "Oldal",
                     "Kerület",
                     "Lift",
                     "Szobák száma",
                     "Alapterület (m2)",
                     "Ár (mFt)",
                     "Négyzetméter ár (Ft/m2)"
            };

            for (int i = 0; i < headers.Length; i++)
            {
                xlSheet.Cells[1, i = 1] = headers[i];
            }

            object[,] values = new object[lakasok.Count, headers.Length];

            int counter = 0;
            int floorColumn = 6;
            foreach (Flat f in lakasok)
            {
                values[counter, 0] = f.Code;
                values[counter, 1] = f.Vendor;
                values[counter, 2] = f.Side;
                values[counter, 3] = f.District;

                if (f.Elevator == true) {values[counter, 4] = true;}
                else{values[counter, 4] = false;}

                values[counter, 5] = f.NumberOfRooms;
                values[counter, floorColumn] = f.FloorArea;
                values[counter, 7] = f.Price;
                values[counter, 8] = "=" + GetCell(counter + 2, 8) + _millian + GetCell(counter + 2, 7);
                counter++;
            }

            var range = xlSheet.get_Range(
                GetCell(2,1),
                GetCell(1 + values.GetLength(0), values.GetLength(1)));
            range.Value2 = values;

        }
        private string GetCell(int x, int y)
        {
            string ExcelCoordinate = "";
            int dividend = y;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                ExcelCoordinate = Convert.ToChar(65 + modulo).ToString() + ExcelCoordinate;
                dividend = (int)((dividend - modulo) / 26);
            }
            ExcelCoordinate += x.ToString();

            return ExcelCoordinate;
        }
    }
}
