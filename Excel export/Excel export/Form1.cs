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
        private int _million = (int)Math.Pow(10, 6);

        RealEstateEntities context = new RealEstateEntities();
        List<Flat> lakasok;

        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;

        string[] headers;

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
                FormatTable();

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
            headers = new string[]
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
                xlSheet.Cells[1, i + 1] = headers[i];

            object[,] values = new object[lakasok.Count, headers.Length];

            int counter = 0;
            int floorColum = 6;
            foreach (var lakas in lakasok)
            {
                values[counter, 0] = lakas.Code;
                values[counter, 1] = lakas.Vendor;
                values[counter, 2] = lakas.Side;
                values[counter, 3] = lakas.District;
                values[counter, 4] = lakas.Elevator ?
                    "Van"
                    : "Nincs";
                //if (lakas.Elevator)
                //    values[counter, 4] = "Van";
                //else
                //    values[counter, 4] = "Nincs";
                values[counter, 5] = lakas.NumberOfRooms;
                values[counter, 6] = lakas.FloorArea;
                values[counter, 7] = lakas.Price;                           //H2/G2*1000e
                values[counter, 8] = string.Format("={0}/{1}*{2}",          //= jel
                    "H" + (counter + 2).ToString(),
                    GetCell(counter + 2, floorColum + 1),
                    _million.ToString());
                counter++;
            }

            var range = xlSheet.get_Range(
                             GetCell(2, 1),
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

        private void FormatTable()
        {
            Excel.Range headerRange = xlSheet.get_Range(GetCell(1, 1), GetCell(1, headers.Length));
            headerRange.Font.Bold = true;
            headerRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            headerRange.EntireColumn.AutoFit();
            headerRange.RowHeight = 40;
            headerRange.Interior.Color = Color.LightBlue;
            headerRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);

            int lastRowID = xlSheet.UsedRange.Rows.Count;
            Excel.Range complateTableRange = xlSheet.get_Range(GetCell(1, 1), GetCell(lastRowID, headers.Length));
            //indulási pont     // utolso  sorszám
            complateTableRange.BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlThick);
        }

        private void Form1_Load(object sender, EventArgs e)
        { }
    }
}


