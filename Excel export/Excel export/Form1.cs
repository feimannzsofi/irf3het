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
        List<Flat> Flats;

        RealEstateEntities context = new RealEstateEntities();

        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;

        

        public Form1()
        {
            InitializeComponent();
            LoadData();
        }

        public void LoadData()
        {
            List<Flat> Flats = context.Flats.ToList();
        }
    }
}
