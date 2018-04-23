using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CustomCalculationServiceExample
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {

        TestCustomCalculationService calculationService;

        public Form1() {
            InitializeComponent();
            spreadsheetControl1.BeginUpdate();
            spreadsheetControl1.Document.DocumentSettings.Calculation.Iterative = true;
            calculationService = new TestCustomCalculationService(loggingControl);
            spreadsheetControl1.AddService(typeof(ICustomCalculationService), calculationService);
            spreadsheetControl1.CustomDrawCell += spreadsheetControl1_CustomDrawCell;
            spreadsheetControl1.ActiveWorksheet.Cells["A3"].Formula = "A3*B3";
            spreadsheetControl1.ActiveWorksheet.Cells["B3"].Formula = "8*9";
            spreadsheetControl1.EndUpdate();
        }

        void spreadsheetControl1_CustomDrawCell(object sender, DevExpress.XtraSpreadsheet.CustomDrawCellEventArgs e)
        {
            CellKey cellkey = new CellKey((e.Cell.Worksheet.Index + 1), e.Cell.ColumnIndex, e.Cell.RowIndex);
            if (calculationService.CircularReferencedCells.Contains(cellkey))
            {
                e.Graphics.DrawRectangle(new Pen(Color.Red, 1), e.Bounds);
            }
        }

        private void clearButtonItem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            loggingControl.ClearNodes();
        }

        private void expandButtonItem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            loggingControl.ExpandAll();
        }

        private void collapseButtonItem_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            loggingControl.CollapseAll();
        }
    }
}
