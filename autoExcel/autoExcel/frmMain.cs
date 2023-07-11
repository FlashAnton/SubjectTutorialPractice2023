using System;
using System.Windows.Forms;
using E = Microsoft.Office.Interop.Excel;

namespace autoExcel
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            E.Application oXls = new E.Application();

            E.Workbook oWb = oXls.Workbooks.Add();

            //Создаётся новый лист
            //oWb.Worksheets.Add();

            E.Range oRng = oXls.Range["D2", "F9"];

            oRng.Select();

            oRng.Borders.Weight = 2;



            oRng = oXls.Range["B2", "B9"];

            E.Range currCell;

            for (int i = 1; i <= oRng.Rows.Count; i++)
                for (int j = 1; j <= oRng.Columns.Count; j++)
                {
                    currCell = (E.Range)oRng.Cells[i, j];
                    currCell.Select();
                    switch (i)
                    {
                        case 1:
                            currCell.Borders[E.XlBordersIndex.xlDiagonalDown].Weight = 2;
                            break;
                        case 2:
                            currCell.Borders[E.XlBordersIndex.xlDiagonalUp].Weight = 2;
                            break;
                        case 3:
                            currCell.Borders[E.XlBordersIndex.xlEdgeTop].Weight = 2;
                            break;
                        case 4:
                            currCell.Borders[E.XlBordersIndex.xlEdgeBottom].Weight = 2;
                            break;
                        case 5:
                            currCell.Borders[E.XlBordersIndex.xlEdgeLeft].Weight = 2;
                            break;
                        case 6:
                            currCell.Borders[E.XlBordersIndex.xlEdgeRight].Weight = 2;
                            break;
                        case 7:
                            currCell.Borders[E.XlBordersIndex.xlInsideHorizontal].Weight = 2;
                            break;
                        case 8:
                            currCell.Borders[E.XlBordersIndex.xlInsideVertical].Weight = 2;
                            break;
                    }
                }
            oXls.Visible = true;
        }
    }
}
