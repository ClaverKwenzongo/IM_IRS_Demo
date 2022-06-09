using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace IM_IRS_Demo
{
    public partial class Ribbon1
    {
        Discounts getDiscount = new Discounts();
        Forwards getForwards = new Forwards();
        PresentValue PV = new PresentValue();
        public int countRows()
        {
            //Count how many rows of data they are:
            int count = 1;
            int row = 3; //Start from the third column assigned as the starting column after data preparation.

            while (string.IsNullOrWhiteSpace(Globals.Sheet3.Cells[row, 1].Value?.ToString()) == false)
            {
                count++;
                row++;
            }

            return count;
        }

        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Valuate_Click(object sender, RibbonControlEventArgs e)
        {
            int rowCount = countRows();
        }

        private void findDiscountFs_Click(object sender, RibbonControlEventArgs e)
        {
            int rows = countRows();
            for (int i = 3; i < rows+2; i++)
            {
                int j = 2;
                while (string.IsNullOrWhiteSpace(Globals.Sheet3.Cells[i, j].Value?.ToString()) == false)
                {
                    Globals.Sheet4.Cells[i, j].Value = getDiscount.getDFs(i,j);
                    j++;
                }
            }
        }

        private void findFowardRates_Click(object sender, RibbonControlEventArgs e)
        {
            int rows = countRows();
            for (int i = 3; i < rows + 2; i++)
            {
                int j = 2;
                while (string.IsNullOrWhiteSpace(Globals.Sheet3.Cells[i, j].Value?.ToString()) == false)
                {
                    Globals.Sheet5.Cells[i, j].Value = getForwards.getFwds(i,j);
                    j++;
                }
            }
        }
    }
}
