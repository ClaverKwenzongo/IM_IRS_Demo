﻿namespace IM_IRS_Demo
{
    public class Forwards
    {
        Discounts get_df = new Discounts();

        public double getFwds(int row , int col)
        {
            double FwdRate = 0;
            if (col == 2)
            {
                FwdRate = Globals.Sheet3.Cells[row,col].Value;
            }
            else
            {
                double df_1 = Globals.Sheet4.Cells[row,col-1].Value;
                double df_2 = Globals.Sheet4.Cells[row, col].Value;
                double days_diff = Globals.Sheet9.Cells[row,col].Value - Globals.Sheet9.Cells[row, col - 1].Value;

                FwdRate = (df_1/df_2 - 1)*(365/days_diff);
            }

            return FwdRate;
        }
    }
}
