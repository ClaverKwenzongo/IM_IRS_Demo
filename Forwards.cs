namespace IM_IRS_Demo
{
    public class Forwards
    {
        Discounts get_df = new Discounts();

        public double getFwds(int row , int col, int col_inc, int start_col)
        {
            double FwdRate = 0;
            if (col == start_col)
            {
                FwdRate = Globals.Sheet3.Cells[row,col].Value;
            }
            else
            {
                double df_1 = get_df.getDFs(row,col - col_inc);
                double df_2 = get_df.getDFs(row,col);
                double days_diff = Globals.Sheet9.Cells[row,col].Value - Globals.Sheet9.Cells[row, col - col_inc].Value;

                FwdRate = (df_1/df_2 - 1)*(365/days_diff);
            }

            return FwdRate;
        }
    }
}
