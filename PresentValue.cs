namespace IM_IRS_Demo
{
    public class PresentValue
    {
        Discounts getDiscount = new Discounts();
        Forwards getForwards = new Forwards();
        public double PV_float( double spread, double notional, int row, int col_start, int col_inc, int col_end)
        {
            double pv_float = 0;

            int i = col_start;
            while (i < col_end + 1)
            {
                double tau = Globals.Sheet9.Cells[row,i].Value;
                double df = getDiscount.getDFs(row, i);
                double fwdRate = getForwards.getFwds(row, i, col_inc, col_start);
                double fwd_plus_spread = fwdRate + spread*(0.01/100); //The spread is multiplied by 0.01% to make a conversion from basis points.

                pv_float += notional * fwd_plus_spread*(tau/365) * df;
                i += col_inc;
            }

            return pv_float;
        }

        public double PV_fixed( double notional, int row, int col_start, int col_inc, int col_end)
        {
            double pv_fixed = 0;

            int i = col_start;
            while ( i < col_end + 1)
            {
                double tau = Globals.Sheet9.Cells[row, i].Value;
                double df = getDiscount.getDFs(row, i);
                pv_fixed += notional*(tau/365)* df;
                i += col_inc;
            }
            return pv_fixed;
        }
    }
}
