namespace IM_IRS_Demo
{
    public class PresentValue
    {
        Discounts getDiscount = new Discounts();
        Forwards getForwards = new Forwards();
        public double PV_float( double spread, double notional, int row, int col_start, int col_inc, int col_end)
        {
            double pv_float = 0;
            
            for (int i = col_start; i <= col_end; i += col_inc)
            {
                double tau = Globals.Sheet9.Cells[row,i].Value;
                double df = getDiscount.getDFs(row, i);
                double fwdRate = getForwards.getFwds(row, i);

                pv_float += notional * (fwdRate + spread)*(tau/365) * df;
            }

            return pv_float;
        }

        public double PV_fixed( double notional, int row, int col_start, int col_inc, int col_end)
        {
            double pv_fixed = 0;

            for (int i = col_start; i <= col_end; i += col_inc)
            {
                double tau = Globals.Sheet9.Cells[row, i].Value;
                double df = getDiscount.getDFs(row, i);
                pv_fixed += notional*(tau/365) * df;
            }
            return pv_fixed;
        }
    }
}
