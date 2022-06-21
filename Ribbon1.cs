using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace IM_IRS_Demo
{
    public partial class Ribbon1
    {
        Discounts getDiscount = new Discounts();
        Forwards getForwards = new Forwards();
        PresentValue PV = new PresentValue();
        getPercentile percentile_func = new getPercentile();
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

        //Calculate the total amount exposed by summing the notional of every IRS:
        public double total_exposure()
        {
            int col = 3;
            double total = 0;
            while(string.IsNullOrWhiteSpace(Globals.Sheet1.Cells[8,col].Value?.ToString()) == false)
            {
                total += Globals.Sheet1.Cells[8,col].Value;
                col++;
            }

            return total;
        }

        static DateTime getDate(string _date)
        {
            CultureInfo cultureInfo = new CultureInfo("en-US");
            DateTime wsDate = DateTime.Parse(_date, cultureInfo);

            return wsDate;
        }

        //This function will return the number of rows we need to filter out the data corresponding to the given lookback period.
        public int end_row(DateTime lookback_date)
        {
            int endRow = 3;
            int j = 3;
            //DateTime lookback_date = getDate(lookback_string);
            while(string.IsNullOrWhiteSpace(Globals.Sheet8.Cells[j, 1].Value?.ToString()) == false)
            {
                DateTime date_check = getDate(Globals.Sheet8.Cells[j,1].Value.ToString());
                if (date_check < lookback_date)
                {
                    break;
                }
                else
                {
                    endRow++;
                }
                j++;
            }

            return endRow;
        }

        //This is a function that computes the daily volatility. The volatility floor given in the home spreadsheet is used as the volatility on the last day.
        public double ewma_vol(double vol_0 , double return_0 , double lambda)
        {
            double vol = lambda*Math.Pow(vol_0 , 2) + (1-lambda)*Math.Pow(return_0, 2);
            return vol;
        }

        public int col_start(int pay_freq)
        {
            //This function named col_start will do a look up on the first row of the spreadsheet containing the zero rates. The function will look up 
            //what value corresponds to the payment frequency and sets the column where this value is found as the start column for calculating the first
            //cashflow.

            int start_col = 0;

            int j = 2; //we exclude the first column of the worksheet because it is blank.
            while (string.IsNullOrWhiteSpace(Globals.Sheet3.Cells[1, j].Value?.ToString()) == false)
            {
                if (Globals.Sheet3.Cells[1,j].Value == pay_freq)
                {
                    start_col = j;
                    break;
                }
                else if (Globals.Sheet3.Cells[1, j].Value == pay_freq)
                {
                    start_col = j;
                    break;
                }
                else if (Globals.Sheet3.Cells[1, j].Value == pay_freq)
                {
                    start_col = j;
                    break;
                }
                else if (Globals.Sheet3.Cells[1, j].Value == pay_freq)
                {
                    start_col = j;
                    break;
                }

                j++;
            }

            return start_col;
        }

        
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Valuate_Click(object sender, RibbonControlEventArgs e)
        {
            int col = 3;  //Starting column for the data the user inputs in the home worksheet.

            while(string.IsNullOrWhiteSpace(Globals.Sheet1.Cells[7, col].Value?.ToString()) == false)
            {
                Globals.Sheet8.Cells[2,col-1].Value = Globals.Sheet1.Cells[7, col].Value.ToString() + " yr_swap";

                //Collect the data the user inputs from the home worksheet.
                double tenor_ = Globals.Sheet1.Cells[7, col].Value;
                double notional_ = Globals.Sheet1.Cells[8, col].Value;
                double spread_ = Globals.Sheet1.Cells[9, col].Value;
                string pay_freq_ = Globals.Sheet1.Cells[10, col].Value.ToString();
                /////////////////////////////////////////////////////////////

                //Define a colum increase: this is so that the right column is selected corresponding to the reset point.
                int col_inc = 0;

                //Define a column start: this is so that we know when the first cashflow will be made.
                int col_start_ = 0;

                //Define column end: this is so that we know when the last cashflow will be made.
                int col_end = 0;

                if (pay_freq_.ToUpper() == "MONTHLY")
                {
                    col_start_ = col_start(1);
                    col_inc = 1;
                    col_end = 2 + (int)(col_inc * (12 * tenor_));  //For a 5 year swap with monthly payments, the last cashflow takes place at the month 12*5 = 60.
                }
                else if (pay_freq_.ToUpper() == "QUARTERLY")
                {
                    col_start_ = col_start(3);
                    col_inc = 3;
                    col_end = 2 +  (int) (col_inc * (4 * tenor_));
                }
                else if (pay_freq_.ToUpper() == "SEMI-ANNUALLY")
                {
                    col_start_ = col_start(6);
                    col_inc = 6;
                    col_end = 2 + (int)(col_inc * (2 * tenor_));
                }
                else if (pay_freq_.ToUpper() == "YEARLY")
                {
                    col_start_ = col_start(12);
                    col_inc = 12;
                    col_end = 2 + (int)(col_inc * tenor_);
                }

                int rowCount = countRows();
                //Now we are going to find the daily swap rates
                for (int row = 3; row < rowCount + 2; row++)
                {
                   double pv_float = PV.PV_float(spread_,notional_,row,col_start_,col_inc,col_end);
                   double pv_fixed = PV.PV_fixed(1, notional_, row, col_start_, col_inc, col_end); //pass swap_rate = 1, because at this point we don't know what the swap rate is.

                   Globals.Sheet8.Cells[row, col-1].Value = pv_float / pv_fixed;

                }

                col++;
            }
        }

        private void find_PV_Click(object sender, RibbonControlEventArgs e)
        {
            int col = 3;
            while(string.IsNullOrWhiteSpace(Globals.Sheet1.Cells[7,col].Value?.ToString()) == false)
            {
                Globals.Sheet8.Cells[2, col + 8].Value = Globals.Sheet1.Cells[7, col].Value.ToString() + " yr_swap";

                //Collect the data the user inputs from the home worksheet.
                double tenor_ = Globals.Sheet1.Cells[7, col].Value;
                double notional_ = Globals.Sheet1.Cells[8, col].Value;
                double spread_ = Globals.Sheet1.Cells[9, col].Value;
                string pay_freq_ = Globals.Sheet1.Cells[10, col].Value.ToString();
                /////////////////////////////////////////////////////////////

                //Define a colum increase: this is so that the right column is selected corresponding to the reset point.
                int col_inc = 0;

                //Define a column start: this is so that we know when the first cashflow will be made.
                int col_start_ = 0;

                //Define column end: this is so that we know when the last cashflow will be made.
                int col_end = 0;

                if (pay_freq_.ToUpper() == "MONTHLY")
                {
                    col_start_ = col_start(1);
                    col_inc = 1;
                    col_end = 2 + (int)(col_inc * (12 * tenor_));  //For a 5 year swap with monthly payments, the last cashflow takes place at the month 12*5 = 60.
                }
                else if (pay_freq_.ToUpper() == "QUARTERLY")
                {
                    col_start_ = col_start(3);
                    col_inc = 3;
                    col_end = 2 + (int)(col_inc * (4 * tenor_));
                }
                else if (pay_freq_.ToUpper() == "SEMI-ANNUALLY")
                {
                    col_start_ = col_start(6);
                    col_inc = 6;
                    col_end = 2 + (int)(col_inc * (2 * tenor_));
                }
                else if (pay_freq_.ToUpper() == "YEARLY")
                {
                    col_start_ = col_start(12);
                    col_inc = 12;
                    col_end = 2 + (int)(col_inc * tenor_);
                }

                //Now we find the present value with the swap rate fixed to the swap rate calculated at evaluation date.
                double eval_swap_rate = 0;
                DateTime Eval_date = getDate(Globals.Sheet1.Cells[4, 3].Value.ToString());
                int date_row = 3;
                while (string.IsNullOrWhiteSpace(Globals.Sheet8.Cells[date_row, 1].Value?.ToString()) == false)
                {
                    DateTime ws_date = getDate(Globals.Sheet8.Cells[date_row, 1].Value.ToString());
                    if ( ws_date == Eval_date )
                    {
                        eval_swap_rate = Globals.Sheet8.Cells[date_row, col - 1].Value;
                        break;
                    }
                    date_row++;
                }

                int rowCount = countRows();

                for (int row = 3; row < rowCount + 2; row++)
                {
                    double pv_fixed = PV.PV_fixed(eval_swap_rate, notional_, row, col_start_, col_inc, col_end);
                    Globals.Sheet8.Cells[row,col+8].Value = pv_fixed;
                }

                col++;
            }
        }

        private void find_Portfolio_PV_Click(object sender, RibbonControlEventArgs e)
        {
            //onclick this button calculates the PV of the portfolio by summing the daily PV of the individual IRS products.

            int rows = countRows();

            for (int row = 3; row < rows + 2; row++)
            {
                int col = 11;
                double pv_sum = 0;
                while (string.IsNullOrWhiteSpace(Globals.Sheet8.Cells[row, col].Value?.ToString()) == false)
                {
                    pv_sum += Globals.Sheet8.Cells[row, col].Value;
                    col++;
                }
                Globals.Sheet8.Cells[row, 20].Value = pv_sum;
            }

            //Finding the portfolio Profit and Loss:
            int i_row = 4;
            double returns = 0;

            while (string.IsNullOrWhiteSpace(Globals.Sheet8.Cells[i_row,20].Value?.ToString()) == false)
            {
                //weights = (1 - lambda) * Math.Pow(lambda, exp);
                returns = (Globals.Sheet8.Cells[i_row - 1, 20].Value - Globals.Sheet8.Cells[i_row, 20].Value) / Globals.Sheet8.Cells[i_row, 20].Value;

                //Portfolio PnL:
                Globals.Sheet8.Cells[i_row, 22].Value = returns;

                i_row++;
            }
        }

        private void find_VaR_Click(object sender, RibbonControlEventArgs e)
        {
            //EWMA weights: this list stores the weights used in implementing EWMA
            List<double> ewma_vol_lst = new List<double>();
            //////////////////////////////////////////////////////////////////////////

            //Calculate VaR
            List<double> scaled_returns = new List<double>();
            List<double> unscalled_returns = new List<double>();

            int row_date = 4; //Start in row=4 because there is no return for the most recent date. For generality, this should be initialized as "row_eval_date + 1".
            DateTime unscalled_lookback_date = getDate(Globals.Sheet1.Cells[4, 9].Value.ToString());
            DateTime scalled_lookback_date = getDate(Globals.Sheet1.Cells[4, 6].Value.ToString());
            double scaled_return_ = 0;

            int rows = countRows();
            double vol_floor = Globals.Sheet1.Cells[15, 4].Value;
            Globals.Sheet8.Cells[rows + 1, 24].Value = vol_floor;

            double lambda = Globals.Sheet1.Cells[14, 4].Value;  //Defined Lambda in the home worksheet

            double return_0 = 0;
            int i = 0;
            while (string.IsNullOrWhiteSpace(Globals.Sheet8.Cells[row_date, 1].Value?.ToString()) == false)
            {

                //Get daily volatilities:
                return_0 = Globals.Sheet8.Cells[rows - i + 1, 22].Value;
                double vol = ewma_vol(Globals.Sheet8.Cells[rows - i + 1, 24].Value, return_0, lambda);
                Globals.Sheet8.Cells[rows - i, 24].Value = Math.Max(vol, vol_floor);

                row_date++;
                i++;
            }

            double num_vol = Globals.Sheet8.Cells[3, 24].Value;
            double factor = 0;

            for (int k = 4; k < rows+2; k++)
            {
                DateTime ws_date = getDate(Globals.Sheet8.Cells[k,1].Value.ToString());
                double den_vol = Globals.Sheet8.Cells[k, 24].Value;

                factor = num_vol / den_vol;
                Globals.Sheet8.Cells[k, 26].Value = factor * Globals.Sheet8.Cells[k, 22].Value;
                double scaled_return = Globals.Sheet8.Cells[k, 26].Value;
                double unscalled_return = Globals.Sheet8.Cells[k, 22].Value;

                if (ws_date >= scalled_lookback_date && ws_date >= unscalled_lookback_date)
                {
                    scaled_returns.Add(scaled_return);
                    unscalled_returns.Add(unscalled_return);
                }
                else if (ws_date < scalled_lookback_date && ws_date >= unscalled_lookback_date)
                {
                    unscalled_returns.Add(unscalled_return);
                }
            }

            double[] scaled_returns_array = scaled_returns.ToArray();
            double[] unscaled_returns_array = unscalled_returns.ToArray();

            Globals.Sheet1.Cells[26, 6].Value = percentile_func.Percentile(scaled_returns_array, 0.001)*Math.Sqrt(5);
            Globals.Sheet1.Cells[26, 9].Value = percentile_func.Percentile(unscaled_returns_array, 0.001) * Math.Sqrt(5);
            Globals.Sheet1.Cells[26, 3].Value = (0.75*Globals.Sheet1.Cells[26,6].Value + 0.25*Globals.Sheet1.Cells[26,9].Value)*total_exposure();
        }
    }
}

//to do:
//The code must be able to pick up the evaluation date from the one the user inputs. So that when we need to caculate the PnL, the PnL for the cells corresponding to the row of
//the evaluation date must be null.

//Also, the code must be able select the starting row that corresponds to the row of the date picked as the evaluation date + 1.
