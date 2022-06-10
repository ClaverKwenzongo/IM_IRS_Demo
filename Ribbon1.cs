﻿using Microsoft.Office.Tools.Ribbon;
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
                    col_end = 2 + (int)(col_inc * (6 * tenor_));
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
                   double pv_fixed = PV.PV_fixed(notional_, row, col_start_, col_inc, col_end);

                   Globals.Sheet8.Cells[row, col-1].Value = pv_float / pv_fixed;

                }

                col++;
            }
        }
    }
}
