using System;

namespace IM_IRS_Demo
{
    public class Discounts
    {
        //This class contains the methods to find the discount factors from the zero rates 
        public double getDFs( int row , int col )
        {
            //We need a ticker to check whether the number of days exceeds 365, that way we adjust the divisor of days accordingly

           // double days_divisor = 365;
            //double ticker = Globals.Sheet3.Cells[1,col].Value; //This is defined ticker on the workbook, designed so that we can identify columns in the sheet named "Zeroes"


            double days = Globals.Sheet9.Cells[row,col].Value;
            double rate = Globals.Sheet3.Cells[row,col].Value;

            double df = Math.Pow(Math.E, -rate * days/365);  //For continuously compounded 

            return Math.Round(df, 9);
        }

    }
}
