using System;

namespace IM_IRS_Demo
{
    public class Discounts
    {
        //This class contains the methods to find the discount factors from the zero rates 
        public double getDFs( int row , int col )
        {

            double days = Globals.Sheet9.Cells[row,col].Value;
            double rate = Globals.Sheet3.Cells[row,col].Value;

            double df = Math.Pow(Math.E, -rate * days/365);  //For continuously compounded 

            return df;
        }

    }
}
