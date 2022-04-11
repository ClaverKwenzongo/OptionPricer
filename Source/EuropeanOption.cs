using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MathNet.Numerics.Distributions;


namespace OptionPricer
{
    class EuropeanOption
    {
        private string ticker { get; set; }
        private double Strike { get; set; }
        private string OptionType { get; set; }
        private double TotalDays { get; set; }
        
        public EuropeanOption(string _Ticker_, double strike, string option_type, double totalDays)
        {
            ticker = _Ticker_;
            Strike = strike;
            OptionType = option_type;
            TotalDays = totalDays;            
        }

        int days_in_year = 365;
        public double optionPrice(double SharePrice, double risk_free, double vol, double div_yield, double size)
        {
            double optionPrice = 0;

            double d_1 = (Math.Log(SharePrice / Strike) + (risk_free - div_yield + Math.Pow(vol, 2) / 2) * (TotalDays / days_in_year)) / (vol * Math.Sqrt(TotalDays / days_in_year));
            double d_2 = d_1 - vol*Math.Sqrt(TotalDays/days_in_year);

            if (OptionType == "PUT")
            {
                int psi = -1;
                double N_d1 = Normal.CDF(0,1,psi*d_1);
                double N_d2 = Normal.CDF(0, 1, psi * d_2);
                optionPrice = psi * (SharePrice * Math.Pow(Math.E, -div_yield * TotalDays/days_in_year) * N_d1 - Strike * Math.Pow(Math.E, -risk_free * TotalDays/days_in_year) * N_d2);
                optionPrice = size * optionPrice;
            }
            else if (OptionType == "CALL")
            {
                int psi = 1;
                double N_d1 = Normal.CDF(0, 1, psi * d_1);
                double N_d2 = Normal.CDF(0, 1, psi * d_2);
                optionPrice = psi * (SharePrice * Math.Pow(Math.E, -div_yield * TotalDays/days_in_year) * N_d1 - Strike * Math.Pow(Math.E, -risk_free * TotalDays/days_in_year) * N_d2);
                optionPrice = size * optionPrice;
            }
            else
            {
                throw new Exception("Option type can only be put or call.");
            }
            return optionPrice;
        }

        public double sensitivity(double _SharePrice_, double _risk_free_, double _div_yield_, double _vol_, string greek_type)
        {
            double greek = 0;

            double d_1 = (Math.Log(_SharePrice_ / Strike) + (_risk_free_ - _div_yield_ + Math.Pow(_vol_, 2) / 2) * (TotalDays / days_in_year)) / (_vol_ * Math.Sqrt(TotalDays / days_in_year));

            if (greek_type.ToUpper() == "DELTA")
            {
                if (OptionType == "CALL")
                {
                    greek = Math.Pow(Math.E, -_div_yield_ * TotalDays/days_in_year) * Normal.CDF(0, 1, d_1);
                }
                else if (OptionType == "PUT")
                {
                    greek = -Math.Pow(Math.E, -_div_yield_ * TotalDays/days_in_year) * Normal.CDF(0, 1, -d_1);
                }
                else
                {
                    Console.WriteLine("Enter a valid option type (call or put)");
                }
            }
            else if (greek_type.ToUpper() == "GAMMA")
            {
                greek = Math.Pow(Math.E, -_div_yield_ * TotalDays/days_in_year) * Normal.PDF(0, 1, d_1) / (_SharePrice_ * _vol_ * Math.Sqrt(TotalDays/days_in_year));

            }
            else if (greek_type.ToUpper() == "VEGA")
            {
                greek = _SharePrice_ * Math.Pow(Math.E, -_div_yield_ * TotalDays/days_in_year) * (Math.Sqrt(TotalDays/days_in_year)) * Normal.PDF(0, 1, d_1);

            }

            return greek;
        }

    }
}
