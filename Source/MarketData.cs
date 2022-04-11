// See https://aka.ms/new-console-template for more information

namespace OptionPricer
{

    //This class is for storing the information queried from the given excel file.
    public class MarketData
    {
        public double SharePrice { get; set; }
        public double riskFree { get; set; }
        public double volatility { get; set; }
        public double dividend { get; set; }

    }
}

