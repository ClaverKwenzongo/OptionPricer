// See https://aka.ms/new-console-template for more information
using OfficeOpenXml;
using System;
using System.Globalization;
using ConsoleTables;

namespace OptionPricer
{
    class Program
    {
        static async Task Main(string[] args)
        {
            CultureInfo culture = new CultureInfo("es-ES");

            //Table of results...
            var table = new ConsoleTable("Shares", "Maturity Dates", "Strike Prices", "Number of Shares","Position", "Option Prices", "Delta", "Gamma", "Vega", "Implied Volatility");

            List<double> portfolio = new ();

            //Enter day of entry of option
            Console.WriteLine("Enter option start date: dd/mm/yyyy");
            String myStartDate = Console.ReadLine();
            DateTime start_date = DateTime.Parse(myStartDate, culture);
            var user_start_date = start_date.ToString("dd-MMM-yyyy");

            Console.WriteLine("Portfolio shares separated by commas:");
            string user_list = Console.ReadLine();
            user_list.Split(',');

            foreach (string ticker in user_list.Split(','))
            {
                Console.WriteLine($"What is the strike price of {ticker}:");
                double op_strike = double.Parse(Console.ReadLine());

                Console.WriteLine($"Option type of {ticker}:");
                string user_op_type = Console.ReadLine();

                //Check option type: Put - right to sell (-1), Call - right to buy (+1).
                int psi;
                if (user_op_type.ToUpper() == "PUT" )
                {
                    psi = -1;
                }
                else if (user_op_type.ToUpper() == "CALL" )
                {
                    psi = 1;
                }
                else
                {
                    throw new Exception("Enter either put or call for option type");
                }

                Console.WriteLine($"How many shares of {ticker}:");
                double size = double.Parse(Console.ReadLine());

                Console.WriteLine($"Enter the option position of {ticker} (long or short):");
                string op_position = Console.ReadLine();

                //check option position, for long position: buying (you are giving money to the seller). For short position: selling (you are getting money from the buyer)
                if (op_position.ToUpper() == "LONG" )
                {
                    size = -size;
                }
                else if (op_position.ToUpper() == "SHORT")
                {
                    size = size;
                }
                else
                {
                    throw new Exception("Option position can only be long or short");
                }

                string user_ticker = ticker;

                Console.WriteLine($"What is the maturity date of {ticker}:");
                string _end_date = Console.ReadLine();
                DateTime end_date = DateTime.Parse(_end_date, culture);

                TimeSpan interval = end_date - start_date;
                double days = (double)interval.TotalDays;

                ExcelPackage.LicenseContext = LicenseContext.Commercial;

                //Change this file path...
                var path = new FileInfo(@"C:\Users\ClaverKwenzongo\source\repos\OptionPricer\Market Data.xlsx");

                //////////////////////////////////////////////////////////////////////////////////////////////////////////////// 
                ///

                var europeanOption = new EuropeanOption(user_ticker, op_strike, psi, days);

                double rf = 0;
                double implied_vol = 0;
                double div_yield = 0;
                double _share_price = 0;
                double op = 0;
                double del = 0;
                double gam = 0;
                double veg = 0;
                double imp_vol = 0;

                if (ticker.ToUpper() == "RWX")
                {
                    //Console.WriteLine($"To calculate the implied volatility enter the market price of {ticker} option:");
                    //double mkt_option_price = double.Parse(Console.ReadLine());

                    ///This is the test case.....
                    op = europeanOption.optionPrice(_share_price = 100, rf = 0.05, implied_vol = 0.2, div_yield = 0, size);

                    double mkt_option_price = op/size;

                    var impliedVol = new ImpliedVol(mkt_option_price, op_strike, days, _share_price = 100, rf = 0.05, Math.Pow(10, -10), div_yield = 0);

                    imp_vol = impliedVol.newton_vol(psi);

                    //Add the calculated european call option price into the portfolio list for summation.
                    portfolio.Add(op);

                    op = Math.Round(op / size,5);
                    del = europeanOption.sensitivity(_share_price = 100, rf = 0.05, div_yield = 0, implied_vol = 0.2, "delta");
                    gam = europeanOption.sensitivity(_share_price = 100, rf = 0.05, div_yield = 0, implied_vol = 0.2, "gamma");
                    veg = europeanOption.sensitivity(_share_price = 100, rf = 0.05, div_yield = 0, implied_vol = 0.2, "vega");
                }
                else
                {
                    //Get the share price corresponding to the _Ticker_ entered by the user and the time to tenor of the option.//
                    List<MarketData> marketDatas = await getMarketData.LoadExcelFile(path, user_ticker.ToUpper(), user_start_date, days);

                    foreach (MarketData marketData in marketDatas)
                    {
                        rf = marketData.riskFree;
                        implied_vol = marketData.volatility;
                        div_yield = marketData.dividend;
                        _share_price = marketData.SharePrice;
                    }

                    //Console.WriteLine($"To calculate the implied volatility enter the market price of {ticker} option:");
                    //long mkt_option_price = (long)double.Parse(Console.ReadLine());

                    op = europeanOption.optionPrice(_share_price, rf, implied_vol, div_yield, size);

                    double mkt_option_price = op/size;

                    var impliedVol = new ImpliedVol(mkt_option_price, op_strike, days, _share_price, rf, Math.Pow(10, -10), div_yield);

                    imp_vol = impliedVol.newton_vol(psi);

                    //Add the calculated european call option price into the portfolio list for summation.
                    portfolio.Add(op);

                    op = Math.Round(op / size,5);
                    del = europeanOption.sensitivity(_share_price, rf, div_yield, implied_vol, "delta");
                    gam = europeanOption.sensitivity(_share_price, rf, div_yield, implied_vol, "gamma");
                    veg = europeanOption.sensitivity(_share_price, rf, div_yield, implied_vol, "vega");

                }

                table.AddRow(ticker, _end_date, op_strike, size,op_position, op, del, gam, veg, imp_vol + " %");

            }

            table.Write();

            double portfolio_value = Math.Round(portfolio.Sum(),2);

            var val = string.Format("{0:N2}", portfolio_value);

            Console.WriteLine($"The value of your portfolio is R {val}");

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        }
    }
}

