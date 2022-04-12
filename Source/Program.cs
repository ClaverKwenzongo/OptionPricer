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

            List<string> start_dates = new List<string> { "30/03/2016", "30/03/2012", "30/03/2012", "30/03/2012" };
            List<string> end_dates = new List<string> { "30/03/2017","26/09/2012", "28/06/2012", "25/12/2012" };
            List<string> tickers = new List<string> {"rwx", "npn", "sol", "sab" };
            List<double> strikes = new List<double> {100, 400, 385, 300 };
            List<string> positions = new List<string> {"long","short", "long","short" };
            List<string> optionTypes = new List<string> { "call", "put", "call", "put" };

            foreach (string ticker in tickers)
            {
                double op_strike = strikes[tickers.IndexOf(ticker)];

                string user_op_type = optionTypes[tickers.IndexOf(ticker)];

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

                double size = 500000; //Suppose all the shares are equally sized.

                string op_position = positions[tickers.IndexOf(ticker)];

                if (op_position.ToUpper() == "LONG" )
                {
                    size = size;
                }
                else if (op_position.ToUpper() == "SHORT")
                {
                    size = -size;
                }
                else
                {
                    throw new Exception("Option position can only be long or short");
                }

                string user_ticker = ticker;

                DateTime end_date = DateTime.Parse(end_dates[tickers.IndexOf(ticker)], culture);
                DateTime start_date = DateTime.Parse(start_dates[tickers.IndexOf(ticker)], culture);
                var user_start_date = start_date.ToString("dd-MMM-yyyy");

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

                    ///This is the test case.....

                    op = europeanOption.optionPrice(_share_price = 100, rf = 0.05, implied_vol = 0.2, div_yield = 0, size);

                    double mkt_option_price = op/size;

                    var impliedVol = new ImpliedVol(mkt_option_price, op_strike, days, _share_price = 100, rf = 0.05, Math.Pow(10, -10), div_yield = 0);

                    imp_vol = impliedVol.newton_vol(psi);

                    //We dont have to add this in the calculation of the portfolio value since it is a made up share.

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

                table.AddRow(ticker, end_dates[tickers.IndexOf(ticker)], op_strike, size,op_position, op, del, gam, veg, imp_vol + " %");

            }

            table.Write();

            double portfolio_value = Math.Round(portfolio.Sum(),2);

            var val = string.Format("{0:N2}", portfolio_value);

            Console.WriteLine($"The value of your portfolio is R {val}");

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        }
    }
}

