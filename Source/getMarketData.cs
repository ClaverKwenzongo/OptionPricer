using OfficeOpenXml;
using System.Globalization;
// See https://aka.ms/new-console-template for more information
namespace OptionPricer
{
    internal static class getMarketData
    {
        ///////////Get the share price from the excel file.
        public static async Task<List<MarketData>> LoadExcelFile(FileInfo path, string shareName, string date, double tenor)
        {
            List<MarketData> result = new();

            using var package = new ExcelPackage(path);


            await package.LoadAsync(path);

            //Get the worksheet with the share prices
            var ws = package.Workbook.Worksheets[0];

            //Get the worksheet with the interest rates
            var ws_rates = package.Workbook.Worksheets[1];

            //Get the worksheet with volatilities
            var ws_vol = package.Workbook.Worksheets[2];

            //Get the worksheet with the divident yields
            var ws_div = package.Workbook.Worksheets[3];

            ///////////////This method allows us to convert the dates from the excel sheet in a format that can be compared with the date the user inputs.
            ///This line of code is not included in the getDateRow method because, we need to reuse it for getting share prices
            static DateTime getDate(string _date)
            {
                CultureInfo cultureInfo = new CultureInfo("en-US");
                DateTime wsDate = DateTime.Parse(_date, cultureInfo);

                return wsDate;
            }
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ///Get the excel row that contains the date corresponding to the start date entered by the user.
            static int getDateRow(ExcelWorksheet xl_ws, string _user_date)
            {
                int date_r = 0;
                int row2 = 3;
                ExcelWorksheet _ws = xl_ws;
                //Look for the risk free rates that corresponds to the tenor date calculated from the user's entered dates
                while (string.IsNullOrWhiteSpace(_ws.Cells[row2, 1].Value?.ToString()) == false)
                {
                    //First we find the row where the date that is equal to the tenor date (end date) entered by the user
                    DateTime ws_rates_date = getDate(_ws.Cells[row2, 1].Value.ToString());
                    var ws_rates_date_con = ws_rates_date.ToString("dd-MMM-yyyy");

                    if (ws_rates_date_con == _user_date)
                    {
                        date_r = row2;
                        break;
                    }

                    row2++;
                }
                return date_r;
            }

            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ///This function when called it returns the interest rate, implied volatility and the dividend yield depending on the excel spreadsheet, the start date, row where the start
            ///date is situated and the column to begin the search (this is because the dividend and volatility datas do not all begin from the same column for individual shares.

            static double getPercentage(ExcelWorksheet per_xl_ws, double _op_tenor, int _date_row, int _start_col)
            {
                double percent = 0;
                ExcelWorksheet percent_ws = per_xl_ws;
                int col2 = _start_col;

                while (string.IsNullOrWhiteSpace(percent_ws.Cells[_date_row, col2].Value?.ToString()) == false)
                {
                    //Pick out the rates that corresponds to the start date entered by the user.
                    //The start date entered by the user may not be in the same row for the share prices and the interest rates.
                    //This is why this is done separately from the share price.

                    int rate_col = _start_col;
                    //Check whether the column that follows the current column is not empty. This allows for you to compare tenors
                    //between two tenors, for the purpose of interpolation.
                    if (string.IsNullOrWhiteSpace(percent_ws.Cells[2, col2 + 1].Value?.ToString()) == false)
                    {
                        //Check whether the days to tenor is exactly equal to the tenor in the excel file
                        if (_op_tenor == double.Parse(percent_ws.Cells[2, col2].Value.ToString()))
                        {
                            percent = double.Parse(percent_ws.Cells[_date_row, col2].Value.ToString());
                            break;
                        }
                        else if (_op_tenor == double.Parse(percent_ws.Cells[2, col2 + 1].Value.ToString()))
                        {
                            percent = double.Parse(percent_ws.Cells[_date_row, col2 + 1].Value.ToString());
                            break;
                        }
                        //In the case the tenor calculated is between two values, do a very simple interpolation: Just take the average of the two risk free rates.
                        //In the case the number of days to tenor is between two values, call the interpolation function here. Here we do a basic calculation and take the average.
                        else if (double.Parse(percent_ws.Cells[2, col2].Value.ToString()) < _op_tenor && _op_tenor < double.Parse(percent_ws.Cells[2, col2 + 1].Value.ToString()))
                        {
                            percent = (double.Parse(percent_ws.Cells[_date_row, col2].Value.ToString()) + double.Parse(percent_ws.Cells[_date_row, col2 + 1].Value.ToString())) / 2;
                            break;
                        }
                    }
                    else
                    {
                        percent = double.Parse(percent_ws.Cells[_date_row, col2].Value.ToString());
                        break;
                    }

                    col2++;
                }

                return percent;
            }

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ////////// This is a function to get the starting column from where the look up for the volatilities and dividends yields of a given share will begin //////////

            static int getStartCol(ExcelWorksheet _xl_ws, string _share_name)
            {
                int s_col = 0;

                ////First we need to count how many tenors (col_count) they are before we can switch columns to look up the tenors data for each share. Assuming dividends and implied volatility
                //// of all the shares have the same tenors.
                int col_count = 1;
                int col3 = 2;
                while (string.IsNullOrWhiteSpace(_xl_ws.Cells[2, col3].Value?.ToString()) == false)
                {
                    col_count++;
                    col3++;
                }

                int _col = 2;

                //Lets pick the column from where to begin the look up. Because all the dividend(implied volatility) are in the same sheet, we need to know from which column we need
                //to begin the look up so we get the right dividend/volatility corresponding to the share _Ticker_ entered by the user.

                int start_col = 0;
                while (string.IsNullOrWhiteSpace(_xl_ws.Cells[1, _col].Value?.ToString()) == false)
                {
                    //Get the first three letters of the string for comparison with the user's inputed _Ticker_.
                    string xl_share = _xl_ws.Cells[1, _col].Value.ToString();

                    if (xl_share.Substring(0, 3) == _share_name)
                    {
                        s_col = _col;
                        break;
                    }
                    _col = _col + col_count; //Because the dividend yield and volatilities are all in the same excel sheet, we need to know where the data
                    //for specific share entered by the user begins on the worksheet.
                }

                return s_col;
            }

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ////The codes below will extract the share price of the share corresponding to the _Ticker_ entered by the user and the start date entered by the user
            int col = 1;
            int date_col = 1;
            int share_col = 0;

            //Pick out the date column and the share name column
            while (string.IsNullOrWhiteSpace(ws.Cells[2, col].Value?.ToString()) == false)
            {
                //Pick out the column that contains share prices of the share name entered by the user.
                if (ws.Cells[2, col].Value.ToString() == shareName)
                {
                    share_col = col;
                    break;
                }

                col++;
            }

            int row = 3;
            double share_price = 0;
            //Look for the date to the tenor the user entered in the date column in excel
            while (string.IsNullOrWhiteSpace(ws.Cells[row, date_col].Value?.ToString()) == false)
            {
                DateTime ws_date = getDate(ws.Cells[row, date_col].Value.ToString());
                var ws_date_con = ws_date.ToString("dd-MMM-yyyy");
                if (ws_date_con == date)
                {
                    share_price = Math.Round(double.Parse(ws.Cells[row, share_col].Value.ToString()), 2);
                    break;
                }

                row++;
            }
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //// Get the interest rates values - code below ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            int date_row = getDateRow(ws_rates, date);
            double risk_free = getPercentage(ws_rates, tenor, date_row, 2);

            ////Get implied volatility values - code below/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            ////Lets pick the column from where to begin the look up
            int start_col_vol = getStartCol(ws_vol, shareName);
            ////////////////////////////////////////////////////////

            int date_row_vol = getDateRow(ws_vol, date);
            double _vol = getPercentage(ws_vol, tenor, date_row_vol, start_col_vol);
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ////Get the dividend yield values - code below///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            ////Lets get the column from where we need to begin to look up
            int start_col_div = getStartCol(ws_div, shareName);
            int date_row_div = getDateRow(ws_div, date);
            double _div = getPercentage(ws_div, tenor, date_row_div, start_col_div);
            ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            MarketData sharePriceData = new();
            sharePriceData.SharePrice = share_price;
            sharePriceData.riskFree = risk_free;
            sharePriceData.volatility = _vol;
            sharePriceData.dividend = _div;

            result.Add(sharePriceData);

            return result;
        }
    }
}