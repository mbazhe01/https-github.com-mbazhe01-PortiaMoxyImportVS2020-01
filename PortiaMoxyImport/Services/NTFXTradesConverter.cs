using PortiaMoxyImport.Entities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Services
{
    public class NTFXTradesConverter : IConvertNTFXTradesToAIM
    {
        private  List<NTFXTradeDTO> _trades;
        private readonly string _outputFilePath;
        private readonly List<string> _flipCurrencyList;
        private readonly HashSet<string> _adjusterUsed = new HashSet<string>();

        private readonly TradeAdjusterFactory _adjusterFactory;

        public NTFXTradesConverter(List<NTFXTradeDTO> trades, string outputFilePath, List<string> flipCurrencyList)
        {
            _trades = trades;
            _outputFilePath = outputFilePath;

            _adjusterFactory = new TradeAdjusterFactory(
                new BuyUsdNonUsdAdjuster(flipCurrencyList),
                new BuyNonUsdUsdAdjuster(flipCurrencyList),
                new SellUsdNonUsdAdjuster(),
                new SellNonUsdUsdAdjuster(flipCurrencyList),
                new BuyNonUsdNonUsdAdjuster(flipCurrencyList),
                new SellNonUsdNonUsdAdjuster());
            _flipCurrencyList = flipCurrencyList;
        }

        public NTFXTradesConverter(List<NTFXTradeDTO> trades, string outputFilePath)
        {
            if (trades == null) throw new ArgumentNullException(nameof(trades));
            if (string.IsNullOrWhiteSpace(outputFilePath)) throw new ArgumentNullException(nameof(outputFilePath));

            _trades = trades;
            _outputFilePath = outputFilePath;

            //string flipCurrencies = Util.getAppConfigVal("FlipRateCurrencies");
            //_flipCurrencyList = flipCurrencies.Split(',').Select(c => c.Trim().ToUpper()).ToList();
        }


        private string[] ConvertTradeToAIMRow(NTFXTradeDTO trade)
        {
            // adjust trade base on Buy/Sell and base currency if needed

           // trade = AdjustTrade(trade);
            //trade = _adjusterFactory.GetAdjuster(trade).Adjust(trade);

            // 47 columns (0..46)
            var fields = new string[47];

            // Column 0: Account
            fields[0] = trade.Account;

            // Column 1: Buy/Sell ("by" or "sl")
            fields[1] = MapSide(trade.BuySell);

            // Column 2: empty

            // Column 3: empty (no mapping required)
            fields[3] = "";

            // Column 4: "-XXX FWD CASH-"
            fields[4] = "-" + trade.Currency + " FWD CASH-";

            // Column 5: Trade date MMddyyyy
            fields[5] = trade.TradeDate.ToString("MMddyyyy", CultureInfo.InvariantCulture);

            // Column 6: Value date MMddyyyy
            fields[6] = trade.ValueDate.ToString("MMddyyyy", CultureInfo.InvariantCulture);

            // Column 7: empty

            // Column 8: Amount (trade currency)
            fields[8] = FormatAmount(trade.Amount);

            // 9–10: empty

            // Column 11: empty (no mapping required)
            fields[11] = "";

            // Column 12: "-XXX FWD CASH-" for other currency
            fields[12] = "-" + trade.OtherCurrency + " FWD CASH-";

            // Column 13: NT forward rate (rounded)
            //fields[13] = FormatRate6Decimals(trade.ForwardRate);

            // Column 14: empty

            // Column 15: precise rate (Amount / OtherAmount)
            fields[15] = FormatRate6Decimals(trade.ForwardRate);

            // Column 16: "y"
            fields[16] = "y";

            // Column 17: Other amount
            fields[17] = FormatAmount(trade.OtherAmount);

            // 18–23: empty

            // Column 24: Broker
            fields[24] = trade.Broker;

            // 25–27: empty

            // Column 28: "n"
            fields[28] = "n";

            // Column 29: "254"
            fields[29] = "254";

            // 30–40: empty

            // Column 41: "1"
            fields[41] = "1";

            // 42–43: empty

            // Column 44: "n"
            fields[44] = "n";

            // Column 45: "y"
            fields[45] = "y";

            // Column 46: empty

          


            return fields;
        }

        private  NTFXTradeDTO AdjustTrade(NTFXTradeDTO trade)
        {
            try
            {

                if (trade == null) throw new ApplicationException("AdjustTrade: Undefined trade. ? ? ? ");
                
                if (trade.BuySell == "B")
                {
                    // Buy trade 
                    if (trade.Currency.ToUpper().Equals("USD"))
                    {
                        // buy side is USD
                        if (!_flipCurrencyList.Contains(trade.OtherCurrency))
                        {
                            trade.ForwardRate = 1 / trade.ForwardRate;
                        }
                    }
                    else
                    {
                        // buy side is Non USD
                        throw new ApplicationException("AdjustTrade Error:  Buy & Buy side is NON USD not implemented.");
                    }

                    return trade;
                }

                if (trade.BuySell == "S")
                {
                    // Sell trade -> make it buy
                    trade.BuySell = "B";

                    if (trade.Currency.ToUpper().Equals("USD"))
                    {
                        throw new ApplicationException("AdjustTrade Error:  Sell & Sell side is USD not implemented.");
                    }
                    else
                    {
                        // sell side is NOT USD
                        if (_flipCurrencyList.Contains(trade.OtherCurrency))
                        {
                           

                        }else
                        {
                            trade.ForwardRate = 1 / trade.ForwardRate;
                        }
                        // swap amounts and currencies
                        Decimal originalAmount = trade.Amount;
                        trade.Amount = trade.OtherAmount;
                        trade.OtherAmount = originalAmount;
                        string originalCurrency = trade.Currency;
                        trade.Currency = trade.OtherCurrency;
                        trade.OtherCurrency = originalCurrency;
                    }
                   
                    return trade;
                }

                return trade;
            }
            catch (Exception ex)
                {
                throw new ApplicationException(
                    "AdjustTrade Error: ", ex);
            }
        }

        public void Convert()
        {
            try
            {
                var dir = Path.GetDirectoryName(_outputFilePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                using (var writer = new StreamWriter(_outputFilePath, false, Encoding.ASCII))
                {
                    foreach (var trade in _trades)
                    {
                        var fields = ConvertTradeToAIMRow(trade);
                        writer.WriteLine(string.Join(",", fields));
                    }
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException("Error during conversion of NTFX trades to AIM format.", ex);
            }
        }

        public HashSet<string> ConvertWithAdjuster()
        {
            
            try
            {
                var dir = Path.GetDirectoryName(_outputFilePath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                using (var writer = new StreamWriter(_outputFilePath, false, Encoding.ASCII))
                {
                    foreach (var trade in _trades)
                    {
                        string implementationMsg = "";
                        var adjuster = _adjusterFactory.GetAdjuster(trade);
                        if (trade.Currency.Equals("EUR"))
                            trade.Currency = trade.Currency;
                        var adjustedTrade = adjuster.Adjust(trade);
                        // log adjuster usage
                        if (!adjuster.IsImplemented)
                            implementationMsg = "Not implemented yet. Trade pass through.";

                            _adjusterUsed.Add(adjuster.GetType().Name + " : "  + implementationMsg);
                        
                        var fields = ConvertTradeToAIMRow(adjustedTrade);
                        writer.WriteLine(string.Join(",", fields));
                    }
                }

                return _adjusterUsed;
            }
            catch (Exception ex)
            {
                throw new ApplicationException("ConverWithAdjuster: Error during conversion of NTFX trades to AIM format.", ex);
            }
        }

        private static string MapSide(string buySell)
        {
            if (string.Equals(buySell, "B", StringComparison.OrdinalIgnoreCase))
                return "by";

            if (string.Equals(buySell, "S", StringComparison.OrdinalIgnoreCase))
                return "sl";

            throw new ArgumentException("Unknown BuySell flag: " + buySell);
        }

        private static string FormatAmount(decimal amount)
        {
            decimal nonNegative = Math.Abs(amount);   // ensures positive value
            // No thousand separators, invariant culture, keeps decimals
            return nonNegative.ToString("0.##########", CultureInfo.InvariantCulture);
        }

        private static string FormatRate4Decimals(decimal rate)
        {
            decimal nonNegative = Math.Abs(rate);   // ensures positive value
            return nonNegative.ToString("0.####", CultureInfo.InvariantCulture);
        }

        private static string FormatRate6Decimals(decimal rate)
        {
            decimal nonNegative = Math.Abs(rate);
            decimal rounded = Math.Round(nonNegative, 6, MidpointRounding.AwayFromZero);
            //return rounded.ToString("0.000000", CultureInfo.InvariantCulture);
            return nonNegative.ToString("0.######", CultureInfo.InvariantCulture);
        }

        private static string FormatPreciseRate(decimal rate)
        {
            return rate.ToString("0.###############", CultureInfo.InvariantCulture);
        }


    }
}
