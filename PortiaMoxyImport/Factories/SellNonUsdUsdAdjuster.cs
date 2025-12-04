using PortiaMoxyImport.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Services
{
    public class SellNonUsdUsdAdjuster : TradeAdjusterBase
    {
        private readonly List<string> _flipCurrencyList;

        public SellNonUsdUsdAdjuster(List<string> flipCurrencyList)
        {
            _flipCurrencyList = flipCurrencyList;
        }

        protected override NTFXTradeDTO AdjustCore(NTFXTradeDTO trade)
        {
            // Normalize Sell to Buy by flipping legs & rate
            var buySell = "B";
            decimal forwardRate;

            // Base is NON-USD, Other is USD
            if (!_flipCurrencyList.Contains(trade.OtherCurrency))
            {
                forwardRate = 1m / trade.ForwardRate;
            }
            else
            {
                forwardRate = trade.ForwardRate;
            }

            // Swap legs
            var newAmount = trade.OtherAmount;
            var newOtherAmount = trade.Amount;
            var newCurrency = trade.OtherCurrency;
            var newOtherCurrency = trade.Currency;

            return new NTFXTradeDTO(
                trade.TradeDate,
                trade.Account,
                buySell,
                newCurrency,
                newAmount,
                newOtherCurrency,
                forwardRate,
                newOtherAmount,
                trade.ValueDate,
                trade.Broker);
        }
    }

}
