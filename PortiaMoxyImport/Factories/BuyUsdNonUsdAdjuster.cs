using PortiaMoxyImport.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Services
{
    public class BuyUsdNonUsdAdjuster : TradeAdjusterBase
    {
        private readonly List<string> _flipCurrencyList;

        public BuyUsdNonUsdAdjuster(List<string> flipCurrencyList)
        {
            _flipCurrencyList = flipCurrencyList;
        }

        protected override NTFXTradeDTO AdjustCore(NTFXTradeDTO trade)
        {
            IsImplemented = true;
            decimal forwardRate;

            // Your old logic: Buy, Currency = USD, Other = NONUSD
            // If other currency not in flip list, invert the rate
            if (!_flipCurrencyList.Contains(trade.OtherCurrency))
            {
                forwardRate = 1m / trade.ForwardRate;
            }
            else
            {
                forwardRate = trade.ForwardRate;
            }

            return new NTFXTradeDTO(
                trade.TradeDate,
                trade.Account,
                "B",
                trade.Currency,
                trade.Amount,
                trade.OtherCurrency,
                forwardRate,
                trade.OtherAmount,
                trade.ValueDate,
                trade.Broker);
        }

        
    }

}
