using PortiaMoxyImport.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Services
{
    public class BuyNonUsdUsdAdjuster : TradeAdjusterBase
    {
        private readonly List<string> _flipCurrencyList;

        public BuyNonUsdUsdAdjuster(List<string> flipCurrencyList)
        {
            _flipCurrencyList = flipCurrencyList;
        }
        protected override NTFXTradeDTO AdjustCore(NTFXTradeDTO trade)
        {
            string currency = trade.Currency.ToUpperInvariant();
            IsImplemented = true;
            decimal forwardRate;

            if (!_flipCurrencyList.Contains(trade.Currency))
            {
                                
                forwardRate =  trade.ForwardRate;
            }
            else
            {
                forwardRate = 1m/trade.ForwardRate;
            }

            // TODO: implement exact rule for Buy, base NON-USD, other USD
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
