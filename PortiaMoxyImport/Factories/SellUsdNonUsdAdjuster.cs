using PortiaMoxyImport.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Services
{
    public class SellUsdNonUsdAdjuster : TradeAdjusterBase
    {
        protected override NTFXTradeDTO AdjustCore(NTFXTradeDTO trade)
        {
            IsImplemented = false;
            // TODO: implement rule for Sell, base USD, other NON-USD
            return new NTFXTradeDTO(
                trade.TradeDate,
                trade.Account,
                "S",
                trade.Currency,
                trade.Amount,
                trade.OtherCurrency,
                trade.ForwardRate,
                trade.OtherAmount,
                trade.ValueDate,
                trade.Broker);
        }
    }
}
