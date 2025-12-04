using PortiaMoxyImport.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Services
{
    public class BuyNonUsdNonUsdAdjuster : TradeAdjusterBase
    {
        protected override NTFXTradeDTO AdjustCore(NTFXTradeDTO trade)
        {
            // TODO: implement rule for Buy NON-USD / NON-USD
            return new NTFXTradeDTO(
                trade.TradeDate,
                trade.Account,
                "B",
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
