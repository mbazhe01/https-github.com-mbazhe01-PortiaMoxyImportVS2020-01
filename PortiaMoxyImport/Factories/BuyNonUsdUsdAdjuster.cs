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
        protected override NTFXTradeDTO AdjustCore(NTFXTradeDTO trade)
        {
            // TODO: implement exact rule for Buy, base NON-USD, other USD
            return new NTFXTradeDTO(
                trade.TradeDate,
                trade.Account,
                "B",
                trade.Currency,
                trade.Amount,
                trade.OtherCurrency,
                trade.ForwardRate,
                trade.Amount,
                trade.ValueDate,
                trade.Broker);
        }
    }
}
