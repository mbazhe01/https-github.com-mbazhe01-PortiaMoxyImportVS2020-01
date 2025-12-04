using PortiaMoxyImport.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Services
{
    public abstract class TradeAdjusterBase : ITradeAdjuster
    {
        public NTFXTradeDTO Adjust(NTFXTradeDTO trade)
        {
            if (trade == null) throw new ArgumentNullException(nameof(trade));
            if (trade.ForwardRate==null) throw new ApplicationException("ForwardRate is required.");
           

            return AdjustCore(trade);
        }

        protected abstract NTFXTradeDTO AdjustCore(NTFXTradeDTO trade);
    }
}
