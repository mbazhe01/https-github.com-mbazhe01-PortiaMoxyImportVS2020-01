using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Entity
{
    class MoxyTrade
    {
        private String portCode { get; set; }
        private String tranCode { get; set; }
        private String comment { get; set; }
        private String symbol { get; set; }

        private String tradeDate { get; set; }

        private String settleDate { get; set; }

        private String origCostDate { get; set; }

        private String quantity { get; set; }

        private String srcDstType { get; set; }

        private String srcDstSymbol { get; set; }

        private String tradeDateFXRate { get; set; }

        private String settleDateFXRate { get; set; }

        private String origFXRate { get; set; }

        private String tradeAmount { get; set; }

        private String userDef2 { get; set; }

        private String ipCounter { get; set; }

    }
}
