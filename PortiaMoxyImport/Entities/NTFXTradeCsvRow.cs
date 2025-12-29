using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Entities
{
    internal class NTFXTradeCsvRow
    {
        public DateTime TradeDate { get; set; }
        public string Account { get; set; }
        public string BuySell { get; set; }
        public string Currency { get; set; }
        public decimal Amount { get; set; }
        public string OtherCurrency { get; set; }
        public decimal ForwardRate { get; set; }
        public decimal OtherAmount { get; set; }
        public DateTime ValueDate { get; set; }
        public string Broker { get; set; }

        public override string ToString()
        {
            return $"{TradeDate.ToShortDateString()}, {Account}, {BuySell}, {Currency}, {Amount}, {OtherCurrency}, {ForwardRate}, {OtherAmount}, {ValueDate.ToShortDateString()}, {Broker}";
        }
    }
}
