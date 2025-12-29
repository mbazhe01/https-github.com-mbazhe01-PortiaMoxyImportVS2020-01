using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Entities
{

    /// <summary>
    /// This class represents a Data Transfer Object (DTO) for Nothrn Trust FX trades
    /// we download from their SFTP server.
    /// </summary>
    public class NTFXTradeDTO
    {
        public DateTime TradeDate { get; }
        public string Account { get; set; }
        public string BuySell { get; set; }          // "B" or "S"
        public string Currency { get; set; }
        public decimal Amount { get; set; }
        public string OtherCurrency { get; set; }
        public decimal ForwardRate { get; set; }
        public decimal OtherAmount { get; set; }
        public DateTime ValueDate { get; }
        public string Broker { get; }

        public NTFXTradeDTO(
            DateTime tradeDate,
            string account,
            string buySell,
            string currency,
            decimal amount,
            string otherCurrency,
            decimal forwardRate,
            decimal otherAmount,
            DateTime valueDate ,
            string broker )
        {
            TradeDate = tradeDate;
            Account = account;
            BuySell = buySell;
            Currency = currency;
            Amount = amount;
            OtherCurrency = otherCurrency;
            ForwardRate = forwardRate;
            OtherAmount = otherAmount;
            ValueDate = valueDate;
            Broker = broker;
        }


        public override String ToString()
        {
            return $"NTFXTradeDTO[TradeDate={TradeDate.ToShortDateString()}, Account={Account}, BuySell={BuySell}, Currency={Currency}, Amount={Amount}, OtherCurrency={OtherCurrency}, ForwardRate={ForwardRate}, OtherAmount={OtherAmount}, ValueDate={ValueDate.ToShortDateString()}, Broker={Broker}]";
        }

    }// end class
}
