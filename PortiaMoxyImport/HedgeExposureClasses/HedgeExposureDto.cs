using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.HedgeExposureClasses
{
    public sealed class HedgeExposureDto
    {
        public string AccountName { get; set; }
        public string AccountId { get; set; }
        public decimal TotalBaseHedgeExposure { get; set; }
        public decimal TotalLocalHedgeExposure { get; set; }
        public decimal TotalLocalHedgeTrades { get; set; }
        public decimal TotalBaseHedgeTrades { get; set; }
        public string BaseCurrency { get; set; }
        public string LocalCurrencyCode { get; set; }
        public decimal TargetHedgeRatio { get; set; }
        public decimal HedgeRatioLowerBound { get; set; }
        public decimal HedgeRatioUpperBound { get; set; }
        public decimal HedgeRatio { get; set; }
        public DateTime LedgerDate { get; set; }
        public string ValidationStatus { get; set; }
        public decimal TotalBaseMtm { get; set; }
        public decimal BaseAmountToBeAdjusted { get; set; }
    }
}
