using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Entities
{
    public class HedgeExposureDTO
    {
        public sealed class HedgeExposureDto
        {
            public string AccountName { get; }
            public string AccountId { get; }
            public decimal TotalBaseHedgeExposure { get; }
            public decimal TotalLocalHedgeExposure { get; }
            public decimal TotalLocalHedgeTrades { get; }
            public decimal TotalBaseHedgeTrades { get; }
            public string BaseCurrency { get; }
            public string LocalCurrencyCode { get; }
            public decimal TargetHedgeRatio { get; }
            public decimal HedgeRatioLowerBound { get; }
            public decimal HedgeRatioUpperBound { get; }
            public decimal HedgeRatio { get; }
            public DateTime LedgerDate { get; }
            public string ValidationStatus { get; }
            public decimal TotalBaseMtm { get; }
            public decimal BaseAmountToBeAdjusted { get; }

            public HedgeExposureDto(
                string accountName,
                string accountId,
                decimal totalBaseHedgeExposure,
                decimal totalLocalHedgeExposure,
                decimal totalLocalHedgeTrades,
                decimal totalBaseHedgeTrades,
                string baseCurrency,
                string localCurrencyCode,
                decimal targetHedgeRatio,
                decimal hedgeRatioLowerBound,
                decimal hedgeRatioUpperBound,
                decimal hedgeRatio,
                DateTime ledgerDate,
                string validationStatus,
                decimal totalBaseMtm,
                decimal baseAmountToBeAdjusted)
            {
                AccountName = accountName ?? throw new ArgumentNullException(nameof(accountName));
                AccountId = accountId ?? throw new ArgumentNullException(nameof(accountId));
                BaseCurrency = baseCurrency ?? throw new ArgumentNullException(nameof(baseCurrency));
                LocalCurrencyCode = localCurrencyCode ?? throw new ArgumentNullException(nameof(localCurrencyCode));
                ValidationStatus = validationStatus ?? throw new ArgumentNullException(nameof(validationStatus));

                TotalBaseHedgeExposure = totalBaseHedgeExposure;
                TotalLocalHedgeExposure = totalLocalHedgeExposure;
                TotalLocalHedgeTrades = totalLocalHedgeTrades;
                TotalBaseHedgeTrades = totalBaseHedgeTrades;
                TargetHedgeRatio = targetHedgeRatio;
                HedgeRatioLowerBound = hedgeRatioLowerBound;
                HedgeRatioUpperBound = hedgeRatioUpperBound;
                HedgeRatio = hedgeRatio;
                LedgerDate = ledgerDate;
                TotalBaseMtm = totalBaseMtm;
                BaseAmountToBeAdjusted = baseAmountToBeAdjusted;
            }
        }
    }
}
