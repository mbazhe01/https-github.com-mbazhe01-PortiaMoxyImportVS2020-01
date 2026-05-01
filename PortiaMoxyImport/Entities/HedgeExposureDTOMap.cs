using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static PortiaMoxyImport.Entities.HedgeExposureDTO;

namespace PortiaMoxyImport.Entities
{
    public class HedgeExposureDTOMap : ClassMap<HedgeExposureDto>
    {
        public HedgeExposureDTOMap()
        {
            Map(m => m.AccountName).Name("account name");
            Map(m => m.AccountId).Name("account id");
            Map(m => m.TotalBaseHedgeExposure).Name("total base hedge exposure");
            Map(m => m.TotalLocalHedgeExposure).Name("total local hedge exposure");
            Map(m => m.TotalLocalHedgeTrades).Name("total local hedge trades");
            Map(m => m.TotalBaseHedgeTrades).Name("total base hedge trades");
            Map(m => m.BaseCurrency).Name("base currency");
            Map(m => m.LocalCurrencyCode).Name("local currency code");
            Map(m => m.TargetHedgeRatio).Name("target hedge ratio");
            Map(m => m.HedgeRatioLowerBound).Name("hedge ratio lower bound");
            Map(m => m.HedgeRatioUpperBound).Name("hedge ratio upper bound");
            Map(m => m.HedgeRatio).Name("hedge ratio");
            Map(m => m.LedgerDate)
                .Name("ledger date")
                .TypeConverterOption
                .Format("dd/MM/yyyy");
            Map(m => m.ValidationStatus).Name("validation_status");
            Map(m => m.TotalBaseMtm).Name("total base MTM");
            Map(m => m.BaseAmountToBeAdjusted).Name("base amount to be adjusted");
        }
    }
}
