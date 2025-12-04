using CsvHelper.Configuration;
using PortiaMoxyImport.Entities;

namespace PortiaMoxyImport.Services
{
    internal sealed class NTFXTradeCsvRowMap : ClassMap<NTFXTradeCsvRow>
    {
        public NTFXTradeCsvRowMap()
        {
            Map(m => m.TradeDate)
                .Name("Trade date")
                .TypeConverterOption.Format("MM/dd/yyyy");

            Map(m => m.Account)
                .Name("Account");

            Map(m => m.BuySell)
                .Name("B/S");

            Map(m => m.Currency)
                .Name("Currency");

            Map(m => m.Amount)
                .Name("Amount");

            Map(m => m.OtherCurrency)
                .Name("Other Currency");

            Map(m => m.ForwardRate)
                .Name("Fwd Rate");

            Map(m => m.OtherAmount)
                .Name("Other Amt");

            Map(m => m.ValueDate)
                .Name("Val Date")
                .TypeConverterOption.Format("MM/dd/yyyy");

            Map(m => m.Broker)
                .Name("Broker");
        }
    }
}
