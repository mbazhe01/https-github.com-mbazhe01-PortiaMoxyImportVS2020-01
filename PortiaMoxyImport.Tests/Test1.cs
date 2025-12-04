using PortiaMoxyImport.Services;


namespace PortiaMoxyImport.Tests
{
    [TestClass]
    public class NTFXTradesReaderTests
    {
        [TestMethod]
        public async Task GetTradesFromFileAsync_ParsesValidCsvCorrectly()
        {
            // Arrange: create a temp CSV file
            string tempFilePath = Path.GetTempFileName();

            string csvContent =
                "Trade date,Account,B/S,Currency,Amount,Other Currency,Fwd Rate,Other Amt,Val Date,Broker" + Environment.NewLine +
                "12/22/2025,ACC123,B,USD,100000,EUR,1.10,90000,12/24/2025,NTBROKER" + Environment.NewLine +
                "12/23/2025,ACC456,S,EUR,50000,USD,1.05,52000,12/26/2025,NTBROKER2" + Environment.NewLine;

            File.WriteAllText(tempFilePath, csvContent);

            var reader = new NTFXTradesReader(tempFilePath);

            try
            {
                // Act
                var trades = await reader.GetTradesFromFileAsync(tempFilePath);

                // Assert
                Assert.AreEqual(2, trades.Count, "Expected two trades to be parsed.");

                var first = trades[0];
                Assert.AreEqual(new DateTime(2025, 12, 22), first.TradeDate);
                Assert.AreEqual("ACC123", first.Account);
                Assert.AreEqual("B", first.BuySell);
                Assert.AreEqual("USD", first.Currency);
                Assert.AreEqual(100000m, first.Amount);
                Assert.AreEqual("EUR", first.OtherCurrency);
                Assert.AreEqual(1.10m, first.ForwardRate);
                Assert.AreEqual(90000m, first.OtherAmount);
                Assert.AreEqual(new DateTime(2025, 12, 24), first.ValueDate);
                Assert.AreEqual("NTBROKER", first.Broker);
            }
            finally
            {
                // Cleanup
                if (File.Exists(tempFilePath))
                {
                    File.Delete(tempFilePath);
                }
            }
        }
    }
}
