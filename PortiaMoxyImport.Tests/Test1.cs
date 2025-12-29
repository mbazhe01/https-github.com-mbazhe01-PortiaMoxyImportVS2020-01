using Microsoft.VisualStudio.TestTools.UnitTesting;
using PortiaMoxyImport.Entities;
using PortiaMoxyImport.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Linq;

namespace PortiaMoxyImport.Tests
{
    [TestClass]
    public class NTFXTradesReaderTests
    {
       
        
            [TestMethod]
            public void SanityTest_ShouldRun()
            {
                Assert.AreEqual(1, 1);
            }
        [TestMethod]
        public async Task GetTradesFromFileAsync_ParsesValidCsvCorrectly()
        {
            // Arrange: create a temp CSV file
            string tempFilePath = Path.Combine(
                Path.GetTempPath(),
                "NTFX_Test_" + Guid.NewGuid().ToString("N") + ".csv");

            string csvContent =
                "Trade date,Account,B/S,Currency,Amount,Other Currency,Fwd Rate,Other Amt,Val Date,Broker" + Environment.NewLine +
                "12/22/2025,ACC123,B,USD,100000,EUR,1.10,90000,12/24/2025,NTBROKER" + Environment.NewLine +
                "12/23/2025,ACC456,S,EUR,50000,USD,1.05,52000,12/26/2025,NTBROKER2" + Environment.NewLine;

            File.WriteAllText(tempFilePath, csvContent);

            var reader = new NTFXTradesReader(tempFilePath);

            try
            {
                // Act
                List<NTFXTradeDTO> trades = await reader.GetTradesFromFileAsync();

                // Assert
                Assert.IsNotNull(trades, "Trades list should not be null.");
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

                var second = trades[1];
                Assert.AreEqual(new DateTime(2025, 12, 23), second.TradeDate);
                Assert.AreEqual("ACC456", second.Account);
                Assert.AreEqual("S", second.BuySell);
                Assert.AreEqual("EUR", second.Currency);
                Assert.AreEqual(50000m, second.Amount);
                Assert.AreEqual("USD", second.OtherCurrency);
                Assert.AreEqual(1.05m, second.ForwardRate);
                Assert.AreEqual(52000m, second.OtherAmount);
                Assert.AreEqual(new DateTime(2025, 12, 26), second.ValueDate);
                Assert.AreEqual("NTBROKER2", second.Broker);
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

        [TestMethod]
        public async Task GetTradesFromFileAsync_FileNotFound_ThrowsApplicationException()
        {
            // Arrange
            string invalidFilePath = Path.Combine(
                Path.GetTempPath(),
                "definitely_does_not_exist_" + Guid.NewGuid().ToString("N") + ".csv");

            var reader = new NTFXTradesReader(invalidFilePath);

            // Act + Assert
            var ex = await Assert.ThrowsExceptionAsync<ApplicationException>(
                async () => await reader.GetTradesFromFileAsync());

            // Inner exception should be FileNotFoundException
            Assert.IsInstanceOfType(ex.InnerException, typeof(FileNotFoundException));
            StringAssert.Contains(ex.Message, "Error reading NTFX trades from file");
        }

        [TestMethod]
        public async Task GetTradesFromFileAsync_RowWithMissingAccount_ThrowsApplicationExceptionWithValidationMessage()
        {
            // Arrange: create a temp CSV file with a missing Account on the second row
            string tempFilePath = Path.Combine(
                Path.GetTempPath(),
                "NTFX_Test_" + Guid.NewGuid().ToString("N") + ".csv");

            string csvContent =
                 "Trade date,Account,B/S,Currency,Amount,Other Currency,Fwd Rate,Other Amt,Val Date,Broker" + Environment.NewLine +
                 "12/22/2025,ACC123,B,USD,100000,EUR,1.10,90000,12/24/2025,NTBROKER" + Environment.NewLine +
                 "12/23/2025,,S,EUR,50000,USD,1.05,52000,12/26/2025,NTBROKER2" + Environment.NewLine;

            File.WriteAllText(tempFilePath, csvContent);

            var reader = new NTFXTradesReader(tempFilePath);

            try
            {
                // Act + Assert: validation should fail due to empty Account
                var ex = await Assert.ThrowsExceptionAsync<ApplicationException>(
                    async () => await reader.GetTradesFromFileAsync());

                // Outer wrapper thrown by GetTradesFromFileAsync
                StringAssert.Contains(ex.Message, "Error reading NTFX trades from file");

                // Inner exception comes from validateNTFXTradeCsvRow
                Assert.IsNotNull(ex.InnerException, "Inner exception should not be null.");
                Assert.IsInstanceOfType(ex.InnerException, typeof(ApplicationException));

                StringAssert.Contains(
                    ex.InnerException.Message,
                    "Invalid account",
                    "Validation message should indicate invalid account.");
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

        // ---------------------------------------
        //  FACTORY TEST HELPERS
        // ---------------------------------------

        private NTFXTradeDTO CreateSampleTrade(string side, string baseCurrency, string otherCurrency)
        {
            return new NTFXTradeDTO(
                tradeDate: DateTime.Today,
                account: "ACC1",
                buySell: side,
                currency: baseCurrency,
                amount: 100m,
                otherCurrency: otherCurrency,
                forwardRate: 1.25m,
                otherAmount: 125m,
                valueDate: DateTime.Today.AddDays(2),
                broker: "BRK"
            );
        }

        private TradeAdjusterFactory CreateFactory()
        {
            // Real adjusters, same as production
            var flipCurrencyList = new List<string>();

            return new TradeAdjusterFactory(
                buyUsdNonUsdAdjuster: new BuyUsdNonUsdAdjuster(flipCurrencyList),   // IsImplemented = true :contentReference[oaicite:1]{index=1}
                buyNonUsdUsdAdjuster: new BuyNonUsdUsdAdjuster(flipCurrencyList),                    // IsImplemented = false :contentReference[oaicite:2]{index=2}
                sellUsdNonUsdAdjuster: new SellUsdNonUsdAdjuster(),                   // IsImplemented = false :contentReference[oaicite:3]{index=3}
                sellNonUsdUsdAdjuster: new SellNonUsdUsdAdjuster(flipCurrencyList),   // IsImplemented = true :contentReference[oaicite:4]{index=4}
                buyNonUsdNonUsdAdjuster: new BuyNonUsdNonUsdAdjuster(flipCurrencyList),                 // IsImplemented = false :contentReference[oaicite:5]{index=5}
                sellNonUsdNonUsdAdjuster: new SellNonUsdNonUsdAdjuster()                 // IsImplemented = false :contentReference[oaicite:6]{index=6}
            );
        }

        // Each DataRow: side, baseCcy, otherCcy, expectNotImplementedMessage, expectedAdjusterTypeName
        [DataTestMethod]
        [DataRow("B", "USD", "EUR", false, "BuyUsdNonUsdAdjuster")]      // Implemented
        [DataRow("B", "EUR", "USD", true, "BuyNonUsdUsdAdjuster")]      // Not implemented
        [DataRow("S", "USD", "EUR", true, "SellUsdNonUsdAdjuster")]     // Not implemented
        [DataRow("S", "EUR", "USD", false, "SellNonUsdUsdAdjuster")]     // Implemented
        [DataRow("B", "EUR", "GBP", true, "BuyNonUsdNonUsdAdjuster")]   // Not implemented
        [DataRow("S", "EUR", "GBP", true, "SellNonUsdNonUsdAdjuster")]  // Not implemented
        public void FactoryLoop_Should_Log_Expected_ImplementationMsg(
            string side,
            string baseCurrency,
            string otherCurrency,
            bool expectNotImplementedMsg,
            string expectedAdjusterTypeName)
        {
            // Arrange
            var factory = CreateFactory();
            var trade = CreateSampleTrade(side, baseCurrency, otherCurrency);
            var adjusterUsed = new List<string>();

            // This simulates your production foreach-loop for ONE trade
            var adjuster = factory.GetAdjuster(trade);
            var adjustedTrade = adjuster.Adjust(trade);

            string implementationMsg = string.Empty;

            if (!((TradeAdjusterBase)adjuster).IsImplemented) // property is on base class :contentReference[oaicite:8]{index=8}
            {
                implementationMsg = "Not implemented yet. Trade pass through.";
            }

            adjusterUsed.Add(adjuster.GetType().Name + " : " + implementationMsg);

            // Assert: we logged the correct adjuster
            Assert.AreEqual(1, adjusterUsed.Count);
            var logged = adjusterUsed.Single();

            StringAssert.StartsWith(logged, expectedAdjusterTypeName + " : ");

            if (expectNotImplementedMsg)
            {
                // For unimplemented adjusters we expect *that* message
                StringAssert.Contains(logged, "Not implemented yet. Trade pass through.");
            }
            else
            {
                // For implemented ones we expect that message NOT to appear
                Assert.IsFalse(
                    logged.Contains("Not implemented yet. Trade pass through."),
                    "Did not expect 'Not implemented yet' message for an implemented adjuster.");
            }
        }

    }//eof-class
}
