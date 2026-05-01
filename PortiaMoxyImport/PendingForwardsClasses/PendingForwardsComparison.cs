using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;

namespace PortiaMoxyImport.PendingForwardsClasses
{
    /// <summary>
    /// Compares Portia forwards data against Northern Trust PDF data for TB10 and TB20.
    /// </summary>
    public static class PendingForwardsComparison
    {
        // Currencies that use a wider tolerance for LocalAmt comparison
        private static readonly HashSet<string> WideToleanceCurrencies =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "KRW", "JPY" };

        private const decimal LocalAmtToleranceWide = 1.00m;
        private const decimal LocalAmtToleranceNarrow = 0.01m;
        private const decimal USDAmtTolerance = 0.01m;
        private const decimal ExRateTolerance = 0.01m;

        /// <summary>
        /// Runs the comparison for both portfolios and returns results grouped by portfolio.
        /// </summary>
        public static ComparisonResult Compare(
            DataTable portiaTB10,
            DataTable portiaTB20,
            DataTable pdfTB10,
            DataTable pdfTB20)
        {
            if (portiaTB10 == null) throw new ArgumentNullException("portiaTB10");
            if (portiaTB20 == null) throw new ArgumentNullException("portiaTB20");
            if (pdfTB10 == null) throw new ArgumentNullException("pdfTB10");
            if (pdfTB20 == null) throw new ArgumentNullException("pdfTB20");

            try
            {
                List<ComparisonRow> tb10Rows = ComparePortfolio(portiaTB10, pdfTB10, "55093");
                List<ComparisonRow> tb20Rows = ComparePortfolio(portiaTB20, pdfTB20, "55090");

                //return ComparisonResult.Ok(tb10Rows, tb20Rows);
                return ComparisonResult.Ok(tb10Rows, tb20Rows, pdfTB10, pdfTB20);
            }
            catch (Exception ex)
            {
                return ComparisonResult.Failure(
                    string.Format("Error during comparison: {0}", ex.Message));
            }
        }

        /// <summary>
        /// Compares a single portfolio's Portia rows against the corresponding PDF DataTable.
        /// Writes a diagnostics file to C:\Temp for inspection.
        /// </summary>
        private static List<ComparisonRow> ComparePortfolio(
            DataTable portiaData, DataTable pdfData, string portfolio)
        {
            List<ComparisonRow> results = new List<ComparisonRow>();
            List<string> diagnostics = new List<string>();

            foreach (DataRow portiaRow in portiaData.Rows)
            {
                string tranType = portiaRow["TranType"].ToString().Trim().ToLower();
                string currency = portiaRow["currency"].ToString().Trim().ToUpper();
                string broker = portiaRow["broker"].ToString().Trim();
                DateTime tradeDate = Convert.ToDateTime(portiaRow["trade_date"]);
                DateTime settleDate = Convert.ToDateTime(portiaRow["settle_date"]);
                decimal localAmt = Math.Abs(Convert.ToDecimal(portiaRow["LocalAmt"]));
                decimal usdAmt = Math.Abs(Convert.ToDecimal(portiaRow["USDAmt"]));
                decimal exchangeRate = Convert.ToDecimal(portiaRow["ExchangeRate"]);

                string tradeDateStr = tradeDate.ToString("MM/dd/yyyy");
                string settleDateStr = settleDate.ToString("MM/dd/yyyy");

                // Build a comparison row with Portia data populated
                ComparisonRow result = new ComparisonRow
                {
                    Portfolio = portfolio,
                    TranType = tranType,
                    Currency = currency,
                    TradeDate = tradeDate,
                    SettleDate = settleDate,
                    LocalAmt = localAmt,
                    USDAmt = usdAmt,
                    ExchangeRate = exchangeRate,
                    Broker = broker,
                    IsMatched = false
                };

                // Find matching PDF row, capturing the mismatch reason if not found
                string mismatchReason;
                DataRow pdfMatch = FindPdfMatchWithDiagnostics(
                    pdfData, tranType, currency,
                    tradeDateStr, settleDateStr,
                    localAmt, usdAmt, exchangeRate,
                    diagnostics,
                    out mismatchReason);

                if (pdfMatch != null)
                {
                    decimal pdfLocalAmt, pdfUsdAmt, pdfContractRate;
                    string pdfSettleDate;

                    if (tranType == "by")
                    {
                        // Buy: matched on Receivable side, USD amount is on Payable side
                        pdfLocalAmt = Math.Abs(GetDecimal(pdfMatch, "RecLocalAmount"));
                        pdfUsdAmt = Math.Abs(GetDecimal(pdfMatch, "PayLocalAmount"));
                        pdfContractRate = GetDecimal(pdfMatch, "RecContractRate");
                        pdfSettleDate = pdfMatch["RecSettleDate"].ToString();
                    }
                    else
                    {
                        // Sell: matched on Payable side, USD amount is on Receivable side
                        pdfLocalAmt = Math.Abs(GetDecimal(pdfMatch, "PayLocalAmount"));
                        pdfUsdAmt = Math.Abs(GetDecimal(pdfMatch, "RecLocalAmount"));
                        pdfContractRate = GetDecimal(pdfMatch, "PayContractRate");
                        pdfSettleDate = pdfMatch["PaySettleDate"].ToString();
                    }

                    result.PdfLocalAmt = pdfLocalAmt;
                    result.PdfUSDAmt = pdfUsdAmt;
                    result.PdfContractRate = pdfContractRate;
                    result.PdfSettleDate = pdfSettleDate;
                    result.LocalAmtVariance = localAmt - pdfLocalAmt;
                    result.USDAmtVariance = usdAmt - pdfUsdAmt;
                    result.ExRateVariance = exchangeRate - pdfContractRate;
                    result.IsMatched = true;
                    result.MismatchReason = string.Empty;
                }
                else
                {
                    result.MismatchReason = mismatchReason;
                }

                results.Add(result);
            }

            // Always write diagnostics file so we can inspect matches and failures
            string diagPath = Path.Combine(
                @"C:\Temp",
                string.Format("PendingForwards_Diag_{0}.txt", portfolio));
            System.IO.File.WriteAllLines(diagPath, diagnostics);

            // Sort: matched rows first, unmatched rows at the bottom
            results.Sort((a, b) => b.IsMatched.CompareTo(a.IsMatched));

            return results;
        }

        /// <summary>
        /// Finds a matching PDF row for a given Portia row, logging the reason
        /// for each failed match attempt to the diagnostics list.
        /// Returns the mismatch reason via out parameter when no match is found.
        /// </summary>
        private static DataRow FindPdfMatchWithDiagnostics(
            DataTable pdfData,
            string tranType,
            string currency,
            string tradeDateStr,
            string settleDateStr,
            decimal localAmt,
            decimal usdAmt,
            decimal exchangeRate,
            List<string> diagnostics,
            out string mismatchReason)
        {
            mismatchReason = string.Empty;

            // Determine which PDF columns to use based on transaction type
            string currencyCol = tranType == "by" ? "RecCurrency" : "PayCurrency";
            string tradeDateCol = tranType == "by" ? "RecTradeDate" : "PayTradeDate";
            string settleDateCol = tranType == "by" ? "RecSettleDate" : "PaySettleDate";
            string localAmtCol = tranType == "by" ? "RecLocalAmount" : "PayLocalAmount";
            string usdAmtCol = tranType == "by" ? "PayLocalAmount" : "RecLocalAmount";
            string contractRateCol = tranType == "by" ? "RecContractRate" : "PayContractRate";

            decimal localAmtTolerance = WideToleanceCurrencies.Contains(currency)
                ? LocalAmtToleranceWide
                : LocalAmtToleranceNarrow;

            // Key used in diagnostic messages to identify this Portia row
            string portiaKey = string.Format(
                "[{0} {1} TradeDate={2} SettleDate={3} LocalAmt={4} USDAmt={5} ExRate={6}]",
                tranType.ToUpper(), currency,
                tradeDateStr, settleDateStr,
                localAmt, usdAmt, exchangeRate);

            bool anyCurrencyMatch = false;
            bool anyDateMatch = false;

            foreach (DataRow pdfRow in pdfData.Rows)
            {
                string pdfCurrency = pdfRow[currencyCol].ToString().Trim();

                // Step 1: currency must match
                if (!string.Equals(pdfCurrency, currency, StringComparison.OrdinalIgnoreCase))
                    continue;

                anyCurrencyMatch = true;

                // Step 2: trade date must match — check both RecTradeDate and RecSettleDate
                // because PdfPig sometimes swaps these for certain PDF layouts
                string pdfTradeDate = pdfRow[tradeDateCol].ToString();
                string pdfSettleDate = pdfRow[settleDateCol].ToString();

                bool tradeDateMatch = pdfTradeDate == tradeDateStr;
                bool settleDateMatch = pdfSettleDate == settleDateStr;

                // Accept if dates match in either normal or swapped order
                bool datesOk = (tradeDateMatch && settleDateMatch)
                            || (pdfTradeDate == settleDateStr && pdfSettleDate == tradeDateStr);

                if (!tradeDateMatch && !datesOk)
                {
                    string msg = string.Format(
                        "TradeDate mismatch: PDF={0} expected={1}", pdfTradeDate, tradeDateStr);
                    diagnostics.Add(string.Format("{0} FAIL {1}", portiaKey, msg));
                    mismatchReason = msg;
                    continue;
                }

                if (!datesOk)
                {
                    string msg = string.Format(
                        "SettleDate mismatch: PDF={0} expected={1}", pdfSettleDate, settleDateStr);
                    diagnostics.Add(string.Format("{0} FAIL {1}", portiaKey, msg));
                    mismatchReason = msg;
                    continue;
                }

                anyDateMatch = true;

                anyDateMatch = true;

                // Step 4: LocalAmt within tolerance
                decimal pdfLocalAmt = Math.Abs(GetDecimal(pdfRow, localAmtCol));
                decimal localAmtDiff = Math.Abs(localAmt - pdfLocalAmt);
                if (localAmtDiff > localAmtTolerance)
                {
                    string msg = string.Format(
                        "LocalAmt mismatch: PDF={0} expected={1} diff={2} tolerance={3}",
                        pdfLocalAmt, localAmt, localAmtDiff, localAmtTolerance);
                    diagnostics.Add(string.Format("{0} FAIL {1}", portiaKey, msg));
                    mismatchReason = msg;
                    continue;
                }

                // Step 5: USDAmt within tolerance
                decimal pdfUsdAmt = Math.Abs(GetDecimal(pdfRow, usdAmtCol));
                decimal usdAmtDiff = Math.Abs(usdAmt - pdfUsdAmt);
                if (usdAmtDiff > USDAmtTolerance)
                {
                    string msg = string.Format(
                        "USDAmt mismatch: PDF={0} expected={1} diff={2} tolerance={3}",
                        pdfUsdAmt, usdAmt, usdAmtDiff, USDAmtTolerance);
                    diagnostics.Add(string.Format("{0} FAIL {1}", portiaKey, msg));
                    mismatchReason = msg;
                    continue;
                }

                // Step 6: ExchangeRate within tolerance
                decimal pdfContractRate = GetDecimal(pdfRow, contractRateCol);
                decimal exRateDiff = Math.Abs(exchangeRate - pdfContractRate);
                if (exRateDiff > ExRateTolerance)
                {
                    string msg = string.Format(
                        "ExRate mismatch: PDF={0} expected={1} diff={2} tolerance={3}",
                        pdfContractRate, exchangeRate, exRateDiff, ExRateTolerance);
                    diagnostics.Add(string.Format("{0} FAIL {1}", portiaKey, msg));
                    mismatchReason = msg;
                    continue;
                }

                // All criteria passed — match found
                mismatchReason = string.Empty;
                diagnostics.Add(string.Format("{0} MATCHED", portiaKey));
                return pdfRow;
            }

            // Log summary reason for no match
            if (!anyCurrencyMatch)
            {
                mismatchReason = string.Format(
                    "No PDF row found with currency '{0}' in column '{1}'",
                    currency, currencyCol);
            }
            else if (!anyDateMatch)
            {
                mismatchReason = "Currency matched but no PDF row passed date checks";
            }
            else
            {
                mismatchReason = "Dates matched but failed on numeric comparison";
            }

            diagnostics.Add(string.Format("{0} NO MATCH — {1}", portiaKey, mismatchReason));
            return null;
        }

        /// <summary>
        /// Safely extracts a decimal value from a DataRow, returning 0 if null or invalid.
        /// </summary>
        private static decimal GetDecimal(DataRow row, string columnName)
        {
            if (row[columnName] == DBNull.Value) return 0m;
            decimal result;
            return decimal.TryParse(
                row[columnName].ToString(),
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out result) ? result : 0m;
        }
    }
}