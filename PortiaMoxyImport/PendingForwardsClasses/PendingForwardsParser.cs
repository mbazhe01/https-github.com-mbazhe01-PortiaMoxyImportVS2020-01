

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace PortiaMoxyImport.Services
{
    public static class PendingForwardsParser
    {
        // TODO: take into considatration holiday calendar.


        private static readonly Regex InstrumentRegex =
            new Regex(@"^(USD/\w+\s+FWD\s+\d{8}\s+\w+)", RegexOptions.Compiled);

        private static readonly Regex DateRegex =
            new Regex(@"\b(\d{2}/\d{2}/\d{4})\b", RegexOptions.Compiled);

        // Matches a SecId — 16 char alphanumeric token that is not a date or known keyword
        private static readonly Regex SecIdRegex =
            new Regex(@"\b([A-Z0-9]{16})\b", RegexOptions.Compiled);

        private static readonly string[] SkipPrefixes = new[]
        {
        "FUND:",
        "As of Date:",
        "Pending Forwards Report",
        "Payable Receivable",
        "Description",
        "Local Amount",
        "Currency Contract Rate",
        "Ex-Rate Unrealized",
        "Trade Date Settle Date",
        "Total Base Cost",
        "Total Unrealized"
    };

        public static DataTable ParsePdf(string pdfPath, DateTime reportDate)
        {
            List<string> lines = ExtractLinesFromPdf(pdfPath);
            return ParseLinesToDataTable(lines, reportDate);
        }

        public static void DumpRawLines(string pdfPath, string outputPath)
        {
            List<string> lines = ExtractLinesFromPdf(pdfPath);
            System.IO.File.WriteAllLines(outputPath, lines);
        }

        private static DataTable ParseLinesToDataTable(List<string> lines, DateTime tradeDate)
        {
            string tradeDateStr = tradeDate.ToString("MM/dd/yyyy");

            DataTable dt = new DataTable("PendingForwards");
            dt.Columns.Add("Instrument", typeof(string));
            dt.Columns.Add("PayCurrency", typeof(string));
            dt.Columns.Add("PayLocalAmount", typeof(decimal));
            dt.Columns.Add("PayContractRate", typeof(decimal));
            dt.Columns.Add("PayExRate", typeof(decimal));
            dt.Columns.Add("PayTradeDate", typeof(string));
            dt.Columns.Add("PaySettleDate", typeof(string));
            dt.Columns.Add("PaySecId", typeof(string));
            dt.Columns.Add("RecCurrency", typeof(string));
            dt.Columns.Add("RecLocalAmount", typeof(decimal));
            dt.Columns.Add("RecContractRate", typeof(decimal));
            dt.Columns.Add("RecExRate", typeof(decimal));
            dt.Columns.Add("RecTradeDate", typeof(string));
            dt.Columns.Add("RecSettleDate", typeof(string));
            dt.Columns.Add("RecSecId", typeof(string));
            dt.Columns.Add("Broker", typeof(string));

            // Remove header/footer lines
            List<string> dataLines = new List<string>();
            foreach (string line in lines)
            {
                if (!ShouldSkip(line))
                    dataLines.Add(line);
            }

            // ADD HERE:
            System.IO.File.WriteAllLines(@"C:\Temp\tb10_datalines.txt", dataLines);

            string lastInstrument = string.Empty;
            int i = 0;

            while (i < dataLines.Count)
            {
                string currentLine = dataLines[i].Trim();
                Match match = InstrumentRegex.Match(currentLine);

                if (match.Success)
                {
                    // New instrument header found
                    lastInstrument = match.Groups[1].Value.Trim();
                    i++;
                    continue;
                }

                // Check if this line looks like the start of a data block
                // (a numbers line — starts with a negative number or digit)
                if (!string.IsNullOrWhiteSpace(lastInstrument) && IsNumbersLine(currentLine))
                {
                    // TEMP: log detected numbers lines
                    System.IO.File.AppendAllText(@"C:\Temp\numberslines.txt",
                        currentLine.Substring(0, Math.Min(60, currentLine.Length)) + "\r\n");

                    // Collect this block: numbers line, currency line, date line(s)
                    List<string> block = new List<string>();
                    int j = i;
                    while (j < dataLines.Count
                        && !InstrumentRegex.IsMatch(dataLines[j].Trim())
                        && block.Count < 6)
                    {
                        string l = dataLines[j].Trim();
                        if (!string.IsNullOrWhiteSpace(l))
                            block.Add(l);
                        j++;
                    }

                    if (block.Count >= 2)
                    {
                        DataRow row = BuildDataRow(dt, lastInstrument, block);
                        //if (row != null
                        //    && row["PayTradeDate"].ToString() == tradeDateStr
                        //    && row["RecTradeDate"].ToString() == tradeDateStr)
                        //{
                        //    dt.Rows.Add(row);
                        //}
                        if(row != null)
                        {
                            dt.Rows.Add(row);
                        }
                       
                    }

                    i = j;
                }
                else
                {
                    i++;
                }
            }

            return dt;
        }

        /// <summary>
        /// Returns true if the line looks like a numbers line —
        /// starts with a negative sign or digit, indicating the start of a trade data block.
        /// </summary>
        private static bool IsNumbersLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line)) return false;
            string trimmed = line.TrimStart();
            return trimmed.StartsWith("-") || char.IsDigit(trimmed[0]);
        }

        private static DataRow BuildDataRow(DataTable dt, string instrument, List<string> block)
        {
            if (dt == null) throw new ArgumentNullException(nameof(dt));
            if (string.IsNullOrWhiteSpace(instrument)) throw new ArgumentNullException(nameof(instrument));
            if (block == null) throw new ArgumentNullException(nameof(block));

            try
            {
                string numbersLine = block.Count > 0 ? block[0] : "";
                string currencyLine = block.Count > 1 ? block[1] : "";

                // Combine remaining lines for date extraction
                string dateBlock = "";
                for (int k = 2; k < block.Count; k++)
                    dateBlock += " " + block[k];
                dateBlock = dateBlock.Trim();

                // --- Numbers line ---
                // Fix smashed negatives: "000.00-7,956" -> "000.00 -7,956"
                string cleanedNums = Regex.Replace(numbersLine, @"(\d)(-)", "$1 $2");
                string[] nums = SplitTokens(cleanedNums);

                // Payable: first 3 values, Receivable: next 3 values
                decimal? payLocalAmt = ParseDecimal(nums.ElementAtOrDefault(0) ?? "");
                // nums[1] = PayCurrentMarketValue (not stored)
                // nums[2] = PayBaseCost (not stored)
                decimal? recLocalAmt = ParseDecimal(nums.ElementAtOrDefault(3) ?? "");
                // nums[4] = RecCurrentMarketValue (not stored)
                // nums[5] = RecBaseCost (not stored)

                // --- Currency line ---
                // Format: PayCurrency PayContractRate PayExRate PayUnrealizedGL RecCurrency RecContractRate RecExRate RecUnrealizedGL
                string[] cur = SplitTokens(currencyLine);
                string payCurrency = cur.ElementAtOrDefault(0) ?? "";
                decimal? payContractRate = ParseDecimal(cur.ElementAtOrDefault(1) ?? "");
                decimal? payExRate = ParseDecimal(cur.ElementAtOrDefault(2) ?? "");
                // cur[3] = PayUnrealizedGL (not stored)
                string recCurrency = cur.ElementAtOrDefault(4) ?? "";
                decimal? recContractRate = ParseDecimal(cur.ElementAtOrDefault(5) ?? "");
                decimal? recExRate = ParseDecimal(cur.ElementAtOrDefault(6) ?? "");
                // cur[7] = RecUnrealizedGL (not stored)

                // --- Date block ---
                // Three layout patterns exist depending on PDF column ordering:
                //
                // Pattern A: PayTD SecId NT PaySD NT RecTD RecSD SecId  → allDates=[PayTD,PaySD,RecTD,RecSD] correct
                // Pattern B: SecId NT PayTD PaySD NT RecSD SecId RecTD  → allDates=[PayTD,PaySD,RecSD,RecTD] swap Rec
                // Pattern C: PayTD SecId NT PaySD NT RecSD SecId RecTD  → allDates=[PayTD,PaySD,RecSD,RecTD] swap Rec
                //
                // Reliable detection: count dates before the first "NORTHERN TRUST".
                // Pattern A has exactly 1 date before first NT; Patterns B and C have 0 or 2.
                List<string> allDates = new List<string>();
                foreach (Match m in DateRegex.Matches(dateBlock))
                    allDates.Add(m.Value);

                List<string> allSecIds = new List<string>();
                foreach (Match m in SecIdRegex.Matches(dateBlock))
                    allSecIds.Add(m.Value);

                int ntIndex = dateBlock.IndexOf("NORTHERN TRUST");
                string beforeNT = ntIndex >= 0 ? dateBlock.Substring(0, ntIndex) : dateBlock;
                int datesBeforeNT = DateRegex.Matches(beforeNT).Count;

                // Pattern B: SecId leads Pay side (0 dates before first NT)
                // Pattern C: Pay side has date then SecId then NT then SettleDate — 1 date before NT but Rec pair still swapped
                // Detect Pattern C by checking if rec dates are in descending order (SettleDate before TradeDate)
                bool recPairSwapped = datesBeforeNT != 1;

                // Additional check for Pattern C: even with 1 date before NT,
                // if allDates[2] > allDates[3] (settle before trade), the pair is swapped
                if (!recPairSwapped && allDates.Count >= 4)
                {
                    if (DateTime.TryParse(allDates[2], out DateTime d2) &&
                        DateTime.TryParse(allDates[3], out DateTime d3) &&
                        d2 > d3)
                    {
                        recPairSwapped = true;
                    }
                }

                string payTradeDate = allDates.ElementAtOrDefault(0) ?? "";
                string paySettleDate = allDates.ElementAtOrDefault(1) ?? "";
                string recTradeDate = recPairSwapped
                    ? (allDates.ElementAtOrDefault(3) ?? "")
                    : (allDates.ElementAtOrDefault(2) ?? "");
                string recSettleDate = recPairSwapped
                    ? (allDates.ElementAtOrDefault(2) ?? "")
                    : (allDates.ElementAtOrDefault(3) ?? "");

                string paySecId = allSecIds.ElementAtOrDefault(0) ?? "";
                string recSecId = allSecIds.ElementAtOrDefault(1) ?? "";

                string broker = dateBlock.Contains("NORTHERN TRUST") ? "NORTHERN TRUST" : "";

                DataRow row = dt.NewRow();
                row["Instrument"] = instrument;
                row["PayCurrency"] = payCurrency;
                row["PayLocalAmount"] = payLocalAmt.HasValue ? (object)payLocalAmt.Value : DBNull.Value;
                row["PayContractRate"] = payContractRate.HasValue ? (object)payContractRate.Value : DBNull.Value;
                row["PayExRate"] = payExRate.HasValue ? (object)payExRate.Value : DBNull.Value;
                row["PayTradeDate"] = payTradeDate;
                row["PaySettleDate"] = paySettleDate;
                row["PaySecId"] = paySecId;
                row["RecCurrency"] = recCurrency;
                row["RecLocalAmount"] = recLocalAmt.HasValue ? (object)recLocalAmt.Value : DBNull.Value;
                row["RecContractRate"] = recContractRate.HasValue ? (object)recContractRate.Value : DBNull.Value;
                row["RecExRate"] = recExRate.HasValue ? (object)recExRate.Value : DBNull.Value;
                row["RecTradeDate"] = recTradeDate;
                row["RecSettleDate"] = recSettleDate;
                row["RecSecId"] = recSecId;
                row["Broker"] = broker;
                return row;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"BuildDataRow failed parsing block for instrument [{instrument}]: {ex.Message}", ex);
            }
        }

        private static List<string> ExtractLinesFromPdf(string pdfPath)
        {
            List<string> lines = new List<string>();

            using (PdfDocument doc = PdfDocument.Open(pdfPath))
            {
                foreach (Page page in doc.GetPages())
                {
                    IEnumerable<IGrouping<double, Word>> wordsByLine = page.GetWords()
                        .GroupBy(w => Math.Round(w.BoundingBox.Bottom, 0))
                        .OrderByDescending(g => g.Key);

                    foreach (IGrouping<double, Word> lineGroup in wordsByLine)
                    {
                        string lineText = string.Join(" ", lineGroup
                            .OrderBy(w => w.BoundingBox.Left)
                            .Select(w => w.Text));

                        if (!string.IsNullOrWhiteSpace(lineText))
                            lines.Add(lineText);
                    }
                }
            }

            return lines;
        }

        private static bool ShouldSkip(string line)
        {
            string trimmed = line.Trim();
            if (string.IsNullOrWhiteSpace(trimmed)) return true;

            foreach (string prefix in SkipPrefixes)
            {
                if (trimmed.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            // Skip standalone page lines like "Page 1 of 15"
            if (Regex.IsMatch(trimmed, @"^Page\s+\d+\s+of\s+\d+$", RegexOptions.IgnoreCase))
                return true;

            return false;
        }

        private static string[] SplitTokens(string s)
        {
            return s.Trim().Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static decimal? ParseDecimal(string s)
        {
            s = s.Replace(",", "").Trim();
            decimal d;
            if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                return d;
            return (decimal?)null;
        }
    }

}
