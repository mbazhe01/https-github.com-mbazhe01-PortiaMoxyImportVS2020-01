using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;


namespace PortiaMoxyImport.HedgeExposureClasses
{
    public class HedgeExposureReader
    {

        public MoxyDatabase database;

        public HedgeExposureReader(MoxyDatabase database) { this.database = database; }


        /// <summary>
        /// read hedge exposure data from a Northern Trust csv file and return a list of HedgeExposureDto objects.
        /// The method handles quoted fields that may contain commas and parses dates in the dd/MM/yyyy format. 
        /// If the file does not exist, it returns an empty list.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public List<HedgeExposureDto> ReadFile(string filePath, List<NTPortiaPortDto> ntPortiaPortMap)
        {
            var results = new List<HedgeExposureDto>();

            // Ensure the file exists before attempting to read
            if (!File.Exists(filePath)) return results;

            // Use ReadLines to be memory efficient for larger files
            var lines = File.ReadLines(filePath);

            // Skip the header row
            var dataRows = lines.Skip(1);

            foreach (var line in dataRows)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;

                // This Regex splits by comma but ignores commas inside double quotes
                string[] columns = Regex.Split(line, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

                if (columns.Length >= 16)
                {
                    string ntAcctId = columns[1];
                    string portiaPort = ntPortiaPortMap.FirstOrDefault(x => x.NTAccttId == ntAcctId)?.PortiaPort ?? "U/A";

                    string ledgerDateStr = columns[16].Trim(); // "3/13/2026 0:00"
                    var ok = DateTime.TryParseExact(
                        ledgerDateStr,
                        "MM/dd/yyyy",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out var ledgerDate);

                    var totalLocalHedgeExposure = ParseDecimal(columns[2]);
                    var totalBaseMTM = ParseDecimal(columns[14]);
                    var baseAmountToBeAdjusted = totalLocalHedgeExposure - totalBaseMTM;

                    results.Add(new HedgeExposureDto
                    {

                        AccountName = columns[0].Trim('"'),
                        // : replace NT account id with Portia portfolio
                        AccountId = portiaPort,
                        TotalBaseHedgeExposure = ParseDecimal(columns[2]),
                        TotalLocalHedgeExposure = ParseDecimal(columns[3]),
                        TotalLocalHedgeTrades = ParseDecimal(columns[4]),
                        TotalBaseHedgeTrades = ParseDecimal(columns[5]),
                        BaseCurrency = columns[6],
                        LocalCurrencyCode = columns[7],
                        TargetHedgeRatio = ParseDecimal(columns[8]),
                        HedgeRatioLowerBound = ParseDecimal(columns[9]),
                        HedgeRatioUpperBound = ParseDecimal(columns[10]),
                        HedgeRatio = ParseDecimal(columns[11]),

                        // Handles the dd/MM/yyyy format (e.g., 17/02/2026)
                        LedgerDate = ledgerDate,

                        ValidationStatus = columns[13],
                        TotalBaseMtm = ParseDecimal(columns[14]),
                        BaseAmountToBeAdjusted = baseAmountToBeAdjusted,
                    });
                }
            }

            return results;
        }



        /// <summary>
        /// read hedge exposure data from a Northern Trust csv file and return a list of HedgeExposureDto objects.
        /// The method handles quoted fields that may contain commas and parses dates in the dd/MM/yyyy format. 
        /// If the file does not exist, it returns an empty list.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>

        public List<PortiaHdgExposureDto> ReadPortiaFile(string filePath)
        {
            var results = new List<PortiaHdgExposureDto>();

            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Portia file was not found.", filePath);
            }

            ValidateCommaSeparated(filePath);

            var rowNumber = 1;

            foreach (var line in File.ReadLines(filePath).Skip(1))
            {
                rowNumber++;

                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                string[] columns = null;

                try
                {
                    columns = Regex.Split(line, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)");

                    for (int i = 0; i < columns.Length; i++)
                    {
                        columns[i] = UnwrapCsvField(columns[i]);
                    }

                    if (columns.Length < 7)
                    {
                        throw new FormatException(
                            "Expected at least 7 columns, but found " + columns.Length + ".");
                    }

                    results.Add(new PortiaHdgExposureDto
                    {
                        AsOfDate = DateTime.ParseExact(columns[0], "MM/dd/yy", CultureInfo.InvariantCulture),
                        Account = columns[1].Trim('"'),
                        Country = columns[2],
                        MarketValueStocks = ParseDecimal(columns[3]),
                        MarketValueForwards = ParseDecimal(columns[4]),
                        HedgeAmount = ParseDecimal(columns[5]),
                        Security = columns[6].Trim('"')
                    });
                }
                catch (Exception ex)
                {
                    var parsedColumns = columns == null
                        ? "Columns could not be parsed."
                        : string.Join(
                            Environment.NewLine,
                            columns.Select((value, index) => "[" + index + "] = " + value));

                    throw new FormatException(
                        "Error reading Portia file." + Environment.NewLine +
                        "File: " + filePath + Environment.NewLine +
                        "Row: " + rowNumber + Environment.NewLine +
                        "Raw line: " + line + Environment.NewLine +
                        "Parsed columns:" + Environment.NewLine +
                        parsedColumns,
                        ex);
                }
            }

            return results;
        }

        private static string UnwrapCsvField(string value)
        {
            if (string.IsNullOrEmpty(value))
                return value;

            value = value.Trim();

            if (value.Length >= 2 && value.StartsWith("\"") && value.EndsWith("\""))
            {
                value = value.Substring(1, value.Length - 2);
                value = value.Replace("\"\"", "\""); // handle escaped quotes
            }

            return value;
        }


        private decimal ParseDecimal(string value)
        {
            if (decimal.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
            {
                return result;
            }
            return 0m;
        }

        public static void ValidateCommaSeparated(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("CSV file not found.", filePath);
            }

            var lines = File.ReadLines(filePath)
                            .Where(l => !string.IsNullOrWhiteSpace(l))
                            .Take(10)
                            .ToList();

            if (lines.Count == 0)
            {
                throw new FormatException("CSV file is empty or contains only blank lines.");
            }

            // Simple delimiter detection
            int commaCount = lines.Sum(l => l.Count(c => c == ','));
            int semicolonCount = lines.Sum(l => l.Count(c => c == ';'));
            int tabCount = lines.Sum(l => l.Count(c => c == '\t'));

            char detectedDelimiter =
                new[] {
            (Delimiter: ',', Count: commaCount),
            (Delimiter: ';', Count: semicolonCount),
            (Delimiter: '\t', Count: tabCount)
                }
                .OrderByDescending(x => x.Count)
                .First().Delimiter;

            if (detectedDelimiter != ',')
            {
                throw new FormatException(
                    $"File does not appear to be comma-separated. Detected '{detectedDelimiter}' as the delimiter.");
            }

            int? expectedColumns = null;
            int rowNumber = 0;

            foreach (var line in lines)
            {
                rowNumber++;

                var columns = Regex.Split(line, ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)")
                                   .Select(UnwrapCsvField)
                                   .ToList();

                // 🔑 Fix: remove trailing empty column(s) caused by trailing commas
                while (columns.Count > 0 && string.IsNullOrWhiteSpace(columns.Last()))
                {
                    columns.RemoveAt(columns.Count - 1);
                }

                if (expectedColumns == null)
                {
                    expectedColumns = columns.Count;
                }
                else if (columns.Count != expectedColumns)
                {
                    throw new FormatException(
                        $"Inconsistent column count detected. Expected {expectedColumns}, found {columns.Count} at row {rowNumber}." +
                        Environment.NewLine +
                        $"Line: {line}");
                }
            }
        }
    }
}