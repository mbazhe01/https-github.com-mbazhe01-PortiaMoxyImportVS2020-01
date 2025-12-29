using PortiaMoxyImport.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Redesign
{
    /// <summary>
    /// redesigned PortiaMoxyManager class  
    /// </summary>
    internal class PortiaMoxyManagerR
    {
        private readonly IConversionReporter _reporter;

        public PortiaMoxyManagerR(IConversionReporter reporter)
        {
            _reporter = reporter ?? throw new ArgumentNullException(nameof(reporter));
        }

        internal void convertPortiaToMoxy(List<FileConversionDTO> fileConversions)
        {
            if (fileConversions == null)
                throw new ArgumentNullException(nameof(fileConversions));

            try
            {
                foreach (var conversion in fileConversions)
                {
                    if (conversion == null)
                    {
                        _reporter.Warning("Skipping null conversion entry in fileConversions.");
                        continue;
                    }

                    _reporter.Info(
                        $"Converting '{conversion.SourceFilePath}' → '{conversion.TargetFilePath}' as '{conversion.FileType}'.");

                    try
                    {
                        // This is the method we refactored earlier that does the per-file switch.
                        var result = ConvertToMoxy(conversion);

                        if (!result.Success)
                        {
                            var details = string.IsNullOrWhiteSpace(result.ErrorMessage)
                                ? "Unknown error."
                                : result.ErrorMessage;

                            _reporter.Error(
                                $"Conversion failed for '{conversion.SourceFilePath}': {details}");

                            // Decide: abort the whole batch, or just continue?
                            // Right now we log and continue with the rest.
                            continue;
                        }

                        _reporter.Success(
                            $"Conversion completed for '{conversion.SourceFilePath}'.");
                    }
                    catch (PortiaMoxyConversionException ex)
                    {
                        // Known conversion-level error – log and abort the batch (or continue if you prefer).
                        _reporter.Error(
                            $"Critical conversion error for '{conversion.SourceFilePath}': {ex.Message}");
                        throw; // rethrow to caller to signal batch failure
                    }
                    catch (Exception ex)
                    {
                        // Unexpected error – wrap in a domain-specific exception so callers know what failed.
                        _reporter.Error(
                            $"Unexpected error while converting '{conversion.SourceFilePath}': {ex.Message}");

                        throw new PortiaMoxyConversionException(
                            "Error in PortiaMoxyManager.convertPortiaToMoxy.",
                            ex);
                    }
                }
            }
            catch
            {
                // Let the exception bubble up after logging; no extra wrapping here
                // to avoid double-wrapping and losing stack clarity.
                throw;
            }
        }

        public ConversionResult ConvertToMoxy(FileConversionDTO file)
        {
            if (file is null)
                throw new ArgumentNullException(nameof(file));

            try
            {
                if (string.IsNullOrWhiteSpace(file.FileType))
                {
                    const string msg = "File type is required.";
                    _reporter.Error(msg);
                    return ConversionResult.Fail(msg);
                }

                if (string.IsNullOrWhiteSpace(file.SourceFilePath))
                {
                    const string msg = "Source file path is required.";
                    _reporter.Error(msg);
                    return ConversionResult.Fail(msg);
                }

                if (string.IsNullOrWhiteSpace(file.TargetFilePath))
                {
                    const string msg = "Target file path is required.";
                    _reporter.Error(msg);
                    return ConversionResult.Fail(msg);
                }

                var normalizedType = file.FileType.Trim().ToLowerInvariant();
                var inPath = file.SourceFilePath;
                var outPath = file.TargetFilePath;

                switch (normalizedType)
                {
                    case "holiday":
                        return WrapIntResult(convertHoliday24(file.SourceFilePath, file.TargetFilePath), "Holiday");

                    case "groups":
                        //return WrapIntResult(convertGroups24(inPath, outPath), "Groups");

                    case "price":
                        //return WrapIntResult(convertPrice24(inPath, outPath), "Price");

                    case "currency":
                        //return WrapIntResult(convertCurrency24(inPath, outPath), "Currency");

                    case "portfolio":
                        {
                            //var result = convertPortfolio24(inPath, outPath);
                            //if (result.Item1 == -1)
                            //{
                            //    const string msg = "Portfolio conversion failed.";
                            //    _reporter.Error(msg);
                            //    return ConversionResult.Fail(msg);
                            //}

                            //// Update the manager-level portfolio set for later taxlot processing
                            //if (result.Item2 != null)
                            //{
                            //    //_hsPortfolios.Clear();
                            //    //_hsPortfolios.UnionWith(result.Item2);
                            //}

                            _reporter.Success("Portfolio conversion completed successfully.");
                            return ConversionResult.Ok();
                        }

                    case "security":
                       // return WrapIntResult(convertSecurity24(inPath, outPath), "Security");

                    case "taxlot":
                        // Uses _hsPortfolios populated during portfolio conversion
                      //return WrapIntResult(convertTaxLots24(inPath, outPath, _hsPortfolios), "TaxLot");

                    case "custodian":
                       //eturn WrapIntResult(convertCustodian(inPath, outPath), "Custodian");

                    case "broker":
                       //eturn WrapIntResult(convertBrokers24(inPath, outPath), "Broker");

                    case "sector":
                        //return WrapIntResult(convertSectors(inPath, outPath), "Sector");

                    case "industry":
                       // return WrapIntResult(convertIndustry(inPath, outPath), "Industry");

                    case "sectype":
                       // return WrapIntResult(convertSecType24(inPath, outPath), "Security Type");

                    default:
                        {
                            var msg = $"Unknown file type '{file.FileType}'.";
                            _reporter.Warning(msg);
                            return ConversionResult.Fail(msg);
                        }
                }
            }
            catch (Exception ex)
            {
                // Preserve stack trace by using inner exception rather than rewriting the message only
                throw new PortiaMoxyConversionException(
                    "Error in PortiaMoxyManager.ConvertToMoxy.",
                    ex);
            }
        }

        private ConversionResult WrapIntResult(int result, string operationName)
        {
            if (result == -1)
            {
                var msg = $"{operationName} conversion failed.";
                _reporter.Error(msg);
                return ConversionResult.Fail(msg);
            }

            _reporter.Success($"{operationName} conversion completed successfully.");
            return ConversionResult.Ok();
        }

        public sealed class PortiaMoxyConversionException : Exception
        {
            public PortiaMoxyConversionException(string message, Exception innerException)
                : base(message, innerException)
            {
            }
        }

        public int convertGeneric(string inputFilePath, string outputFilePath)
        {
            //
            // Removes double quotes & replaces commas with tabs.
            // Returns number of processed rows or -1 on failure.
            //

            if (string.IsNullOrWhiteSpace(inputFilePath))
                return -1;

            if (string.IsNullOrWhiteSpace(outputFilePath))
                return -1;

            var processedRows = 0;
            var outputLines = new List<string>();

            try
            {
                foreach (var line in File.ReadLines(inputFilePath))
                {
                    // Skip empty or "No Data" rows
                    if (line.IndexOf("No Data", StringComparison.OrdinalIgnoreCase) >= 0)
                        continue;

                    // Transform
                    var normalized = line
                        .Replace("\"", string.Empty)
                        .Replace(",", "\t");

                    // Mandatory validation (same pattern as other converters)
                    if (checkMandatoryValues(normalized) == -1)
                        return -1;

                    outputLines.Add(normalized);
                    processedRows++;
                }

                File.WriteAllLines(outputFilePath, outputLines);
                return processedRows;
            }
            catch
            {
                return -1;
            }
        }

        public int checkMandatoryValues(string line)
        {
            //
            // Checks that the first two positions in the tab-separated string have values.
            // Returns 0 on success, -1 if a mandatory value is missing or an error occurs.
            //

            if (line == null)
            {
                _reporter.Error("checkMandatoryValues: input line is null.");
                return -1;
            }

            try
            {
                var values = line.Split('\t');

                // index 0 must exist and be non-empty
                if (values.Length == 0 || string.IsNullOrWhiteSpace(values[0]))
                {
                    _reporter.Error(
                        $"!!!---> checkMandatoryValues: value is missing at index 0 for line: {line}");
                    return -1;
                }

                // index 1 must exist and be non-empty if there is at least 2 columns
                if (values.Length > 1 && string.IsNullOrWhiteSpace(values[1]))
                {
                    _reporter.Error(
                        $"!!!---> checkMandatoryValues: value is missing at index 1 for line: {line}");
                    return -1;
                }

                // all mandatory checks passed
                return 0;
            }
            catch (Exception ex)
            {
                _reporter.Error($"checkMandatoryValues: {ex.Message}");
                return -1;
            }
        }

        public int convertHoliday24(string inputFilePath, string outputFilePath)
        {
            int rtn = 0;

            try
            {
                // Optional: still emit a logical header via reporter, instead of UI helper
                _reporter.Info("Holiday Conversion");

                // Validate that the file has the expected number of columns (7)
                if (!isValidColNumber(inputFilePath, 7))
                {
                    // Assume isValidColNumber logs details itself if needed
                    return -1;
                }

                // Delegate the actual line transformation to the generic converter
                rtn = convertGeneric(inputFilePath, outputFilePath);

                // IMPORTANT:
                // We no longer call ShowGreenText here; success/failure messaging
                // is handled centrally by WrapIntResult, which looks at rtn.
                // The (rtn - 1) “holidays loaded” message was just display logic
                // and doesn’t affect the contract (only -1 vs non-negative matters).

                return rtn;
            }
            catch (Exception ex)
            {
                // Log a detailed error, but keep the int contract for WrapIntResult
                _reporter.Error($"convertHoliday24: {ex.Message}");
                return -1;
            }
        }

        bool isValidColNumber(string filePath, int requiredColumnCount)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                {
                    _reporter.Error("isValidColNumber: file path is empty.");
                    return false;
                }

                // Read all lines one time
                var lines = File.ReadAllLines(filePath);

                if (lines == null || lines.Length == 0)
                {
                    _reporter.Error("isValidColNumber: file is empty: " + filePath);
                    return false;
                }

                // Try to find first actual data row (skip header)
                string dataLine = null;

                for (int i = 1; i < lines.Length; i++)
                {
                    var line = lines[i];

                    if (!string.IsNullOrWhiteSpace(line) &&
                        line.IndexOf("No Data", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        dataLine = line;
                        break;
                    }
                }

                if (dataLine == null)
                {
                    _reporter.Warning("isValidColNumber: no data rows found in file: " + filePath);
                    return false;
                }

                var columnCount = dataLine.Split(',').Length;

                if (columnCount != requiredColumnCount)
                {
                    _reporter.Error(
                        "isValidColNumber: column count mismatch. Required: " +
                        requiredColumnCount + ", Provided: " + columnCount);
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                _reporter.Error("isValidColNumber: " + ex.Message);
                return false;
            }
        }



    }// eo class
}
