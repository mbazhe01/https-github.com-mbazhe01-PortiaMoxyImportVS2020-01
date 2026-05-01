using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace PortiaMoxyImport.PendingForwardsClasses
{
    /// <summary>
    /// Exports the pending forwards comparison results to an Excel file with four tabs:
    /// - 55090: comparison results for TB20
    /// - 55093: comparison results for TB10
    /// - PDF TB20: raw parsed data from TB20 PDF
    /// - PDF TB10: raw parsed data from TB10 PDF
    /// Uses Excel Interop with SafeReleaseCom patterns for proper COM object cleanup.
    /// </summary>
    public static class PendingForwardsExporter
    {
        public static ExportResult Export(
            ComparisonResult comparison,
            DateTime tradeDate,
            string outputFolder)
        {
            if (comparison == null)
                throw new ArgumentNullException("comparison");
            if (string.IsNullOrWhiteSpace(outputFolder))
                throw new ArgumentNullException("outputFolder");

            Excel.Application xlApp = null;
            Excel.Workbooks xlWorkbooks = null;
            Excel.Workbook xlWorkbook = null;

            try
            {
                if (!Directory.Exists(outputFolder))
                    return ExportResult.Failure(
                        string.Format("Output folder not found: {0}", outputFolder));

                // Build filename e.g. PendingForwards_Comparison_04232026_mbazhenov.xlsx
                string userName = System.Environment.UserName;
                string fileName = string.Format(
                    "PendingForwards_Comparison_{0}_{1}.xlsx",
                    tradeDate.ToString("MMddyyyy"),
                    userName);
                string filePath = Path.Combine(outputFolder, fileName);

                xlApp = new Excel.Application();
                xlApp.Visible = false;
                xlApp.DisplayAlerts = false;

                xlWorkbooks = xlApp.Workbooks;
                xlWorkbook = xlWorkbooks.Add();

                // Sheet 1: comparison results for portfolio 55090 (TB20)
                WriteComparisonTab(xlWorkbook, "55090", comparison.TB20Rows, 1);

                // Sheet 2: comparison results for portfolio 55093 (TB10)
                WriteComparisonTab(xlWorkbook, "55093", comparison.TB10Rows, 2);

                // Sheet 3: raw PDF data for TB20
                WritePdfDataTab(xlWorkbook, "PDF TB20", comparison.PdfTB20Data, 3);

                // Sheet 4: raw PDF data for TB10
                WritePdfDataTab(xlWorkbook, "PDF TB10", comparison.PdfTB10Data, 4);

                // Remove any extra default sheets Excel created
                while (xlWorkbook.Sheets.Count > 4)
                {
                    Excel.Worksheet xlExtra =
                        (Excel.Worksheet)xlWorkbook.Sheets[5];
                    xlExtra.Delete();
                    SafeReleaseCom(xlExtra);
                }

                xlWorkbook.SaveAs(
                    filePath,
                    Excel.XlFileFormat.xlOpenXMLWorkbook,
                    Type.Missing, Type.Missing, false, false,
                    Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing);

                return ExportResult.Ok(filePath);
            }
            catch (Exception ex)
            {
                return ExportResult.Failure(
                    string.Format("PendingForwardsExporter.Export: Error exporting comparison to Excel: {0}", ex.Message));
            }
            finally
            {
                if (xlApp != null)
                {
                    xlApp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                    xlApp.DisplayAlerts = true;
                }

                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(false);
                    SafeReleaseCom(xlWorkbook);
                }

                SafeReleaseCom(xlWorkbooks);

                if (xlApp != null)
                {
                    xlApp.Quit();
                    SafeReleaseCom(xlApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Writes a comparison results tab with Portia vs PDF data and variance columns.
        /// </summary>
        private static void WriteComparisonTab(
            Excel.Workbook xlWorkbook,
            string tabName,
            List<ComparisonRow> rows,
            int sheetIndex)
        {
            Excel.Worksheet xlSheet = null;
            Excel.Range xlHeaderRange = null;
            Excel.Range xlDataRange = null;
            Excel.Range xlFullRange = null;

            try
            {
                xlSheet = GetOrAddSheet(xlWorkbook, sheetIndex);
                xlSheet.Name = tabName;

                string[] headers = new[]
                {
                    "TranType", "Currency", "Trade Date", "Settle Date",
                    "Portia LocalAmt", "Portia USDAmt", "Portia ExRate",
                    "PDF LocalAmt", "PDF USDAmt", "PDF ContractRate", "PDF SettleDate",
                    "LocalAmt Variance", "USDAmt Variance", "ExRate Variance",
                    "Broker", "Status", "Mismatch Reason"
                };

                int colCount = headers.Length;
                int rowCount = rows.Count;

                // Write header row
                object[,] headerArray = new object[1, colCount];
                for (int c = 0; c < colCount; c++)
                    headerArray[0, c] = headers[c];

                xlHeaderRange = xlSheet.Range[
                    xlSheet.Cells[1, 1],
                    xlSheet.Cells[1, colCount]];
                xlHeaderRange.Value2 = headerArray;

                // Style header
                xlHeaderRange.Font.Bold = true;
                xlHeaderRange.Font.Color = ColorToOleColor(System.Drawing.Color.White);
                xlHeaderRange.Interior.Color = ColorToOleColor(
                    System.Drawing.ColorTranslator.FromHtml("#1A3A5C"));
                xlHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlHeaderRange.RowHeight = 25;

                if (rowCount > 0)
                {
                    // Build data array
                    object[,] dataArray = new object[rowCount, colCount];

                    for (int r = 0; r < rowCount; r++)
                    {
                        ComparisonRow row = rows[r];

                        dataArray[r, 0] = row.TranType.ToUpper();
                        dataArray[r, 1] = row.Currency;
                        dataArray[r, 2] = row.TradeDate.ToString("MM/dd/yyyy");
                        dataArray[r, 3] = row.SettleDate.ToString("MM/dd/yyyy");
                        dataArray[r, 4] = (object)row.LocalAmt;
                        dataArray[r, 5] = (object)row.USDAmt;
                        dataArray[r, 6] = (object)row.ExchangeRate;
                        dataArray[r, 7] = row.IsMatched ? (object)row.PdfLocalAmt.Value : "";
                        dataArray[r, 8] = row.IsMatched ? (object)row.PdfUSDAmt.Value : "";
                        dataArray[r, 9] = row.IsMatched ? (object)row.PdfContractRate.Value : "";
                        dataArray[r, 10] = row.IsMatched ? (object)row.PdfSettleDate : "";
                        dataArray[r, 11] = row.IsMatched ? (object)row.LocalAmtVariance.Value : "";
                        dataArray[r, 12] = row.IsMatched ? (object)row.USDAmtVariance.Value : "";
                        dataArray[r, 13] = row.IsMatched ? (object)row.ExRateVariance.Value : "";
                        dataArray[r, 14] = row.Broker;
                        dataArray[r, 15] = row.IsMatched ? "Matched" : "Unmatched";
                        dataArray[r, 16] = row.IsMatched ? "" : row.MismatchReason ?? "";
                    }

                    // Write data in one shot
                    xlDataRange = xlSheet.Range[
                        xlSheet.Cells[2, 1],
                        xlSheet.Cells[rowCount + 1, colCount]];
                    xlDataRange.Value2 = dataArray;

                    // Number formats
                    FormatColumn(xlSheet, rowCount, 5, "#,##0.00");     // Portia LocalAmt
                    FormatColumn(xlSheet, rowCount, 6, "#,##0.00");     // Portia USDAmt
                    FormatColumn(xlSheet, rowCount, 7, "#,##0.000000"); // Portia ExRate
                    FormatColumn(xlSheet, rowCount, 8, "#,##0.00");     // PDF LocalAmt
                    FormatColumn(xlSheet, rowCount, 9, "#,##0.00");     // PDF USDAmt
                    FormatColumn(xlSheet, rowCount, 10, "#,##0.000000"); // PDF ContractRate
                    FormatColumn(xlSheet, rowCount, 12, "#,##0.00");     // LocalAmt Variance
                    FormatColumn(xlSheet, rowCount, 13, "#,##0.00");     // USDAmt Variance
                    FormatColumn(xlSheet, rowCount, 14, "#,##0.000000"); // ExRate Variance

                    // Row highlighting: green = matched, red = unmatched
                    System.Drawing.Color greenColor =
                        System.Drawing.ColorTranslator.FromHtml("#C6EFCE");
                    System.Drawing.Color redColor =
                        System.Drawing.ColorTranslator.FromHtml("#FFC7CE");

                    for (int r = 0; r < rowCount; r++)
                    {
                        Excel.Range xlRow = null;
                        Excel.Range xlRowFull = null;
                        try
                        {
                            xlRow = (Excel.Range)xlSheet.Rows[r + 2];
                            xlRowFull = xlRow.EntireRow;
                            xlRowFull.Interior.Color = rows[r].IsMatched
                                ? ColorToOleColor(greenColor)
                                : ColorToOleColor(redColor);
                        }
                        finally
                        {
                            SafeReleaseCom(xlRowFull);
                            SafeReleaseCom(xlRow);
                        }
                    }
                }

                // Column widths
                int[] colWidths = new[]
                {
                    10, 10, 14, 14,  // TranType, Currency, Trade Date, Settle Date
                    16, 16, 16,      // Portia LocalAmt, USDAmt, ExRate
                    16, 16, 16, 14,  // PDF LocalAmt, USDAmt, ContractRate, SettleDate
                    18, 18, 16,      // Variances
                    16, 12, 40       // Broker, Status, Mismatch Reason
                };

                SetColumnWidths(xlSheet, colWidths);

                FreezeAndFilter(xlSheet, rowCount, colCount);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(
                    "PendingForwardsExporter.WriteComparisonTab: Error writing tab '{0}': {1}",
                    tabName, ex.Message));
            }
            finally
            {
                SafeReleaseCom(xlFullRange);
                SafeReleaseCom(xlDataRange);
                SafeReleaseCom(xlHeaderRange);
                SafeReleaseCom(xlSheet);
            }
        }

        /// <summary>
        /// Writes a raw PDF DataTable to a worksheet tab for inspection.
        /// </summary>
        private static void WritePdfDataTab(
            Excel.Workbook xlWorkbook,
            string tabName,
            DataTable dt,
            int sheetIndex)
        {
            Excel.Worksheet xlSheet = null;
            Excel.Range xlHeaderRange = null;
            Excel.Range xlDataRange = null;

            try
            {
                xlSheet = GetOrAddSheet(xlWorkbook, sheetIndex);
                xlSheet.Name = tabName;

                if (dt == null || dt.Columns.Count == 0)
                    return;

                int colCount = dt.Columns.Count;
                int rowCount = dt.Rows.Count;

                // Write header row from DataTable column names
                object[,] headerArray = new object[1, colCount];
                for (int c = 0; c < colCount; c++)
                    headerArray[0, c] = dt.Columns[c].ColumnName;

                xlHeaderRange = xlSheet.Range[
                    xlSheet.Cells[1, 1],
                    xlSheet.Cells[1, colCount]];
                xlHeaderRange.Value2 = headerArray;

                // Style header
                xlHeaderRange.Font.Bold = true;
                xlHeaderRange.Font.Color = ColorToOleColor(System.Drawing.Color.White);
                xlHeaderRange.Interior.Color = ColorToOleColor(
                    System.Drawing.ColorTranslator.FromHtml("#1A3A5C"));
                xlHeaderRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlHeaderRange.RowHeight = 25;

                if (rowCount > 0)
                {
                    // Build data array from DataTable rows
                    object[,] dataArray = new object[rowCount, colCount];

                    for (int r = 0; r < rowCount; r++)
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            object val = dt.Rows[r][c];
                            dataArray[r, c] = val == DBNull.Value ? "" : val;
                        }
                    }

                    // Write data in one shot
                    xlDataRange = xlSheet.Range[
                        xlSheet.Cells[2, 1],
                        xlSheet.Cells[rowCount + 1, colCount]];
                    xlDataRange.Value2 = dataArray;

                    // Auto-fit column widths for readability
                    Excel.Range xlUsed = null;
                    try
                    {
                        xlUsed = xlSheet.UsedRange;
                        xlUsed.Columns.AutoFit();
                    }
                    finally
                    {
                        SafeReleaseCom(xlUsed);
                    }
                }

                FreezeAndFilter(xlSheet, rowCount, colCount);
            }
            catch (Exception ex)
            {
                throw new Exception(string.Format(
                    "PendingForwardsExporter.WritePdfDataTab: Error writing tab '{0}': {1}",
                    tabName, ex.Message));
            }
            finally
            {
                SafeReleaseCom(xlDataRange);
                SafeReleaseCom(xlHeaderRange);
                SafeReleaseCom(xlSheet);
            }
        }

        /// <summary>
        /// Gets an existing worksheet by index or adds a new one if it doesn't exist.
        /// </summary>
        private static Excel.Worksheet GetOrAddSheet(
            Excel.Workbook xlWorkbook, int sheetIndex)
        {
            if (sheetIndex <= xlWorkbook.Sheets.Count)
                return (Excel.Worksheet)xlWorkbook.Sheets[sheetIndex];

            Excel.Sheets xlSheets = xlWorkbook.Sheets;
            Excel.Worksheet xlSheet = (Excel.Worksheet)xlSheets.Add(
                Type.Missing,
                xlWorkbook.Sheets[xlWorkbook.Sheets.Count]);
            SafeReleaseCom(xlSheets);
            return xlSheet;
        }

        /// <summary>
        /// Sets column widths from an array of widths.
        /// </summary>
        private static void SetColumnWidths(Excel.Worksheet xlSheet, int[] colWidths)
        {
            for (int c = 0; c < colWidths.Length; c++)
            {
                Excel.Range xlCol = null;
                try
                {
                    xlCol = (Excel.Range)xlSheet.Columns[c + 1];
                    xlCol.ColumnWidth = colWidths[c];
                }
                finally
                {
                    SafeReleaseCom(xlCol);
                }
            }
        }

        /// <summary>
        /// Freezes the header row and applies AutoFilter to the data range.
        /// </summary>
        private static void FreezeAndFilter(
            Excel.Worksheet xlSheet, int rowCount, int colCount)
        {
            xlSheet.Activate();
            xlSheet.Application.ActiveWindow.SplitRow = 1;
            xlSheet.Application.ActiveWindow.FreezePanes = true;

            Excel.Range xlFullRange = null;
            try
            {
                xlFullRange = xlSheet.Range[
                    xlSheet.Cells[1, 1],
                    xlSheet.Cells[rowCount + 1, colCount]];
                xlFullRange.AutoFilter(1);
            }
            finally
            {
                SafeReleaseCom(xlFullRange);
            }
        }

        /// <summary>
        /// Applies a number format to a data column from row 2 to rowCount + 1.
        /// </summary>
        private static void FormatColumn(
            Excel.Worksheet xlSheet,
            int rowCount,
            int colIndex,
            string format)
        {
            Excel.Range xlRange = null;
            try
            {
                xlRange = xlSheet.Range[
                    xlSheet.Cells[2, colIndex],
                    xlSheet.Cells[rowCount + 1, colIndex]];
                xlRange.NumberFormat = format;
            }
            finally
            {
                SafeReleaseCom(xlRange);
            }
        }

        /// <summary>
        /// Converts a System.Drawing.Color to an OLE color integer for Excel Interop.
        /// </summary>
        private static int ColorToOleColor(System.Drawing.Color color)
        {
            return System.Drawing.ColorTranslator.ToOle(color);
        }

        /// <summary>
        /// Safely releases a COM object to prevent memory leaks.
        /// </summary>
        private static void SafeReleaseCom(object comObject)
        {
            if (comObject == null) return;
            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
                // Suppress errors during COM release
            }
            finally
            {
                comObject = null;
            }
        }
    }
}