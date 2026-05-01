using PortiaMoxyImport.PendingForwardsClasses;
using System.Collections.Generic;
using System.Data;

public class ComparisonResult
{
    public bool Success { get; set; }
    public string ErrorMessage { get; set; }
    public List<ComparisonRow> TB10Rows { get; set; }
    public List<ComparisonRow> TB20Rows { get; set; }
    public DataTable PdfTB10Data { get; set; }
    public DataTable PdfTB20Data { get; set; }

    public static ComparisonResult Failure(string error) =>
        new ComparisonResult { Success = false, ErrorMessage = error };

    public static ComparisonResult Ok(
        List<ComparisonRow> tb10Rows,
        List<ComparisonRow> tb20Rows,
        DataTable pdfTB10Data,
        DataTable pdfTB20Data) =>
        new ComparisonResult
        {
            Success = true,
            TB10Rows = tb10Rows,
            TB20Rows = tb20Rows,
            PdfTB10Data = pdfTB10Data,
            PdfTB20Data = pdfTB20Data
        };
}