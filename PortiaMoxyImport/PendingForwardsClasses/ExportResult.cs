

namespace PortiaMoxyImport.PendingForwardsClasses
{
    /// <summary>
    /// Result object for the Excel export operation.
    /// </summary>
    public class ExportResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string FilePath { get; set; }

        public static ExportResult Failure(string error) =>
            new ExportResult { Success = false, ErrorMessage = error };

        public static ExportResult Ok(string filePath) =>
            new ExportResult { Success = true, FilePath = filePath };
    }
}
