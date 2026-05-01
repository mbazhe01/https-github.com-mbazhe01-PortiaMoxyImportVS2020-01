using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Entities
{
    public class PendingForwardsResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public string TB10PdfPath { get; set; } // fist source file path
        public string TB20PdfPath { get; set; } // second source file path
        public DataTable DTDataTB10 { get; set; } // data table parsed from first source file
        public DataTable DTDataTB20 { get; set; } // data table parsed from second source file

        public static PendingForwardsResult Failure(string error) =>
        new PendingForwardsResult { Success = false, ErrorMessage = error };

        public static PendingForwardsResult Ok(string tb10Path, string tb20Path, DataTable tb10Data, DataTable tb20Data) =>
            new PendingForwardsResult { Success = true, TB10PdfPath = tb10Path, TB20PdfPath = tb20Path, DTDataTB10 = tb10Data, DTDataTB20 = tb20Data };
    }
}
