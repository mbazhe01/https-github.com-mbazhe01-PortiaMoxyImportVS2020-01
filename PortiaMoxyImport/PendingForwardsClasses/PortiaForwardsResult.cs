using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.PendingForwardsClasses
{
    public class PortiaForwardsResult
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; }
        public DataTable Data { get; set; }

        public static PortiaForwardsResult Failure(string error) =>
            new PortiaForwardsResult { Success = false, ErrorMessage = error };

        public static PortiaForwardsResult Ok(DataTable data) =>
            new PortiaForwardsResult { Success = true, Data = data };
    }
}
