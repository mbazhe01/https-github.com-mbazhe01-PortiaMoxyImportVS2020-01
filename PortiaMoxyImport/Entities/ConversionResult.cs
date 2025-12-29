using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Entities
{
    public sealed class ConversionResult
    {
        public bool Success { get; }
        public string ErrorMessage { get; }

        private ConversionResult(bool success, string errorMessage)
        {
            Success = success;
            ErrorMessage = errorMessage;
        }

        public static ConversionResult Ok() => new ConversionResult(true, null);

        public static ConversionResult Fail(string errorMessage) =>
            new ConversionResult(false, errorMessage);
    }

}
