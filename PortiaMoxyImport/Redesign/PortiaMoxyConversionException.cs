using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Redesign
{
    public sealed class PortiaMoxyConversionException : Exception
    {
        public PortiaMoxyConversionException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
