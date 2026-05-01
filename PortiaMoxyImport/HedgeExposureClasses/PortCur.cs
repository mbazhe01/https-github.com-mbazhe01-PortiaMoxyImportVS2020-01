using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.HedgeExposureClasses
{
    public class PortCur
    {
        public string Port { get; set; }
        public string Currency { get; set; }


        public PortCur() { }

        public PortCur(string port, string currency)
            {
                Port = port;
                Currency = currency;
        }

        public override bool Equals(object obj)
        {
            if (obj is PortCur other)
            {
                return this.Port == other.Port && this.Currency == other.Currency;
            }
            return false;
        }

        public override int GetHashCode()
        {
            // Combines the hash codes of the two strings
            return (Port?.GetHashCode() ?? 0) ^ (Currency?.GetHashCode() ?? 0);
        }
    }
}
