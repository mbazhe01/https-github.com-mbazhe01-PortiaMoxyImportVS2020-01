using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.HedgeExposureClasses
{
    public class PortiaHdgExposureDto
    {

        public DateTime AsOfDate { get; set; }

        public string Account { get; set; }

        public string Country { get; set; }

        public decimal MarketValueStocks { get; set; }

        public decimal MarketValueForwards { get; set; }

        public decimal HedgeAmount { get; set; }

        /// <summary>
        /// this is actualy the currency field. It is named security in the source file but it contains the currency code (e.g. USD, EUR, etc.)
        /// </summary>
        public string Security { get; set; }
    }
}
