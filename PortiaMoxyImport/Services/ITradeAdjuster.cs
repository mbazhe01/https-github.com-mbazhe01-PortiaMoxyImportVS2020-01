using PortiaMoxyImport.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Services
{
    public interface ITradeAdjuster
    {
        NTFXTradeDTO Adjust(NTFXTradeDTO trade);
    }
}
