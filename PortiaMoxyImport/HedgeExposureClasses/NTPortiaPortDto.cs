using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.HedgeExposureClasses
{
    public class NTPortiaPortDto
    {
        public string NTAccttId { get; set; }
        public string PortiaPort { get; set; }
        public NTPortiaPortDto(string ntAccttId, string portiaPort)
        {
            NTAccttId = ntAccttId;
            PortiaPort = portiaPort;
        }
    }

  
}//eon
