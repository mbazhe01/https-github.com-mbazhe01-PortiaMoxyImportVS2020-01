using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Entities
{
    internal class FileConversionDTO
    {
        public string SourceFilePath { get; set; }
        public string TargetFilePath { get; set; }
        public string FileType { get; set; }
    }
}
