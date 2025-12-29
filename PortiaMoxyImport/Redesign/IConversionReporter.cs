using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Redesign
{
    internal interface IConversionReporter
    {
        void Info(string message);
        void Success(string message);
        void Warning(string message);
        void Error(string message);

        void SetStatus(string message);
        void Clear();

    }
}
