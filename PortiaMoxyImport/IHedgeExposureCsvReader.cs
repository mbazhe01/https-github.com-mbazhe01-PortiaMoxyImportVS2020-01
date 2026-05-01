using PortiaMoxyImport.Entities;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
namespace PortiaMoxyImport
{
    internal interface IHedgeExposureCsvReader
    {
        Task<IReadOnlyList<HedgeExposureDTO>> ReadAsync(
        string filePath,
        CancellationToken cancellationToken = default);

        List<HedgeExposureDTO> Read(string filePath);

    }
}
