using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using PortiaMoxyImport.Entities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using static PortiaMoxyImport.Entities.HedgeExposureDTO;

namespace PortiaMoxyImport
{
    internal class HedgeExposureCsvReader : IHedgeExposureCsvReader
    {
        private static readonly CultureInfo CsvCulture = CultureInfo.InvariantCulture;
        private string localFilePath;

        public HedgeExposureCsvReader(string localFilePath)
        {
            this.localFilePath = localFilePath;
        }

        public List<HedgeExposureDTO> Read(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("File path must be provided.", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel file not found.", filePath);

            FileStream stream = null;
            StreamReader reader = null;
            CsvReader csv = null;

            try
            {
                stream = new FileStream(
                    filePath,
                    FileMode.Open,
                    FileAccess.Read,
                    FileShare.Read);

                reader = new StreamReader(stream);

                var config = new Configuration(CultureInfo.InvariantCulture)
                {
                    HasHeaderRecord = true
                };

                csv = new CsvReader(reader, config);
                csv.Configuration.RegisterClassMap<HedgeExposureDTOMap>();

                return csv.GetRecords<HedgeExposureDTO>().ToList();

            }
            catch (HeaderValidationException ex)
            {
                throw new InvalidDataException("CSV header mismatch. Column names may not match expected headers.", ex);
            }
            catch (TypeConverterException ex)
            {
                // Usually a bad number/date somewhere
                throw new InvalidDataException(
                       $"HedgeExposureCsvReader: {ex.Message}.",
                            ex);
            }
            catch (CsvHelperException ex)
            {
                throw new InvalidDataException(
                    $"CSV parsing failed {ex.Message}.",
                    ex);
            }
            catch (IOException ex)
            {
                throw new IOException("Error reading CSV file.", ex);
            }
            finally
            {
                // Dispose in reverse order
                if (csv != null) csv.Dispose();
                if (reader != null) reader.Dispose();
                if (stream != null) stream.Dispose();
            }
        }

        public Task<IReadOnlyList<HedgeExposureDTO>> ReadAsync(string filePath, CancellationToken cancellationToken = default)
        {
            throw new NotImplementedException();
        }
    }// eoc


}
