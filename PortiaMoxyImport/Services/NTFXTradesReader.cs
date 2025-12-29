using CsvHelper;
using CsvHelper.Configuration;
using PortiaMoxyImport.Entities;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
//using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PortiaMoxyImport.Services
{
    public class NTFXTradesReader : IGetNTFXTradesFromFile
    {
        private String _filePath;

        public NTFXTradesReader(String filePath)
        {
            _filePath = filePath;
        }

        public Task<List<NTFXTradeDTO>> GetTradesFromFileAsync()
        {
            return Task.Run(() =>
            {
                var trades = new List<NTFXTradeDTO>();

                try
                {
                    if (!File.Exists(_filePath))
                    {
                        throw new FileNotFoundException("File not found: " + _filePath, _filePath);
                    }

                    using (var reader = new StreamReader(_filePath))
                    using (var csv = new CsvReader(reader)) // ← Correct for CsvHelper 12.2.1
                    {
                        // Configuration must be set THROUGH the Configuration object
                        csv.Configuration.CultureInfo = CultureInfo.InvariantCulture;
                        csv.Configuration.HasHeaderRecord = true;
                        csv.Configuration.TrimOptions = TrimOptions.Trim;
                        csv.Configuration.IgnoreBlankLines = true;

                        csv.Configuration.RegisterClassMap<NTFXTradeCsvRowMap>();

                        foreach (var row in csv.GetRecords<NTFXTradeCsvRow>())
                        {

                            validateNTFXTradeCsvRow(row);

                            var dto = new NTFXTradeDTO(
                                row.TradeDate,
                                row.Account,
                                row.BuySell,
                                row.Currency,
                                row.Amount,
                                row.OtherCurrency,
                                row.ForwardRate,
                                row.OtherAmount,
                                row.ValueDate,
                                row.Broker
                            );

                            trades.Add(dto);
                        }
                    }

                    return trades;
                }
                catch (Exception ex)
                {
                    throw new ApplicationException(
                        "Error reading NTFX trades from file " + _filePath + ".", ex);
                }
            });
        }

        private void validateNTFXTradeCsvRow(NTFXTradeCsvRow row)
        {
            if(String.IsNullOrEmpty(row.Account) )
            {
                throw new ApplicationException(
                    "Invalid account : " + row.Account + " for trade " + row.ToString());
            }
        }

    }
}
