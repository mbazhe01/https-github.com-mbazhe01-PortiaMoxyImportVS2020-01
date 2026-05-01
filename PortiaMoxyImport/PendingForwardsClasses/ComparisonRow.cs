using System;

namespace PortiaMoxyImport.PendingForwardsClasses
{
    /// <summary>
    /// Holds a single comparison result row — Portia data, matched PDF data,
    /// variances, match status, and mismatch reason.
    /// </summary>
    public class ComparisonRow
    {
        // --- Portia fields ---
        public string Portfolio { get; set; }
        public string TranType { get; set; }
        public string Currency { get; set; }
        public DateTime TradeDate { get; set; }
        public DateTime SettleDate { get; set; }
        public decimal LocalAmt { get; set; }
        public decimal USDAmt { get; set; }
        public decimal ExchangeRate { get; set; }
        public string Broker { get; set; }

        // --- PDF matched fields ---
        public decimal? PdfLocalAmt { get; set; }
        public decimal? PdfUSDAmt { get; set; }
        public decimal? PdfContractRate { get; set; }
        public string PdfSettleDate { get; set; }

        // --- Variances ---
        public decimal? LocalAmtVariance { get; set; }
        public decimal? USDAmtVariance { get; set; }
        public decimal? ExRateVariance { get; set; }

        // --- Status ---
        public bool IsMatched { get; set; }
        public string MismatchReason { get; set; }
    }
}