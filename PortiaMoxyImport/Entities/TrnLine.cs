using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PortiaMoxyImport.Entities
{
    class TrnLine
    {
        String portCode;
        String tranCode;
        String comment;
        String secType;
        String symbol;
        DateTime tradeDate;
        DateTime settleDate;
        DateTime origCostDate;
        Double quantity;
        String closeMeth;
        String versusDate;
        String sourceType;
        String sourceSymbol;
        Double tdFXRate;
        Double sdFXRate;
        Double origFXRate;
        String MarkToMarket;
        Double tradeAmt;
        Double origCost;
        String reserved;
        String withholdingTax;
        String exchange;
        Double exchangeFee;
        Double commission;
        String broker;
        String impliedComm;
        String otherFees;
        String commissionPurpose;
        String pledge;
        String lotLocation;
        String destPledge;
        String destLotLocation;
        String origFace;
        String yieldOnCost;
        String durationOnCost;
        String userDef1;
        String userDef2;
        String userDef3;
        String tranId;
        String ipCounter;
        String repl;
        String source;
        String reserved2;
        String omnibus;
        String recon;
        String post;
        String labname;
        String nlandef;
        String dlabdef;
        String slabdef;
        String reserved3;
        String recordDate;
        String reclaimAmt;
        String strategy;
        String reserved4;
        String reserved5;
        String reserved6;
        String reserved7;
        String perfCW;
       

        public TrnLine(string portCode, string tranCode, string secType, string symbol)
        {
            this.PortCode = portCode;
            this.TranCode = tranCode;
            this.SecType = secType;
            this.Symbol = symbol;
        }

        public TrnLine(String portCode, String tranCode,
                String secType, String symbol, DateTime tradeDate, DateTime settleDate,
                Double qty, String closingMeth, Double tdFXRate, Double sdFXRate,
                Double tradeAmt)
        {
            this.PortCode = portCode;
            this.TranCode = tranCode;
            this.SecType = secType;
            this.Symbol = symbol;
            this.TradeDate = tradeDate;
            this.SettleDate = settleDate;
            this.Quantity = qty;
            this.CloseMeth = closingMeth;
            this.TdFXRate = tdFXRate;
            this.SdFXRate = sdFXRate;
            this.TradeAmt = tradeAmt;
        }

        public string PortCode { get => portCode; set => portCode = value; }
        public string TranCode { get => tranCode; set => tranCode = value; }
        public string Comment { get => comment; set => comment = value; }
        public string SecType { get => secType; set => secType = value; }
        public string Symbol { get => symbol; set => symbol = value; }
        public DateTime TradeDate { get => tradeDate; set => tradeDate = value; }
        public DateTime SettleDate { get => settleDate; set => settleDate = value; }
        public DateTime OrigCostDate { get => origCostDate; set => origCostDate = value; }
        public double Quantity { get => quantity; set => quantity = value; }
        public string CloseMeth { get => closeMeth; set => closeMeth = value; }
        public string VersusDate { get => versusDate; set => versusDate = value; }
        public string SourceType { get => sourceType; set => sourceType = value; }
        public string SourceSymbol { get => sourceSymbol; set => sourceSymbol = value; }
        public double TdFXRate { get => tdFXRate; set => tdFXRate = value; }
        public double SdFXRate { get => sdFXRate; set => sdFXRate = value; }
        public double OrigFXRate { get => origFXRate; set => origFXRate = value; }
        public string MarkToMarket1 { get => MarkToMarket; set => MarkToMarket = value; }
        public double TradeAmt { get => tradeAmt; set => tradeAmt = value; }
        public double OrigCost { get => origCost; set => origCost = value; }
        public string Reserved { get => reserved; set => reserved = value; }
        public string WithholdingTax { get => withholdingTax; set => withholdingTax = value; }
        public string Exchange { get => exchange; set => exchange = value; }
        public double ExchangeFee { get => exchangeFee; set => exchangeFee = value; }
        public double Commission { get => commission; set => commission = value; }
        public string Broker { get => broker; set => broker = value; }
        public string ImpliedComm { get => impliedComm; set => impliedComm = value; }
        public string OtherFees { get => otherFees; set => otherFees = value; }
        public string CommissionPurpose { get => commissionPurpose; set => commissionPurpose = value; }
        public string Pledge { get => pledge; set => pledge = value; }
        public string LotLocation { get => lotLocation; set => lotLocation = value; }
        public string DestPledge { get => destPledge; set => destPledge = value; }
        public string DestLotLocation { get => destLotLocation; set => destLotLocation = value; }
        public string OrigFace { get => origFace; set => origFace = value; }
        public string YieldOnCost { get => yieldOnCost; set => yieldOnCost = value; }
        public string DurationOnCost { get => durationOnCost; set => durationOnCost = value; }
        public string UserDef1 { get => userDef1; set => userDef1 = value; }
        public string UserDef2 { get => userDef2; set => userDef2 = value; }
        public string UserDef3 { get => userDef3; set => userDef3 = value; }
        public string TranId { get => tranId; set => tranId = value; }
        public string IpCounter { get => ipCounter; set => ipCounter = value; }
        public string Repl { get => repl; set => repl = value; }
        public string Source { get => source; set => source = value; }
        public string Reserved2 { get => reserved2; set => reserved2 = value; }
        public string Omnibus { get => omnibus; set => omnibus = value; }
        public string Recon { get => recon; set => recon = value; }
        public string Post { get => post; set => post = value; }
        public string Labname { get => labname; set => labname = value; }
        public string Nlandef { get => nlandef; set => nlandef = value; }
        public string Dlabdef { get => dlabdef; set => dlabdef = value; }
        public string Slabdef { get => slabdef; set => slabdef = value; }
        public string Reserved3 { get => reserved3; set => reserved3 = value; }
        public string RecordDate { get => recordDate; set => recordDate = value; }
        public string ReclaimAmt { get => reclaimAmt; set => reclaimAmt = value; }
        public string Strategy { get => strategy; set => strategy = value; }
        public string Reserved4 { get => reserved4; set => reserved4 = value; }
        public string Reserved5 { get => reserved5; set => reserved5 = value; }
        public string Reserved6 { get => reserved6; set => reserved6 = value; }
        public string Reserved7 { get => reserved7; set => reserved7 = value; }
        public string PerfCW { get => perfCW; set => perfCW = value; }
    }
}
