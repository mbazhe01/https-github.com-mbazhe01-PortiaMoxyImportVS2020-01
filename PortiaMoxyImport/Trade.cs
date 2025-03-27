using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;

namespace PortiaMoxyImport
{
     class Trade
    {
        public string[] items= null ;                                       // Moxy trade split into an array
        string portfolio = string.Empty;                                 // Portfolio number
        protected string portType = string.Empty;                 // NonUs or Us 
        protected string secType = string.Empty;                  // SameCur or DiffCur --> same or different security currency from portfolio currency
        protected string tranType = string.Empty;                 // cash or equity
        protected string tranCode = string.Empty;                 // Buy or Sell
        string dbConnMoxy = string.Empty;                          // db connection to Moxy
        string dbConnPortia = string.Empty;                         // db connection to Portia
        string tradingCurrencyStoredProc = string.Empty;
        string reportingCurrencyStoredProc = string.Empty;
        string sellRuleStoredProc = string.Empty;
        string lastCrossRateStoredProc = string.Empty;
        protected string securityCur = string.Empty;
        protected String tradeCur = string.Empty;
        protected String repCur = string.Empty;
        public bool doNotInclude = false;
        private static int tradeCnt = 0;                                // processed trades counter
        protected MoxyDatabase md;
        PortiaDatabase pd; 
        public TextBox screen;
        public TextBox Screen
        {
            get
            {
                return screen;
            }

            set
            {
                screen = value;
            }
        }

       
        /// <summary>
        ///                     Constructor
        /// </summary>
        public Trade(TextBox aScreen, string aLine, string aDbConnMoxy, string aDbConnPortia, string aTradingCurrencySP, string aLastCrossRateSP, string aReportingCurrencySP, string aSellRuleSP)
        {
            int rtn = 0;
            tradeCnt++;
            Screen  = aScreen;
            dbConnMoxy = aDbConnMoxy;
            dbConnPortia = aDbConnPortia;
            tradingCurrencyStoredProc = aTradingCurrencySP;
            reportingCurrencyStoredProc = aReportingCurrencySP;
            sellRuleStoredProc = aSellRuleSP;
            lastCrossRateStoredProc = aLastCrossRateSP;
            items = aLine.Split(',');
            Screen.AppendText("\r\n ---\r\n Tran code: " + items[1] + " Line in the file: " + tradeCnt );

          if (tradeCnt==38 )
            {
                rtn = 0;
            }

          if(items[0].Equals("24132"))
            {
                rtn = 0;
            }

            //
            // when lot location is null or nothing --> set default lot location: 254
            //
            if (String.IsNullOrEmpty(items[29]))
            {
                items[29] = "254";
            }

            md = new MoxyDatabase(dbConnMoxy, screen);
            pd = new PortiaDatabase(dbConnPortia, screen, tradingCurrencyStoredProc, lastCrossRateStoredProc, sellRuleStoredProc);
            rtn = md.getISOCurrency(items[3], ref securityCur);
            rtn = pd.getTradingCurrency(tradingCurrencyStoredProc, items[0].Replace("fc", string.Empty), ref tradeCur);
            rtn = pd.getReportingCurrency(reportingCurrencyStoredProc, items[0].Replace("fc", string.Empty), ref repCur);

            if (String.IsNullOrEmpty(tradeCur)) {tradeCur="USD";}

            clearSettleFX(ref items);

            // based on these values the appropriate conversion function selected
            if (items[4].Equals("$cash")) { tranType = "cash"; } else { tranType = "equity"; }
            if (tradeCur.Equals("USD")) { portType = "Us"; } else { portType = "NonUs"; }
            if (securityCur.Equals(tradeCur)) { secType = "SameCur"; } else { secType = "DiffCur"; }
            if ((items[1].Equals("sl") || items[1].Equals("SL"))){tranCode = "Sell";} else {tranCode = "Buy";}

            rtn = 0;

        } // end of constructor

        /// <summary>
        ///                     Converts trade to Portia specs.
        /// </summary>
        /// <returns></returns>
        public int convert()
        {

            string funcName; 
            try
            {
                               
                funcName = tranType  +portType + secType + tranCode;
                Type thisType = this.GetType();

                // testing new function
                if (funcName == "cashNonUsDiffCurSell")
                    funcName  = "cashNonUsDiffCurSell02";

                MethodInfo theMethod = thisType.GetMethod(funcName);
                                                                                                                                                                                 
                if (theMethod == null)                 
                {
                    screen.AppendText(Globals.saveErr(String.Format("\r\n-->{0}: Function {1} not found ? ? ?", GetCurrentMethod(), funcName)));
                }
                else
                {
                    theMethod.Invoke(this, null);
                }

                
            }
            catch (Exception  ex)
            {
                
                screen.AppendText(String.Format("{0}: {1}", GetCurrentMethod (), ex.Message  ));
                Globals.WriteErrorLog(ex.ToString());
            }

            return 0;
        }// end of convert()


         /// <summary>
         /// 1.
         /// </summary>
         /// <returns></returns>
        public int cashNonUsSameCurBuy()
        {

            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
            doNotInclude = true;
            return 0;
        }

         /// <summary>
         /// 2.
         /// </summary>
         /// <returns></returns>
        public int cashNonUsSameCurSell()
        {

            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
            doNotInclude = true;
            return 0;
        }

         /// <summary>
         /// 3.
         /// </summary>
         /// <returns></returns>

        public int cashNonUsDiffCurBuy()
        {
            int rtn = 0;
            Double crossRateNum, qtyNum = 0, tradeAmt = 0;
            String crossRate = items[36];
            String qty = items[8];
            Double origTradeAmt=0;
            String secType = items[3];
            string tradeDate = items[5];
            string conversionInstruction = null ;

            qtyNum =  Double.Parse( items[8]) ; 
            tradeAmt = Double.Parse(items[17]);

            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
                                 
            // replace Moxy src & dest symbols with Portia format like -CAD CASH-
            rtn = md.convertSymbolToPortiaCash(ref items[3], ref items[4]);
            rtn = md.convertSymbolToPortiaCash(ref items[11], ref items[12]);

            if (castToTradingCurrency(tradeCur, ref items) == -1) { return -1; }
                       
            //if (!Double.TryParse(crossRate, out crossRateNum))
            // cross rate is not available
            //{
                rtn = md.getCrossRateCash(int.Parse(items[39]), tradeDate , tradeCur, secType,  items[0], ref crossRate, ref conversionInstruction);
            //}

            if (Double.TryParse(crossRate.Trim(), out crossRateNum))
            {
                if (Double.TryParse(qty, out qtyNum))
                {

                    crossRateNum = Math.Round(crossRateNum, Globals.RNDNUM);
                    if (conversionInstruction.Equals("d"))
                    {
                        tradeAmt = Math.Round(qtyNum / crossRateNum, Globals.RNDNUM);
                    }
                    else
                    {
                        tradeAmt = Math.Round(qtyNum * crossRateNum, Globals.RNDNUM);
                    }

                    origTradeAmt = Double.Parse(items[17]);
                    items[17] = tradeAmt.ToString();
                }
                else
                {
                    Screen.AppendText(Globals.saveErr(String.Format("Function {0} : Qty is unavailable for the trade-->{1}  ", GetCurrentMethod(), String.Join(",", items))));
                }
            }
            else
            {
                Screen.AppendText(Globals.saveErr(Environment.NewLine + String.Format("Function {0} : Cross rate is unavailable for the trade-->{1}  ", GetCurrentMethod(), String.Join(",", items))));
                Screen.AppendText(Globals.saveErr(Environment.NewLine + String.Format("Function {0} : Can not convert {1} to number. ", GetCurrentMethod(), crossRate)));

            }

            //Sec2Port
            if (conversionInstruction.Equals("d"))
            {
                items[13] = crossRate;
            }
            else
            {
                if (Double.TryParse(crossRate, out crossRateNum))
                {
                    Double tmp = 1 / crossRateNum;
                    items[13] = (1 / crossRateNum).ToString();
                }

            }

            //items[13] =Math.Round((qtyNum / tradeAmt), Globals.RNDNUM).ToString();

            // re-evaluate sec to port
            if (tradeCur.Equals(securityCur))
            {
                // sec2port
                items[13] = "1";
            }


            // sec2base 
            if (securityCur.Equals("USD")) { items[14] = "1"; } else { items[14] = string.Empty; }

            // sec2cbal
            items[15] = Math.Round((qtyNum/tradeAmt), Globals.RNDNUM  ).ToString() ;
            
            return rtn;
        }
         
        public int cashNonUsDiffCurSell02()
        {
            int rtn = 0;
            Double crossRateNum = 0;
            String crossRate = string.Empty;
            Double origQty = 0 ;
            Double origAmt = 0;
            string tradeDate = items[5];
            string conversionInstruction = string.Empty;

            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
            try
            {
            
                // preserve the original amounts
                origQty = Double.Parse(items[8]);
                origAmt = Double.Parse(items[17]);

                // swap to make sell a buy - Portia AIM can take only buys
                string tmp = items[3];
                string tmp2 = items[4];
                items[3] = items[11];
                items[4] = items[12];
                items[11] = tmp;
                items[12] = tmp2;
                tmp = items[8];
                items[8] = items[17];
                items[17] = tmp;
                                
                // replace Moxy src & dest symbols with Portia format like -CAD CASH-
                rtn = md.convertSymbolToPortiaCash(ref items[3], ref items[4]);
                rtn = md.convertSymbolToPortiaCash(ref items[11], ref items[12]);

                if (castToTradingCurrency(tradeCur, ref items) == -1) { return -1; }
                //items[9] = getSellingRule(items[9], items[0]);

                // get cross rate & conversion insrtuction
                String secType = items[11];
                if (!Double.TryParse(crossRate, out crossRateNum))
                // cross rate is not available
                {            
                    rtn = md.getCrossRateCash(int.Parse(items[39]), tradeDate, tradeCur, secType, items[0], ref crossRate, ref conversionInstruction);
                }

                // replace sell with buy with trading currency
                if (items[1].ToString().Equals(items[1].ToString().ToUpper()))
                    items[1] = "BY";
                else
                    items[1] = "by";

                items[4] = String.Format("-{0} CASH-", tradeCur);

                // convert qty
                //if (Double.TryParse(crossRate, out crossRateNum))
                //{

                //        if (conversionInstruction.Equals("d"))
                //        {
                //            items[8] = Math.Round((origQty  / crossRateNum), Globals.RNDNUM).ToString();
                //        }
                //        else
                //        {
                //            items[8] = Math.Round((origQty * crossRateNum), Globals.RNDNUM).ToString();
                //        }

                //}
                //else
                //{
                //    Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : Cross Rate is unavailable for the trade-->  " + String.Join(",", items)));
                //    Screen.AppendText(Globals.saveErr(Environment.NewLine + String.Format("Function {0} : Can not convert {1} to number. ", GetCurrentMethod(), crossRate)));
                //}

                // sec2port
                if (Double.TryParse(crossRate, out crossRateNum))
                {
                    if (conversionInstruction.Equals("d"))
                        items[13] = Math.Round((1 / crossRateNum), Globals.RNDNUM).ToString();
                    else
                        items[13] = Math.Round((crossRateNum), Globals.RNDNUM).ToString();
                }
                //
                // reevaluate new security currency after flip
                //
                rtn = md.getISOCurrency(items[3], ref securityCur);
                if (tradeCur.Equals(securityCur))
                {
                    // sec2port
                    items[13] = "1";
                } 

                // sec2Base

                //items[14] =Math.Round( (Double.Parse(items[8]) / origAmt), Globals.RNDNUM).ToString();
                if (securityCur.Equals("USD")) { items[14] = "1"; } else { items[14] = string.Empty; }
                // sec2cbal
                //if (securityCur.Equals("EUR") || securityCur.Equals("GBP") || securityCur.Equals("AUD") || securityCur.Equals("NZD"))
                if(conversionInstruction.Equals("d") )
                //items[15] = Math.Round(Double.Parse(items[17]) / Double.Parse(items[8]), Globals.RNDNUM).ToString();
                    items[15] = (1/Double.Parse(crossRate)).ToString();
                else
                    //items[15] = Math.Round(Double.Parse(items[8]) / Double.Parse(items[17]), Globals.RNDNUM).ToString();
                    items[15] = crossRate;
                

            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

            return rtn;

        }



         /// <summary>
         /// 4.
         /// </summary>
         /// <returns></returns>
        public int cashNonUsDiffCurSell()
        {
            int rtn = 0;
           
            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
            try
            {
                Double crossRateNum, qtyNum = 0;
                String crossRate = items[36];
                String qty = items[8];
                Double origTradeAmt ;
                string tradeDate = items[5];
                string conversionInstruction = string.Empty;
                //string securityCur = string.Empty;
                               
                //origTradeAmt = Double.Parse(items[17]);  
                // swap src dest symbols
                string tmp = items[3];
                string tmp2 = items[4];
                items[3] = items[11];
                items[4] = items[12];
                items[11] = tmp;
                items[12] = tmp2;

                tmp = items[8];
                items[8] = items[17];
                items[17] = tmp;

                //origTradeAmt = Double.Parse(items[17]);
                origTradeAmt = Double.Parse(items[8]);

                // replace Moxy src & dest symbols with Portia format like -CAD CASH-
                rtn = md.convertSymbolToPortiaCash(ref items[3], ref items[4]);
                rtn = md.convertSymbolToPortiaCash(ref items[11], ref items[12]);

                if (castToTradingCurrency(tradeCur, ref items) == -1) { return -1; }
                //items[9] = getSellingRule(items[9], items[0]);
                String secType = items[11];

                if (!Double.TryParse(crossRate, out crossRateNum))
                // cross rate is not available
                {
                    //rtn = md.getCrossRate(int.Parse(items[39]), tradeDate, tradeCur, secType, items[0], ref crossRate, ref conversionInstruction);
                    //rtn = md.getCrossRate02(int.Parse(items[39]), tradeDate, tradeCur, secType, items[0], ref crossRate, ref conversionInstruction);
                    rtn = md.getCrossRateCash(int.Parse(items[39]), tradeDate, tradeCur, secType, items[0], ref crossRate, ref conversionInstruction);
                }
                
                // replace sell with buy with trading currency
                if (items[1].ToString().Equals(items[1].ToString().ToUpper()))
                    items[1] = "BY";
                else
                    items[1] = "by";
                           
                items[4] = String.Format("-{0} CASH-", tradeCur);
                              
       
                if (Double.TryParse(crossRate, out crossRateNum))
                {
                    if (Double.TryParse(qty, out qtyNum))
                    {
                        if (conversionInstruction.Equals("d"))
                        {
                            items[8] = Math.Round((qtyNum / crossRateNum), Globals.RNDNUM).ToString();
                        }
                        else
                        {
                            items[8] = Math.Round((qtyNum * crossRateNum), Globals.RNDNUM).ToString();
                        }

                        items[17] = qty.ToString();          //  trade amount becomes orig qty 
                    }
                    else
                    {
                        Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : Qty is unavailable for the trade-->  " + String.Join(",", items)));
                    }
                }
                else
                {
                    Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : Cross Rate is unavailable for the trade-->  " + String.Join(",", items)));
                    Screen.AppendText(Globals.saveErr(Environment.NewLine + String.Format("Function {0} : Can not convert {1} to number. ", GetCurrentMethod(), crossRate)));
                }

                // sec2Base
                items[14] = (Math.Round (Double.Parse(items[8]) / origTradeAmt, Globals.RNDNUM )).ToString();
                // sec2port
                if (conversionInstruction.Equals("d"))
                    items[13] = Math.Round((1 / crossRateNum), Globals.RNDNUM ).ToString();
                else
                    items[13] = Math.Round(( crossRateNum), Globals.RNDNUM).ToString();

               // sec2cbal
                items[15] = Math.Round(Double.Parse(items[8]) / Double.Parse(qty) , Globals.RNDNUM).ToString();
                //
                // reevaluate new security currency after flip
                //
                rtn = md.getISOCurrency(items[3], ref securityCur);
                if (tradeCur.Equals(securityCur))
                {
                    // sec2port
                    items[13] = "1";                                  
                }

            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

            return 0;
        }

         /// <summary>
         /// 5.
         /// </summary>
         /// <returns></returns>
        public int cashUsSameCurBuy()
        {

            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
            doNotInclude = true;
            return 0;
        }
         /// <summary>
         /// 6.
         /// </summary>
         /// <returns></returns>

        public int cashUsSameCurSell()
        {

            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
            doNotInclude = true;
            return 0;
        }

         /// <summary>
         /// 7.
         /// </summary>
         /// <returns></returns>
        public int cashUsDiffCurBuy()
        {
            int rtn = 0;
            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
            try
            {            

                // replace Moxy src & dest symbols with Portia format like -CAD CASH-
                rtn = md.convertSymbolToPortiaCash(ref items[3], ref items[4]);
                rtn = md.convertSymbolToPortiaCash(ref items[11], ref items[12]);

                // sec2Base
                items[14] = items[13];
                
                // sec2Port
                if (tradeCur.Equals(repCur))
                    items[13] = items[13];
                else 
                    items[13] = string.Empty;
                
                // sec2Cbal
                Double qtyNum = 0, tradeAmt = 0;
                if (Double.TryParse(items[8], out qtyNum) && Double.TryParse(items[17], out tradeAmt) && !String.IsNullOrEmpty(items[17]))
                {
                    items[15] = Math.Round((qtyNum / tradeAmt), Globals.RNDNUM).ToString();
                }
                else
                {
                    Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : Qty or Trade Amount is unavailable for the trade-->  " + String.Join(",", items)));
                }
                

                //items[15] = items[13];

            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

 
            return 0;
        }

         /// <summary>
         /// 8.
         /// </summary>
         /// <returns></returns>
        public int cashUsDiffCurSell()
        {
            int rtn = 0;
            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
            try
            {

                if ( items[0].Equals( "55093")) {
                    rtn = 0;
                }

                // swap src dest symbols
                string tmp = items[3];
                items[3] = items[11];
                items[11] = tmp;

                tmp = items[4];
                items[4] = items[12];
                items[12] = tmp;


                // replace Moxy src & dest symbols with Portia format like -CAD CASH-
                rtn = md.convertSymbolToPortiaCash(ref items[3], ref items[4]);
                rtn = md.convertSymbolToPortiaCash(ref items[11], ref items[12]);

                // replace sell with buy
                if (items[1].ToString().Equals(items[1].ToString().ToUpper()))
                    items[1] = "BY";
                else
                    items[1] = "by";

                //items[9] = getSellingRule(items[9], items[0]);


                // qty & tradeAmt
                Double qtyNum = 0;
                String qty = items[8];

                if (Double.TryParse(qty, out qtyNum))
                {
                    items[8] = items[17];
                    items[17] = qtyNum.ToString();
                }
                else
                {
                    Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : Qty is unavailable for the trade-->  " + String.Join(",", items)));
                }

                // sec2Base
                String fxRate = items[13];
                Double fxRateNum = 0;

                if (Double.TryParse(fxRate, out fxRateNum) && fxRateNum != 0)
                {
                    items[14] = Math.Round((1 / fxRateNum), Globals.RNDNUM).ToString();
                }
                else
                {
                    Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : FX Rate is unavailable for the trade-->  " + String.Join(",", items)));
                }


                // sec2CBal
                Double tradeAmt = 0;
                if (Double.TryParse(items[17], out tradeAmt) && !String.IsNullOrEmpty(items[17]) && Double.TryParse(items[8], out qtyNum )  )
                {
                    items[15] = Math.Round((qtyNum / tradeAmt), Globals.RNDNUM).ToString();
                }
                else
                {
                    Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : Trade Amount is unavailable for the trade-->  " + String.Join(",", items)));
                }


                // sec2Port
                items[13] = items[14];

                //
                // reevaluate new security currency after flip
                //
                rtn = md.getISOCurrency(items[3], ref securityCur);
                if (tradeCur.Equals(securityCur))
                {
                    // sec2port
                    items[13] = "1"; 
 
                    // sec2base 
                    items[14] = "1";

                }

            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }
            return 0;
        }

         /// <summary>
         /// 9.
         /// </summary>
         /// <returns></returns>
        public int equityNonUsSameCurBuy()
        {
            int rtn=0;
            Double crossRateNum;
            String crossRate = items [36];
            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
            string tradeDate = items[5];
            string conversionInstruction = null;
            try
            {
                if (!Double.TryParse(crossRate, out crossRateNum))
                {
                    // cross rate is not available
                    rtn = md.getCrossRate02(int.Parse(items[39]), tradeDate , tradeCur, items[3], items[0], ref crossRate, ref conversionInstruction);
                }

                //
                // when security currency is USD settle it in USD
                // when security currency is not USD settle it in trading currency
                //
                //string part1 = items[11].ToString().Substring(0, 2);
               
                //if (securityCur.Equals("USD"))
                //{
                //    items[11] = part1 + securityCur.Substring(0, 2).ToLower();
                //}
                //else
                //{
                //    items[11] = part1 + tradeCur.Substring(0, 2).ToLower();
                //}

                // sec2Base 
                items[14] = items[13];

                // sec2port
                items[13] = crossRate;

                //sec2cbal
                items[15] = "1";

                // usupervised tran
                string unsupMsg = unsupervisedCheck(items[3].ToString());
                if (!String.IsNullOrEmpty(unsupMsg))
                    items[12] = unsupMsg;

                //if (items[3].ToString().Equals("usus"))
                //{
                //    items[12] = "-UNSUP USD-";
                //}


            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }
                                      
            return 0;
        }

         /// <summary>
         /// 10.
         /// </summary>
         /// <returns></returns>
        public int equityNonUsSameCurSell()
        {

            int rtn = 0;
            Double crossRateNum;
            String crossRate = items[36];
            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
            string tradeDate = items[5];
            string conversionInstruction = null;

            try
            {
                if (!Double.TryParse(crossRate, out crossRateNum))
                {
                    // cross rate is not available
                    rtn = md.getCrossRate(int.Parse(items[39]), tradeDate , tradeCur, items[3], items[0], ref crossRate, ref conversionInstruction);
                }

                //
                // when security currency is USD settle it in USD
                // when security currency is not USD settle it in trading currency
                //
                //string part1 = items[11].ToString().Substring(0, 2);
               
                //if (securityCur.Equals("USD"))
                //{
                //    items[11] = part1 + securityCur.Substring(0, 2).ToLower();
                //}
                //else
                //{
                //    items[11] = part1 + tradeCur.Substring(0, 2).ToLower();
                //}


                items[9] = getSellingRule(items[9], items[0]);

                // sec2Base 
                items[14] = items[13];

                // sec2port
                items[13] = crossRate;

                //sec2cbal
                items[15] = "1";

                // usupervised tran
                string unsupMsg = unsupervisedCheck(items[3].ToString());
                if (!String.IsNullOrEmpty(unsupMsg))
                    items[12] = unsupMsg;

                //if (items[3].ToString().Equals("usus"))
                //{
                //    items[12] = "-UNSUP USD-";
                //}

            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

            return 0;
        }

         /// <summary>
         /// 11.
         /// </summary>
         /// <returns></returns>
        public int equityNonUsDiffCurBuy()
        {

            int rtn = 0;
            Double crossRateNum;
            String crossRate = items[36];
            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
            string tradeDate = items[5];
            string conversionInstruction = null;

            try            
            {
                string unsupMsg = unsupervisedCheck(items[3].ToString());
                if (!String.IsNullOrEmpty(unsupMsg))
                    items[12] = unsupMsg;
                // usupervised tran
                //if (items[3].ToString().Equals("usus"))
                //{
                //    items[12] = "-UNSUP USD-";
                //}

                   // if (!Double.TryParse(crossRate, out crossRateNum))
                //{
                    // cross rate is not available
                    rtn = md.getCrossRate02(int.Parse(items[39]), tradeDate , tradeCur, items[3], items[0], ref crossRate, ref conversionInstruction);
                //}

                //
                // when security currency is USD settle it in USD
                // when security currency is not USD settle it in trading currency
                //
                //string part1 = items[11].ToString().Substring(0, 2);
                                    
                //if (securityCur.Equals("USD"))
                //{
                //    items[11] = part1 + securityCur.Substring(0, 2).ToLower();
                //}
                //else
                //{
                //    items[11] = part1 + tradeCur.Substring(0, 2).ToLower();
                //}


                // sec2Base 
                //if (items[3].ToUpper().Substring(2, 2).Equals("US"))
                //    items[14] = "1";
                //else
                    items[14] = items[13];

                // sec2port
                if (conversionInstruction.Equals("d"))
                {
                    items[13] = crossRate;      // changed 9/25/17 mikeba
                    
                }
                else
                {
                    if (Double.TryParse(crossRate, out crossRateNum))
                        items[13] = (1 / crossRateNum).ToString();
                }
                   

                //sec2cbal
                //if(items[3].ToUpper().Substring(2,2).Equals(items[11].ToUpper().Substring(2,2))) {
                    items[15] = "1";
                //}
                //else
                //{
                //    items[15] = items[13];
                //}
                
            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

            return 0;
        }

        /// <summary>
        /// 12.
        /// </summary>
        /// <returns></returns>
        public int equityNonUsDiffCurSell()
        {

            int rtn = 0;
            Double crossRateNum;
            String crossRate = items[36];
            string tradeDate = items[5];
            string conversionInstruction = null;

            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));

            try
            {
               
                //if (!Double.TryParse(crossRate, out crossRateNum))
                //{
                    // cross rate is not available
                     //rtn = md.getCrossRate(int.Parse(items[39]), tradeDate, tradeCur, items[3], items[0], ref crossRate, ref conversionInstruction);
                    rtn = md.getCrossRate02(int.Parse(items[39]), tradeDate, tradeCur, items[3], items[0], ref crossRate, ref conversionInstruction);
                //}

                //
                // when security currency is USD settle it in USD
                // when security currency is not USD settle it in trading currency
                //
                //string part1 = items[11].ToString().Substring(0, 2);
               
                //if (securityCur.Equals("USD"))
                //{
                //    items[11] = part1 + securityCur.Substring(0, 2).ToLower();
                //}
                //else
                //{
                //    items[11] = part1 + tradeCur.Substring(0, 2).ToLower();
                //}


                items[9] = getSellingRule(items[9], items[0]);

                // sec2Base 
                //if (items[3].ToUpper().Substring(2, 2).Equals("US"))
                //    items[14] = "1";
                //else
                    items[14] = items[13];

                // sec2port
                if (Double.TryParse(crossRate, out crossRateNum))
                {
                    if (conversionInstruction.Equals("d"))
                    {
                        items[13] = Math.Round((crossRateNum), Globals.RNDNUM).ToString();
                       
                    }
                    else
                    {
                        items[13] = Math.Round((1 / crossRateNum), Globals.RNDNUM).ToString();
                    }

                    //items[13] = Math.Round((crossRateNum), Globals.RNDNUM).ToString();
                    //items[13] = Math.Round((1 / crossRateNum), Globals.RNDNUM).ToString();
                }
                //sec2cbal
                if (items[3].ToUpper().Substring(2, 2).Equals(items[11].ToUpper().Substring(2, 2)))
                {
                    items[15] = "1";
                }
                else
                {
                    items[15] = items[13];
                }

                // usupervised tran
                string unsupMsg = unsupervisedCheck(items[3].ToString());
                if (!String.IsNullOrEmpty(unsupMsg))
                    items[12] = unsupMsg;
                //if (items[3].ToString().Equals("usus"))
                //{
                //    items[12] = "-UNSUP USD-";
                //}


            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

            return 0;
        }

         /// <summary>
         /// 13.
         /// </summary>
         /// <returns></returns>
        public int equityUsSameCurBuy()
        {

            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
           
            try
            {
                // sec2port
                items[13] = "1";
                // sec2base
                items[14] = "1";
                // sec2cbal
                items[15] = "1";

                // usupervised tran
                string unsupMsg = unsupervisedCheck(items[3].ToString());
                if (!String.IsNullOrEmpty(unsupMsg))
                    items[12] = unsupMsg;

                //if (items[3].ToString().Equals("usus"))
                //{
                //    items[12] = "-UNSUP USD-";
                //}

            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

            return 0;
        }

         /// <summary>
         /// 14.
         /// </summary>
         /// <returns></returns>
        public int equityUsSameCurSell()
        {

            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
           
            try
            {
                items[9] = getSellingRule(items[9], items[0]);

                // sec2port
                items[13] = "1";
                // sec2base
                items[14] = "1";
                // sec2cbal
                items[15] = "1";

                // usupervised tran
                string unsupMsg = unsupervisedCheck(items[3].ToString());
                if (!String.IsNullOrEmpty(unsupMsg))
                    items[12] = unsupMsg;


                //if (items[3].ToString().Equals("usus"))
                //{
                //    items[12] = "-UNSUP USD-";
                //}

            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

            return 0;
        }

         /// <summary>
         /// 15.
         /// </summary>
         /// <returns></returns>
        public int equityUsDiffCurBuy()
        {

            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));          

            try
            {
                // usupervised tran
                string unsupMsg = unsupervisedCheck(items[3].ToString());
                if (!String.IsNullOrEmpty(unsupMsg))
                    items[12] = unsupMsg;


                //if (items[3].ToString().Equals("usus"))
                //{
                //    items[12] = "-UNSUP USD-";
                //}

                //
                // when security currency is USD settle it in USD
                // when security currency is not USD settle it in trading currency
                //
                //string part1 = items[11].ToString().Substring(0, 2);

                //if (securityCur.Equals("USD"))
                //{
                //    items[11] = part1 + securityCur.Substring(0, 2).ToLower();
                //}
                //else
                //{
                //    items[11] = part1 + tradeCur.Substring(0, 2).ToLower();
                //}

                // sec2base
                //if (items[3].ToUpper().Substring(2, 2).Equals("US"))
                //    items[14] = "1";
                //else
                items[14] = items[13];
                // sec2cbal


                // sec2port
                //items[13] = items[13];
                if (tradeCur.Equals(repCur))
                    items[13] = items[13];
                else
                    items[13] = string.Empty;

               
                //if (items[3].ToUpper().Substring(2, 2).Equals(items[11].ToUpper().Substring(2, 2)))
                //{
                    items[15] = "1";
                //}
                //else
                //{
                //    items[15] = items[13];
                //}
            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

            return 0;
        }

         /// <summary>
         ///  16.
         /// </summary>
         /// <returns></returns>
        public int equityUsDiffCurSell()
        {

            Screen.AppendText(String.Format("Executing {0} ", GetCurrentMethod()));
           
            try
            {
               
                
                items[9] = getSellingRule(items[9], items[0]);

                // sec2base
                items[14] = items[13];

                // sec2port
                if (repCur.Equals(tradeCur))
                    items[13] = items[13];
                else
                    items[13] = string.Empty;
                               
                  
                // sec2cbal
                //if (items[3].ToUpper().Substring(2, 2).Equals(items[11].ToUpper().Substring(2, 2)))
                //{
                    items[15] = "1";
                //}
                //else
                //{
                //    items[15] = items[13];
                //}

                // usupervised tran
                string unsupMsg = unsupervisedCheck(items[3].ToString());
                if (!String.IsNullOrEmpty(unsupMsg))
                    items[12] = unsupMsg;

                //if (items[3].ToString().Equals("usus"))
                //{
                //    items[12] = "-UNSUP USD-";
                //}

            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }


            return 0;
        }

        public string GetCurrentMethod()
        {
            StackTrace st = new StackTrace();
            StackFrame sf = st.GetFrame(1);

            return sf.GetMethod().Name;
        }

         /// <summary>
         ///    cashToTradingCurrency - to overcome Moxy trading against UDS
         ///                                                for NON US Based portfolios cast cash
         ///                                                trades to trading currency
         /// </summary>
         /// <param name="tradeCur">portfolio's trading currency</param>
         /// <param name="items">array representing the current trade</param>

        public  int castToTradingCurrency(string tradeCur, ref string[] items)
        {
            int rtn = 0;

            try
            {
                if (tradeCur.IndexOf("USD") != -1)
                {
                    rtn = -1;
                    Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " :  Wrong use of function. Could be used for non us portfolios only"));
                }
                else
                {
                    if ( (items[4].IndexOf(tradeCur) == -1) && (items[12].IndexOf(tradeCur) == -1))
                    {
                        if (items[4].IndexOf("USD") != -1)
                        {
                            items[4] = items[4].Replace("USD", tradeCur);
                            items[3] = items[3].Substring(0, 2) + tradeCur.ToLower().Substring(0, 2);     
                            // adjust qty to tradecur

                         }

                        if (items[12].IndexOf("USD") != -1)
                        {
                            items[12] = items[12].Replace("USD", tradeCur);
                            // adjust amt to trade cur

                        }

                    }
                }


            }
            catch (Exception ex)
            {
                rtn = -1;
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }
                return rtn;
        }// end of function

        /// <summary>
        ///     clearSettleFX() - clear settle fx field in a Moxy trade.
        ///                                 This field will be used as Sec2Base place holder.
        /// </summary>
        /// <param name="items">Moxy Trade represented by string array of items.</param>
        private void clearSettleFX(ref string[] items)
        {
            // settleFX is in position 14
            items[14] = string.Empty;
        }

         /// <summary>
         ///    getSellingRule() - convert moxy selling rule to numeric Portia rule
         /// </summary>
         /// <param name="switchCase"></param>
         /// <returns></returns>
        public string getSellingRule(string switchCase, string portfolio)
        {
            string rtn = string.Empty;

            try
            {
                // when selling rule is undefined in Moxy --> get selling rule from portia
                if (String.IsNullOrEmpty(switchCase))
                {
                    rtn = pd.getSellingRule(portfolio);
                    return rtn;
                }

                switch (switchCase)
                {
                    case "f":
                        rtn = "1";               // FIFO Portia code
                        break;
                    case "h":                   // Least Gain
                        rtn = "4";
                        break;
                    case "c":                   // most Gain
                        rtn = "3";
                        break;
                    case "l":                   // LIFO
                        rtn = "2";
                        break;
                    case "a":
                        rtn = "9";
                        break;
                    default:
                        // specific lot
                        rtn = "0";
                        break;
                }
            }
            catch (Exception ex)
            {        
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

         
            return rtn;
        }

        /// <summary>
        ///     check if the security type is unsupervised and return unsupervised
        ///     message with appropriate currency
        /// </summary>
        /// <param name="sectype"></param>
        /// <returns></returns>
        public String unsupervisedCheck(string sectype)
        {
            String rtnMsg = null;
            try
            {
                // check if sectype is unsupervise
                if (!sectype.StartsWith("u", StringComparison.OrdinalIgnoreCase))
                    return rtnMsg;

                rtnMsg = $"-UNSUP {securityCur}-";
                items[34] = "y";   // aim will sent it to Portia as unsupervised

            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

            return rtnMsg;
        }


    }//end of class
}// end of namespace
