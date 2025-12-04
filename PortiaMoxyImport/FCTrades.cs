 using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using System.Globalization;
using System.Windows.Forms;
using System.Diagnostics;

namespace PortiaMoxyImport
{
    /// <summary>
    ///     This class holds forward contact trades
    /// </summary>
    class FCTrades
    {
        ArrayList trades;
        RichTextBox screen;
        /// <summary>
        /// Constructor
        /// </summary>
        public FCTrades(RichTextBox aScreen)
        {
            screen = aScreen; 
            trades = new ArrayList();
         }

        public void addTrade(String trade) {
            trades.Add (trade);
        }

        public String getBrokerForTrade(String trade, ref String tradeDate)
        {
            String broker= String.Empty ;
            string[] tradeItems = trade.Split(',');       
        
            foreach (String s in trades) {
                string[] fcItems = s.Split(',');
                if (fcItems.Length > 0 && fcItems[0].Substring(0, 1).Equals("7".ToString()))
                {
                    fcItems[0] = '2' + fcItems[0].Remove(0, 1);
                }
                // for fc trade to match cash offset compare:
                // Portfolio, tranCode, tradedate, settledate, qty, tradeAmt
               // if (tradeItems[0].Equals(fcItems[0]) && tradeItems[1].Equals(fcItems[1]) && tradeItems[5].Equals(fcItems[5]) && tradeItems[6].Equals(fcItems[6]) && tradeItems[8].Equals(fcItems[8]) && tradeItems[17].Equals(fcItems[17])) 
               
                 //&& tradeItems[1].Equals(fcItems[1])  

                if (tradeItems[0].Equals(fcItems[0]) && tranCodeMatch(tradeItems[1], fcItems[1]) && tradeDateMatch(tradeItems[5], fcItems[6]) && tradeItems[8].Equals(fcItems[8]) && tradeItems[17].Equals(fcItems[17])) 
                {
                    // match found
                    tradeDate  = fcItems[5];  // trade date is not avail for cash tran. set it here
                    broker = fcItems[24];
                    break;
                }

            } // end of For Loop

            return broker;
        }
        /// <summary>
        ///  trancodeMatch() compares two transactions code and return true if they are the same or compatible
        ///  Possible tran code values: by, sl, ss
        /// </summary>
        /// <param name="aTranCode1">tran code 1</param>
        /// <param name="aTranCode2">tran code 2</param>
        /// <returns></returns>
        protected Boolean tranCodeMatch(String aTranCode1, String aTranCode2)
        {
            Boolean rtn = false;

            if (!String.IsNullOrEmpty(aTranCode1) && !String.IsNullOrEmpty(aTranCode2))
            {
                if (aTranCode1.Equals(aTranCode2))
                {
                    return true;
                }
                if (aTranCode1.Equals("ss") && aTranCode2.Equals("sl"))
                {
                    return true;
                }
                if (aTranCode1.Equals("sl") && aTranCode2.Equals("ss"))
                {
                    return true;
                }



            }


            return rtn;
        }

        protected Boolean tradeDateMatch (String aTradeDate1, String aTradeDate2) {
            Boolean rtn = false;

            if (!String.IsNullOrEmpty(aTradeDate1) && !String.IsNullOrEmpty(aTradeDate2))
            {
                if (aTradeDate1.Equals(aTradeDate2))
                {
                    return true;
                }

                // convert to date time
                string[] formats = {"MMddyyyy", "MMdd/yy"};
                
                DateTime dateValue1= DateTime.MaxValue , dateValue2= DateTime.MinValue ;

                // convert first date
                   try
                    {
                        dateValue1 = DateTime.ParseExact(aTradeDate1, formats,
                                                        new CultureInfo("en-US"),
                                                        DateTimeStyles.None);
                        
                    }
                    catch (FormatException ex)
                    {
                    
                        screen.AppendText(Globals.saveErr(String.Format ("/r/n{1}: Unable to convert '{0}' to a date./r/n", aTradeDate1, GetCurrentMethod())));
                        Globals.WriteErrorLog(ex.ToString());
                    }

                    // convert second date
                    try
                    {
                        dateValue2 = DateTime.ParseExact(aTradeDate2, formats,
                                                        new CultureInfo("en-US"),
                                                        DateTimeStyles.None);
                        
                    }
                    catch (FormatException e)
                    {
                        screen.AppendText(Globals.saveErr(String.Format("/r/n{1}: Unable to convert '{0}' to a date./r/n", aTradeDate2, GetCurrentMethod())));
                        Globals.WriteErrorLog(e.ToString());
                    }

                    // compare two dates
                    if (DateTime.Compare(dateValue1, dateValue2) == 0)
                    {
                        return true;
                    }


                }

            

            return rtn;

        }

        public string GetCurrentMethod()
        {
            StackTrace st = new StackTrace();
            StackFrame sf = st.GetFrame(1);

            return sf.GetMethod().Name;
        }


    }// end of class
}
