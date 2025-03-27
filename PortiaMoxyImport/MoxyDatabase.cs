using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;
using System.Diagnostics;
using PortiaMoxyImport.Entities;

namespace PortiaMoxyImport
{
    class MoxyDatabase
    {
        String dbConnection;
        TextBox screen;


        public MoxyDatabase(String aDbConnection, TextBox aScreenTextBox)
        {
            // constructor
            
            dbConnection = aDbConnection;
            if (String.IsNullOrEmpty(dbConnection))
            {
                Globals.saveErr("Moxy connection string is null or empty.");
                throw new ArgumentNullException("Moxy DB Connection"); 
            }
                    

            screen = aScreenTextBox;
        }

        public int convertSymbolToPortiaFwdCash(ref string aSecType, ref string aSymbol)
        {
            int rtn = 0;
            string ISOCode = string.Empty;
            string msg = string.Empty;
            try
            {
                rtn = getISOCashXref(aSecType, ref  ISOCode);
                if (rtn > 0)
                {
                    aSymbol = "-" + ISOCode.Trim() + " FWD CASH-";
                    if (! aSecType.Equals("cacc"))
                        aSecType = aSecType.Substring(0, 2) + ISOCode.Trim().ToLower().Substring(0, 2);
                }
                else
                {
                    screen.AppendText(String.Format("\r\n ISO Code not found for sec type: {0}", aSecType));
                }

            }
            catch (Exception e)
            {
                rtn = -1;
                msg += this.GetCurrentMethod() + ": " + e.Message + "\r\n";
                screen.AppendText(msg);
                Globals.saveErr(msg);
                Globals.WriteErrorLog(e.ToString());

            }

            return rtn;

        }

        public int convertSymbolToPortiaCash(ref string aSecType, ref string aSymbol)
        {
            int rtn = 0;
            string ISOCode = string.Empty;
            string msg = string.Empty;
            try
            {
                rtn = getISOCashXref(aSecType, ref  ISOCode);
                if (rtn > 0)
                {
                    aSymbol = "-" + ISOCode.Trim() + " CASH-";
                    aSecType = aSecType.Substring(0, 2) + ISOCode.Trim().ToLower().Substring(0, 2);    
                }
                else
                {
                    screen.AppendText( String.Format("\r\n ISO Code not found for sec type: {0}", aSecType));
                }

            }
            catch (Exception e)
            {
                rtn = -1;
                msg+=  this.GetCurrentMethod() + ": " + e.ToString() + "\r\n";
                screen.AppendText( msg);
                Globals.saveErr(msg);
                Globals.WriteErrorLog(e.ToString());
            }
                                    
            return rtn;

        }

        public List<TrnLine>  getMoxyExport(DateTime asOfDate)
        {
            List<TrnLine> trades =  new List<TrnLine>();
            try
            {
                // get stored procedure name
                String storedProc = Util.getAppConfigVal("moxyExportSP");
                // get how many days back check trade date
                Int16 backDays = Int16.Parse(Util.getAppConfigVal("tradeDaysBack"));

                // get trades ready for export from moxy
                DataTable table = new DataTable();

                var con = new SqlConnection(dbConnection);
                con.Open();
                var cmd = new SqlCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = storedProc;
                cmd.Connection = con;

                // Add parameters and set values.  
                SqlParameter selectedDate = cmd.Parameters.Add(new SqlParameter( "@asofdate", SqlDbType.DateTime));
                selectedDate.Direction = ParameterDirection.Input;
                selectedDate.Value = asOfDate;

                SqlParameter daysBack = cmd.Parameters.Add(new SqlParameter("@check_back_days", SqlDbType.Int));
                daysBack.Direction = ParameterDirection.Input;
                daysBack.Value = backDays;

                //var da = new SqlDataAdapter(cmd.CommandText, con);
                //da.Fill(table);
               
                using (SqlDataReader rdr = cmd.ExecuteReader())
                {

                    if (rdr.HasRows)
                    {
                        // iterate through results, printing each to console
                        while (rdr.Read())
                        {
                            trades.Add(new TrnLine(
                                                        rdr[1].ToString().Trim(),//port code
                                                        rdr[4].ToString().Trim(),// tran code
                                                        rdr[2].ToString().Trim(),// sec type
                                                        rdr[3].ToString().Trim(), //symbol
                                                        (DateTime)rdr[5], // trade date
                                                         (DateTime)rdr[6], // settle date   
                                                         (Double)rdr[7], // quantity
                                                         rdr[8].ToString().Trim(), // close meth
                                                          (Double)rdr[9], // td fx rate
                                                           (Double) rdr[10], // sd fx rate
                                                            (Double)rdr[11] // trade amt
                                                        )                                       
                                );    
                            screen.AppendText(String.Format("\r\nFound Trade: {0} for sec type: {1}.", rdr[8].ToString().Trim(), rdr[1].ToString().Trim()));
                        }
                    }
                }// eo using

                }
            catch (Exception e)
            {
              
                screen.AppendText(Globals.saveErr(GetCurrentMethod() + ":" + e.Message + "\r\n"));
                Globals.WriteErrorLog(e.ToString());
                throw e;
            }

            return trades;
        }//eof

        public int getISOCurrency(string aSecType, ref string aISOCode)
        {
            int rtn = 0;
           
            try
            {
                using (SqlConnection conn = new SqlConnection(dbConnection))
                {
                    conn.Open();

                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand("usp_GetISOCurrency", conn);

                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;

                    // 3. add parameter to command, which will be passed to the stored procedure
                    if (aSecType.Length == 4)
                        cmd.Parameters.Add(new SqlParameter("@cur", aSecType.Substring(2,2)  ));
                    else
                        cmd.Parameters.Add(new SqlParameter("@cur", aSecType.Substring(0,2) ));


                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {

                        if (rdr.HasRows)
                        {
                            // iterate through results, printing each to console
                            while (rdr.Read())
                            {
                                rtn++;
                                aISOCode = rdr[0].ToString().Trim();
                                screen.AppendText( String.Format("\r\nFound ISO Code: {0} for sec type: {1}.", aISOCode, aSecType));
                            }
                        }


                    } // end of using
                } // end of outter using

                if (String.IsNullOrEmpty(aISOCode))
                {
                    screen.AppendText( Globals.saveErr(String.Format("\r\n--> No ISO Code found for sec type: {0}.", aSecType)));
                }
            }
            catch (Exception e)
            {
                rtn = -1;
                screen.AppendText( Globals.saveErr(GetCurrentMethod() + ":" +  e.Message + "\r\n"));
                Globals.WriteErrorLog(e.ToString());
            }


            if (String.IsNullOrEmpty(aISOCode))
            {
                
                screen.AppendText(  Globals.saveErr(String.Format("\r\nISO Code not found for sec type: {0}.", aSecType)));
               
            }

            return rtn;
        }

        protected int getISOCashXref(string aSecType, ref string aISOCode)
        {
            int rtn = 0;
            string msg = string.Empty;
            try
            {
                aISOCode = string.Empty;
                using (SqlConnection conn = new SqlConnection(dbConnection))
                {
                    conn.Open();

                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand("usp_GetISOCashXref", conn);

                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;

                    // 3. add parameter to command, which will be passed to the stored procedure
                    cmd.Parameters.Add(new SqlParameter("@sectype", aSecType));

                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {

                        if (rdr.HasRows)
                        {
                            // iterate through results, printing each to console
                            while (rdr.Read())
                            {
                                rtn++;
                                aISOCode = rdr[0].ToString().Trim() ;
                                screen.AppendText( String.Format("\r\nFound ISO Code: {0} for sec type: {1}.", aISOCode, aSecType));
                            }
                        }
                      
                       
                    } // end of using
                } // end of connection using

                if (String.IsNullOrEmpty(aISOCode))
                {
                   
                    screen.AppendText( Globals.saveErr(String.Format("\r\n--->Unknown ISO Code: {0} for sec type: {1}.   ? ? ?", aISOCode, aSecType)));
                }

            }
            catch (Exception e)
            {
                rtn = -1;
                screen.AppendText( Globals.saveErr(e.Message + "\r\n"));
                Globals.WriteErrorLog(e.ToString());
            }
                                    
            return rtn;
        }


        public int flipRate(string aCurrency1, string aCurrency2, ref Boolean aFlipRate)
        {
            int rtn = 0;
            int flip = 0;
           
            try
            {

               
                using (SqlConnection conn = new SqlConnection(dbConnection))
                {
                    conn.Open();

                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand("usp_FXConFlipRate", conn);

                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;

                    // 3. add parameter to command, which will be passed to the stored procedure
                    cmd.Parameters.Add(new SqlParameter("@currency1", aCurrency1));
                    cmd.Parameters.Add(new SqlParameter("@currency2", aCurrency2));
                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {

                        if (rdr.HasRows)
                        {
                            // iterate through results, printing each to console
                            while (rdr.Read())
                            {
                                rtn++;
                                flip  = (int) rdr[0];
                                if (flip == 1) { aFlipRate = true; }
                                screen.AppendText(String.Format("\r\nFlip rate for : {0} and {1}.", aCurrency1, aCurrency2));
                            }
                        }

                    } // end of using

                    

                }

             
            }
            catch (Exception e)
            {
                rtn = -1;
                screen.AppendText(Globals.saveErr(GetCurrentMethod() + ": " + e.Message + " ? ? ? \r\n"));
                Globals.WriteErrorLog(e.ToString());
            }

            return rtn;  
        }

        protected int getConversionInstructions(string aCurrency1, string aCurrency2, int aInstructionType, ref string aConversionInstruction)
        {
            int rtn = 0;
           
            try
            {

                using (SqlConnection conn = new SqlConnection(dbConnection))
                {
                    conn.Open();

                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand("usp_GetConversionInstruction02", conn);

                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;

                    // 3. add parameter to command, which will be passed to the stored procedure
                    cmd.Parameters.Add(new SqlParameter("@cur1", aCurrency1));
                    cmd.Parameters.Add(new SqlParameter("@cur2", aCurrency2));
                    cmd.Parameters.Add(new SqlParameter("@instructiontype", aInstructionType));
                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {

                        if (rdr.HasRows)
                        {
                            // iterate through results, printing each to console
                            while (rdr.Read())
                            {
                                rtn++;
                                aConversionInstruction = rdr[0].ToString();
                                screen.AppendText(String.Format("\r\nFound conversion instruction -{0}- for : {1} and {2}.", aConversionInstruction, aCurrency1, aCurrency2 ));
                            }
                        }                      

                    } // end of using
                }

                if (String.IsNullOrEmpty(aConversionInstruction) ) {
                    screen.AppendText( Globals.saveErr(String.Format("\r\n-->Conversion instruction not found for : {0} to {1}.  ? ? ? ", aCurrency1, aCurrency2)));
                    screen.AppendText(Globals.saveErr(String.Format("\r\n-->===>Add conversion instruction in Moxy table tb_CurrencyConversionInstructions <<<=== ")));
                }

            }
            catch (Exception e)
            {
                rtn = -1;
                screen.AppendText(Globals.saveErr(GetCurrentMethod() + ": " +  e.Message + " ? ? ? \r\n"));
                Globals.WriteErrorLog(e.ToString());
            }
            
            return rtn;  
        } //  end of getConversionInstuctions()

        /// <summary>
        ///     getCrossRate Function:  retrieives cross rate of the trade from Moxy.
        ///     
        ///     Note: The cross rate stored in MoxyOrders table in UserDef2 field.
        ///                When Moxy trades get exported to .trn file this field is not
        ///                included in the file. For each trade we have to explicitly retreive
        ///                Moxy for the cross rate.
        /// </summary>
        /// <param name="aTradeMatchId">a unique trade identifier</param>
        /// <param name="aTradingCur">portfolio's trading currency</param>
        /// <param name="aCrossRate">a cross rate</param>
        /// <returns>0/-1</returns>
        public int getCrossRate (int aTradeMatchId, string aTradeDate, string aTradingCur, string aSecType, string aPortfolio, ref string aCrossRate, ref string aConversionInstruction) 
        {
            int rtn = 0;
         
            //double number;
            string secType = string.Empty;
            string secISOCode = string.Empty;
            
            string errMsg = string.Empty;
            string securityCur = string.Empty;
            string tradeDate = string.Empty;
       

            try
            {
              
                using (SqlConnection conn = new SqlConnection(dbConnection))
                {
                    conn.Open();
                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand("usp_GetCrossRate", conn);
                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;
                    // 3. add parameter to command, which will be passed to the stored procedure
                    cmd.Parameters.Add(new SqlParameter("@tradematchid", aTradeMatchId));
                    cmd.Parameters.Add(new SqlParameter("@portfolio", aPortfolio));
                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        // iterate through results
                        while (rdr.Read())
                        {
                            aCrossRate = rdr[0].ToString() ;
                            secType = rdr[1].ToString();
                            tradeDate = rdr[2].ToString();                  
                                                  
                        } // end of while loop

                        //
                        // when cross rate is unavailable from Moxy try to get it from Portia 
                        //
                        if (String.IsNullOrEmpty(secType )) { secType = aSecType;}

                        getISOCurrency(secType, ref securityCur);
                        if (String.IsNullOrEmpty(aCrossRate))
                        {
                           
                            if (PortiaDatabase.getLastCrossRate(securityCur, aTradingCur, aTradeDate, ref aCrossRate) == -1)
                            {
                                screen.AppendText(String.Format(GetCurrentMethod() + "--->No cross rate for : {0} {1} {2} {3} {4} in Portia", securityCur, aTradingCur, tradeDate));
                            }
                        }

                       
                        // get security ISO Code
                        //if (getISOCurrency(aSecType, ref secISOCode) != -1 && getConversionInstructions(secISOCode, aTradingCur, ref conversionInstruction) != -1)
                        if (getISOCurrency(aSecType, ref secISOCode) != -1 && getConversionInstructions(securityCur, aTradingCur,0, ref aConversionInstruction) != -1)
                        {
                            //screen.AppendText(String.Format("\r\nFound cross rate: {0} for trade match id: {1}.", aCrossRate, aTradeMatchId));

                        }    // end of if                           


                    } // end of using
                } // end of outter using
            } // end of try
            catch (Exception e)
            {
                rtn = -1;
                screen.AppendText(  Globals.saveErr( e.Message + "\r\n"));
                Globals.WriteErrorLog(e.Message);
            }
                                    
            return rtn;
        } // end of getCrossRate()

        // Difference from getCrossRate() : sectype goes before trading currency
        public int getCrossRate02(int aTradeMatchId, string aTradeDate, string aSecType, string aTradingCur,string aPortfolio, ref string aCrossRate, ref string aConversionInstruction)
        {
            int rtn = 0;

            //double number;
            string secType = string.Empty;
            string secISOCode = string.Empty;
            string errMsg = string.Empty;
            string securityCur = string.Empty;
            string tradeDate = string.Empty;
            string tradingCur = string.Empty;  

            try
            {

                using (SqlConnection conn = new SqlConnection(dbConnection))
                {
                    conn.Open();
                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand("usp_GetCrossRate", conn);
                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;
                    // 3. add parameter to command, which will be passed to the stored procedure
                    cmd.Parameters.Add(new SqlParameter("@tradematchid", aTradeMatchId));
                    cmd.Parameters.Add(new SqlParameter("@portfolio", aPortfolio));
                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        // iterate through results
                        while (rdr.Read())
                        {
                            aCrossRate = rdr[0].ToString();
                            secType = rdr[1].ToString();
                            tradeDate = rdr[2].ToString();

                        } // end of while loop

                        //
                        // when cross rate is unavailable from Moxy try to get it from Portia 
                        //
                        //if (String.IsNullOrEmpty(secType)) { secType = aSecType; }
                        secType = aSecType;
                        getISOCurrency(aTradingCur , ref tradingCur);
                        if (String.IsNullOrEmpty(aCrossRate))
                        {

                            if (PortiaDatabase.getLastCrossRate(secType, tradingCur, aTradeDate, ref aCrossRate) == -1)
                            {
                                screen.AppendText(String.Format(GetCurrentMethod() + "--->No cross rate for : {0} {1} {2} {3} {4} in Portia", securityCur, aTradingCur, tradeDate));
                            }
                        }

                        // get security ISO Code
                        //if (getISOCurrency(aSecType, ref secISOCode) != -1 && getConversionInstructions(secISOCode, aTradingCur, ref conversionInstruction) != -1)
                        if (getISOCurrency(aSecType, ref secISOCode) != -1 && getConversionInstructions(secType, tradingCur,0, ref aConversionInstruction) != -1)
                        {
                            //screen.AppendText(String.Format("\r\nFound cross rate: {0} for trade match id: {1}.", aCrossRate, aTradeMatchId));

                        }    // end of if                           


                    } // end of using
                } // end of outter using
            } // end of try
            catch (Exception e)
            {
                rtn = -1;
                screen.AppendText(Globals.saveErr(e.Message + "\r\n"));
                Globals.WriteErrorLog(e.Message);
            }

             return rtn;
        } // end of getCrossRate()

        public int getCrossRateCash(int aTradeMatchId, string aTradeDate, string aSecType, string aTradingCur, string aPortfolio, ref string aCrossRate, ref string aConversionInstruction)
        {
            int rtn = 0;

            //double number;
            string secType = string.Empty;
            string secISOCode = string.Empty;
            string errMsg = string.Empty;
            string securityCur = string.Empty;
            string tradeDate = string.Empty;
            string tradingCur = string.Empty;

            try
            {

                using (SqlConnection conn = new SqlConnection(dbConnection))
                {
                    conn.Open();
                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand("usp_GetCrossRateCash", conn);
                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;
                    // 3. add parameter to command, which will be passed to the stored procedure
                    cmd.Parameters.Add(new SqlParameter("@tradematchid", aTradeMatchId));
                    cmd.Parameters.Add(new SqlParameter("@portfolio", aPortfolio));
                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        // iterate through results
                        while (rdr.Read())
                        {
                            aCrossRate = rdr[0].ToString();
                            secType = rdr[1].ToString();
                            tradeDate = rdr[2].ToString();

                        } // end of while loop

                        //
                        // when cross rate is unavailable from Moxy try to get it from Portia 
                        //
                        //if (String.IsNullOrEmpty(secType)) { secType = aSecType; }
                        secType = aSecType;
                        getISOCurrency(aTradingCur, ref tradingCur);
                        if (String.IsNullOrEmpty(aCrossRate))
                        {

                            if (PortiaDatabase.getLastCrossRate(secType, tradingCur, aTradeDate, ref aCrossRate) == -1)
                            {
                                screen.AppendText(String.Format(GetCurrentMethod() + "--->No cross rate for : {0} {1} {2} {3} {4} in Portia", securityCur, aTradingCur, tradeDate));
                            }
                        }

                        // get security ISO Code
                        //if (getISOCurrency(aSecType, ref secISOCode) != -1 && getConversionInstructions(secISOCode, aTradingCur, ref conversionInstruction) != -1)
                        if (getISOCurrency(aSecType, ref secISOCode) != -1 && getConversionInstructions(secType, tradingCur,1, ref aConversionInstruction) != -1)
                        {
                            //screen.AppendText(String.Format("\r\nFound cross rate: {0} for trade match id: {1}.", aCrossRate, aTradeMatchId));

                        }    // end of if                           


                    } // end of using
                } // end of outter using
            } // end of try
            catch (Exception e)
            {
                rtn = -1;
                screen.AppendText(Globals.saveErr(e.Message + "\r\n"));
                Globals.WriteErrorLog(e.Message);
            }

            return rtn;
        } // end of getCrossRateCash()



        /// <summary>
        ///     getTradingCurrency function:  retreives portfolio' s trading currency.
        ///     
        ///     Note: Trading currency might differ from reporting currency.
        /// </summary>
        /// <param name="aTradingCurrencyStoredProc">name of the stored procedure to retreive trading currency</param>
        /// <param name="aPortfolio">portfolio number</param>
        /// <param name="aTradingCurrency">portfolio's trading currency</param>
        /// <returns>0/-1</returns>
        public int getTradingCurrency(string aTradingCurrencyStoredProc, string aPortfolio, ref string aTradingCurrency)
        {
            int rtn = 0;
            string errMsg = string.Empty;
            try
            {
                using (SqlConnection conn = new SqlConnection(dbConnection))
                {
                    conn.Open();

                    // 1.  create a command object identifying the stored procedure
                    SqlCommand cmd = new SqlCommand(aTradingCurrencyStoredProc , conn);

                    // 2. set the command object so it knows to execute a stored procedure
                    cmd.CommandType = CommandType.StoredProcedure;

                    // 3. add parameter to command, which will be passed to the stored procedure
                    cmd.Parameters.Add(new SqlParameter("@portfolio", aPortfolio));

                    // execute the command
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        // iterate through results, printing each to console
                        while (rdr.Read())
                        {
                            aTradingCurrency = rdr[0].ToString();
                            // flip cross rate

                            if (string.IsNullOrWhiteSpace(aTradingCurrency))
                            {
                                errMsg = String.Format("\r\n" + GetCurrentMethod() + ": Can not retrieve trading currency for : {0} for trade match id: {1}.", aTradingCurrency);
                                screen.AppendText( Globals.saveErr(errMsg));

                            }

                            else
                                screen.AppendText( String.Format("\r\nFound trading currency {0} for : {1} ", aTradingCurrency, aPortfolio));
                        } // end of while loop
                    } // end of using
                } // end of outter using
            } // end of try
            catch (Exception e)
            {
                rtn = -1;
                errMsg = "Function " + GetCurrentMethod() + ":" + e.Message + "\r\n";
                screen.AppendText(Globals.saveErr(errMsg + "\r\n"));
                screen.AppendText( Globals.saveErr(errMsg));

                Globals.WriteErrorLog(e.ToString());
            }

            return rtn;
        } // end of getTradeCurrency()

        public string GetCurrentMethod()
        {
            StackTrace st = new StackTrace();
            StackFrame sf = st.GetFrame(1);

            return sf.GetMethod().Name;
        }

        /// <summary>
        /// get the files extracted from Portia to import into Moxy
        /// </summary>
        /// <param name="storedProc"></param>
        /// <returns></returns>
        public DataTable getSrcFiles(string storedProc) {
            DataTable dt = null;
            try
            {
                dt = new DataTable();
                using (var con = new SqlConnection(dbConnection)) 
                using (var cmd = new SqlCommand(storedProc, con))
                using (var da = new SqlDataAdapter(cmd))
                {
                    cmd.CommandType = CommandType.StoredProcedure;
                    da.Fill(dt);
                }

                // set primary key on first column
                DataColumn[] keyColumns = new DataColumn[1];
                keyColumns[0] = dt.Columns["id"];
                dt.PrimaryKey = keyColumns;

            }
             catch (Exception e)
            {
              
                string errMsg = "Function " + GetCurrentMethod() + ":" + e.Message + "\r\n";
                screen.AppendText(Globals.saveErr(errMsg + "\r\n"));
                screen.AppendText(Globals.saveErr(errMsg));
                Globals.WriteErrorLog(e.ToString());
                throw new Exception(errMsg);
            }

            return dt;
        }// eof


    } // end of class
} // end of namespace
