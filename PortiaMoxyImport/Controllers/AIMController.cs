using PortiaMoxyImport.Entities;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PortiaMoxyImport.Controllers
{
    class AIMController
    {
        TextBox tbScreen;
        PortiaDatabase portia;
        MetaData mData;
        MoxyDatabase moxy;
        public AIMController(TextBox screen, PortiaDatabase pd)
        {
            tbScreen = screen;
            portia = pd;
            mData = getAppMetaDataMoxy();
            //md = new MoxyDatabase(mData.moxyConStr, screen);
            moxy = new MoxyDatabase(mData.moxyConStr, screen);
        }// eo constructor

        private MetaData getAppMetaDataMoxy()
        {
            try
            {
                string moxycon = Util.getAppConfigVal("moxyconstr");
                string metadatastoredproc = Util.getAppConfigVal("metadataSP");
                MetaData mData = null;

                SqlConnection sqlConnection1 = new SqlConnection(moxycon);
                SqlCommand cmd = new SqlCommand();
                SqlDataReader reader;

                cmd.CommandText = metadatastoredproc;
                cmd.Connection = sqlConnection1;

                sqlConnection1.Open();

                reader = cmd.ExecuteReader();
                // Data is accessible through the DataReader object here.
                while (reader.Read())
                {
                    mData = new MetaData(reader[0].ToString(), reader[1].ToString(), reader[2].ToString(),
                                                            reader[3].ToString(), reader[4].ToString(), reader[5].ToString(),
                                                            reader[6].ToString(), moxycon, reader[7].ToString(), reader[8].ToString(),
                                                            reader[9].ToString());

                }

                reader.Close();
                sqlConnection1.Close();

                return mData;
            }
            catch (Exception ex)
            {
                Globals.errCnt += 1;
                tbScreen.AppendText("\r\n" + GetCurrentMethod() + ": " + ex.Message);
                Globals.WriteErrorLog(ex.ToString());
                throw new Exception("\r\n" + GetCurrentMethod() + ": " + ex.Message);
            }
        }// eof

        public void convertMoxyTrades(List<TrnLine> trades)
        {
            
            foreach (TrnLine t in trades)
            {

                // based on these values the appropriate conversion function selected
                //if (t.Symbol.Equals("$cash")) { tranType = "cash"; } else { tranType = "equity"; }
                //if (tradeCur.Equals("USD")) { portType = "Us"; } else { portType = "NonUs"; }
                //if (securityCur.Equals(tradeCur)) { secType = "SameCur"; } else { secType = "DiffCur"; }
                //if ((items[1].Equals("sl") || items[1].Equals("SL"))) { tranCode = "Sell"; } else { tranCode = "Buy"; }
                
                String line = convertToAIM(t);

            }
        }

        private string convertToAIM(TrnLine t)
        {
            String tradeCur = null;
            String tranType;
            String securityCur = null;
            String repCur = null;
            String tranCode;
            String portType = null;
            String secType = null;
            try {

                
                int rtn = portia.getTradingCurrency(mData.tradingCurrencyStoredProc, t.PortCode, ref tradeCur);
                if (String.IsNullOrEmpty(tradeCur)) { tradeCur = "USD"; }

                rtn = moxy.getISOCurrency(t.Symbol, ref securityCur);

                rtn = portia.getReportingCurrency(mData.reportingCurrencyStoredProc, t.PortCode, ref repCur);

                // based on these values the appropriate conversion function selected
                if (t.Symbol.Equals("$cash")) { tranType = "cash"; } else { tranType = "equity"; }
                if (tradeCur.Equals("USD")) { portType = "Us"; } else { portType = "NonUs"; }
                if (securityCur.Equals(tradeCur)) { secType = "SameCur"; } else { secType = "DiffCur"; }
                if ((t.TranCode.Equals("sl") || t.TranCode.Equals("SL"))) { tranCode = "Sell"; } else { tranCode = "Buy"; }



            }
            catch (Exception ex)
            {
                tbScreen.AppendText(Globals.saveErr(GetCurrentMethod() + ":" + ex.Message + "\r\n"));
            }

            return "";
        }

        private string GetCurrentMethod()
        {
            StackTrace st = new StackTrace();
            StackFrame sf = st.GetFrame(1);

            return sf.GetMethod().Name;
        }


    }// eo class
}
