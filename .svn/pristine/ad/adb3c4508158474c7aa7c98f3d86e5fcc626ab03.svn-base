using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
namespace PortiaMoxyImport
{
    class TradeEvare
    {
        public string[] items = null;
        TextBox Screen;
        String reconConnStr = null;
        public TradeEvare(TextBox aScreen, string aLine, string aDbConnRecon)
        {
            
            try
            {
                Screen = aScreen;
                reconConnStr = aDbConnRecon;
                items = aLine.Split(',');
            }
            catch (Exception ex)
            {

                Screen.AppendText(String.Format("{0}: {1}", GetCurrentMethod(), ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }
        }

        public int convert()
        {

            string funcName = string.Empty;

            try
            {

            }
            catch (Exception ex)
            {

                Screen.AppendText(String.Format("{0}: {1}", GetCurrentMethod(), ex.Message));
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

    } // end of class

   
}
