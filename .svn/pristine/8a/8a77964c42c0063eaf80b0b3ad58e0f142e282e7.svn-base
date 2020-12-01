using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using System.Diagnostics;
using System.Windows.Forms;
using System.Text.RegularExpressions;


namespace PortiaMoxyImport
{
    class EvareFile
    {
        TextBox Screen;
        String reconConnStr = null;
        String evareFileID = null;
        String filesLocation = null;

        public  EvareFile(TextBox aScreen, String aReconConnStr, String aEvareFileID, String aFilesLocation)
        {
            Screen = aScreen;
            reconConnStr = aReconConnStr;
            evareFileID = aEvareFileID;
            filesLocation = aFilesLocation;
        }

        public string GetCurrentMethod()
        {
            StackTrace st = new StackTrace();
            StackFrame sf = st.GetFrame(1);

            return sf.GetMethod().Name;
        }

        /// <summary>
        ///     getFiles: retrieves Evare files from the specified folder
        /// </summary>
        /// <param name="directory"></param>
        /// <returns></returns>
        public ArrayList getFiles()
        {
            ArrayList files = new ArrayList();
            
            try
            {
                var directory = new DirectoryInfo(filesLocation);
                // all files updated within one day
                DateTime from_date = DateTime.Now.AddDays(-1);
                DateTime to_date = DateTime.Now;
                foreach (var f in directory.GetFiles().Where(file => file.LastWriteTime >= from_date && file.LastWriteTime <= to_date))
                {
                    if (isTranFile(evareFileID, f.Name ) ) {
                        files.Add(f.Name);
                        Screen.AppendText(f.Name);
                        Screen.AppendText(Environment.NewLine);
                    }
                    
                }

                Screen.AppendText("Number of valid Evare files: " + files.Count);
                Screen.AppendText(String.Empty);
            }
            catch (Exception ex)
            {
               
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }
            return files;
        }

        public bool isTranFile(String aEvareFileID, String aFileName)
        {
            bool rtn = false;
            try
            {
                if (aFileName.Contains(aEvareFileID) ){rtn = true;} 
            }
            catch (Exception ex)
            {

                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }
            return rtn;
        }

        public void convertToAIM(String fileName)
        {
            String line = null;
            try
            {

                Screen.AppendText(getBankNameFromFile(filesLocation + fileName));
                Screen.AppendText(Environment.NewLine);

                Stream stream1 = new FileStream(filesLocation + fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                StreamReader sr1 = new StreamReader(stream1);
                do
                {
                    line = sr1.ReadLine();
                    if (String.IsNullOrEmpty(line))
                        continue;

                    TradeEvare trade = new TradeEvare(Screen, line, reconConnStr);

                } while (!(line == null));


                sr1.Close();

            }
            catch (Exception ex)
            {

                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }
        }

        /// <summary>
        ///     getBankName: extract a bank name from the file name
        /// </summary>
        /// <param name="aEvareFileID"></param>
        /// <param name="aFileName"></param>
        /// <returns></returns>
        public String getBankNameFromFile(String aFileName) {
            String bankName = null;
            try
            {
                bankName = Path.GetFileNameWithoutExtension(aFileName); 
                bankName = bankName.Replace(evareFileID, String.Empty);
                bankName = Regex.Replace(bankName, "[^a-zA-Z]", "");
            }
            catch (Exception ex)
            {
                Screen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }


            return bankName;
        }


    }//end of class
}// end of namesapce
