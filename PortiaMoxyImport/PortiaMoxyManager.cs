using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Xml;
using System.Data;
using System.Collections;
using System.Text.RegularExpressions;
using System.Configuration;

//
//  converts files from Portia to Moxy Import Format
//

namespace PortiaMoxyImport
{
    class PortiaMoxyManager
    {
    
        public TextBox screen; // for communication with UI
        public Label status;      // for showing progress of the task
        public PortiaMoxyManager(ref TextBox aScreen, ref Label aStatus)
        {
            screen = aScreen;
            status = aStatus;
        }

        public int checkSourceDestination(String aInFile, String aOutFile)
        {
            //
            //   checks if source & destination files exist 
            // 
            int rtn = 0;

            try
            {
                if (!File.Exists(aInFile))
                {
                    rtn = -1;
                    screen.AppendText( String.Format("?-?-? Source file {0} not found ?-?-?", aInFile)); 
                }

                if (!File.Exists(aOutFile))
                {
                    rtn = -1;
                    screen.AppendText( String.Format("?-?-? Destination file {0} not found ?-?-?", aOutFile));
                }


            }
            catch (Exception ex)
            {
                screen.AppendText( "checkSourceDestination: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;
        }

        public int checkMandatoryValues(string aLine)
        {
            //
            // check if the first two positions in the string have values
            //
            int rtn = 0;
            try
            {
                string[] values = aLine.Split('\t');
                if (values[0] == String.Empty)
                {
                    screen.AppendText( "!!!--->  checkMandatoryValue: Values are missing for " + aLine + Environment.NewLine);
                    return -1;
                }

                
                if (values.Length > 1 && values[0] == String.Empty)
                {
                    screen.AppendText( "!!!--->  checkMandatoryValues: Values are missing for " + aLine + Environment.NewLine);
                    return -1;
                }
                    
             }
            catch (Exception ex)
            {
                screen.AppendText( "checkMandatoryValues: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;

        }


        public int convertGeneric(String aInFile, String aOutFile)
        {
            //
            // - removes double quotes & replaces commas with Tab in a file
            // - returns number of line processed
            //
                        
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            var myList = new List<string>();

            try
            {
                // Open the file to read from. 
                string[] arr = File.ReadAllLines(aInFile);

                foreach (string s in arr)
                {
                    Application.DoEvents();

                    if (s.IndexOf("No Data") != -1) { continue;  } // do not include this line

                    rtn += 1;
                    tmp = s.Replace("\"", "");
                    tmp = tmp.Replace(",", "\t");
                    // check for the values in first and second positions
                    if (checkMandatoryValues(tmp) == -1)
                    {
                        return -1;
                    }

                    myList.Add(tmp);
                    rtn += 1;
                    status.Text = String.Format("Row # {0}", rtn.ToString()); 
                }

                File.WriteAllLines(aOutFile, myList.ToArray());
            }
            catch (Exception ex)
            {
                screen.AppendText( "convertGeneric: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;
        }

        public  int convertHoliday (String aInFile, String aOutFile)
        {
            int rtn = 0;       
            string tmp = string.Empty;  
            string line = string.Empty; 
            var myList = new List<string>();

            try{
                screen.AppendText( "######################" + Environment.NewLine ); 
                screen.AppendText( "###   Holiday Conversion   ###\r\n");
                screen.AppendText( "######################" + Environment.NewLine); 

                rtn = convertGeneric(aInFile, aOutFile);

                if (rtn != -1)
                {
                   screen.AppendText( String.Format("{0} holidays loaded into file {1}\r\n", (rtn - 1 ).ToString(), aOutFile ));
                   screen.AppendText( Environment.NewLine);
                }

              
            }
            catch (Exception ex) {
                screen.AppendText( "convertHoliday: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

           
            return rtn;
        } // end of convertHoliday function

        /// <summary>
        ///  convertBrokers18() - for Moxy 18
        /// </summary>
        /// <param name="aInFile"></param>
        /// <param name="aOutFile"></param>
        /// <returns></returns>
        public int convertBrokers18(String aInFile, String aOutFile)
        {
            int rtn = 0;
                       
            try
            {
                writeHeaderOnScreen("   Broker Conversion   ");

                string reqColNo = ConfigurationManager.AppSettings["BrokerImportRequiredCols"];

                if (isValidColNumber(aInFile, Int32.Parse(reqColNo))) {
                    rtn = convertGeneric(aInFile, aOutFile);
                }
                else
                {
                    rtn = -1;
                }

                if (rtn != -1)
                {
                    screen.AppendText(String.Format("{0} brokers loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText(Environment.NewLine);
                }
                else
                {
                    screen.AppendText("===> Failed to convert broker file. ? ? ?");
                    screen.AppendText(Environment.NewLine);
                }

             
            }
            catch (Exception ex)
            {
                screen.AppendText("convertBrokers: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;

        } // end of convertBrokers18
        
        public int convertBrokers(String aInFile, String aOutFile)
        {
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            var myList = new List<string>();

            try
            {
                writeHeaderOnScreen("   Broker Conversion   ");

                rtn = convertGeneric(aInFile, aOutFile);

                if (rtn != -1)
                {
                    screen.AppendText( String.Format("{0} brokers loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText( Environment.NewLine);
                }


            }
            catch (Exception ex)
            {
                screen.AppendText( "convertBrokers: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;

        } // end of convertBrokers

        public int convertSectors(String aInFile, String aOutFile)
        {
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            var myList = new List<string>();

            try
            {
                screen.AppendText("######################" + Environment.NewLine);
                screen.AppendText("###  Sector Conversion   ###\r\n");
                screen.AppendText("######################" + Environment.NewLine);

                rtn = convertGeneric(aInFile, aOutFile);

                if (rtn != -1)
                {
                    screen.AppendText(String.Format("{0} Sectors loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText(Environment.NewLine);
                }


            }
            catch (Exception ex)
            {
                screen.AppendText("convertSectors: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;

        } // end of convertSectors

        public int convertIndustry(String aInFile, String aOutFile)
        {
            int rtn = 0;
            string tmp = string.Empty;
           // string line = string.Empty;
            var myList = new List<string>();

            try
            {
                writeHeaderOnScreen("Industry Group Conversion");

                //rtn = convertGeneric(aInFile, aOutFile);

                // Open the source file.
                string[] arr = File.ReadAllLines(aInFile);
                foreach (string s in arr)
                {
                    Application.DoEvents();
                    String str =replaceAllInsideQuotesCommasWithTildas(s);
                    if (s.IndexOf("No Data") != -1) { continue; } // do not include this line

                    // split the string to array
                    String[] indArr = str.Split(',');
               
                    if (indArr.Length == 8)
                    {
                         // make Industry Group Name to containg IndGrpId and Ind Grp Name 
                        indArr[4] = indArr[0] + "  " + indArr[4];

                        // convert array back to string
                        str = String.Join(",", indArr);
                    }
                    // replace double quotes, commas, and tildas back to commas
                    tmp = str.Replace("\"", "").Replace(",", "\t").Replace('~', ',');
           
                    // check for the values in first and second positions
                    if (checkMandatoryValues(tmp) == -1)  {   return -1; }

                    myList.Add(tmp);
                    rtn += 1;
                    status.Text = String.Format("Row # {0}", rtn.ToString());
                }

                File.WriteAllLines(aOutFile, myList.ToArray());

                if (rtn != -1)
                {
                    screen.AppendText(String.Format("{0} industries loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText(Environment.NewLine);
                }

            }
            catch (Exception ex)
            {
                screen.AppendText("convertIndustry: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;

        } // end of convertIndustry


        /// <summary>
        ///  could be multiple group files from portia
        /// </summary>
        /// <param name="aInFile"></param>
        /// <param name="aOutFile"></param>
        /// <returns></returns>
        public int convertGroups18(String aInFile, String aOutFile)
        {
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            string fileGrpName = string.Empty;
            string dir = string.Empty;
            //var myList = new List<string>();

            try
            {
                writeHeaderOnScreen("Industry Group Conversion");

                string reqColNo = ConfigurationManager.AppSettings["GroupImportRequiredCols"];

                // get the incoming group file name without extension
                fileGrpName = Path.GetFileNameWithoutExtension(aInFile);
                dir = Path.GetDirectoryName(aInFile);

                // find all group files & combine them into one
                string[] files = Directory.GetFiles(dir, "*" + fileGrpName + "*", SearchOption.TopDirectoryOnly);
                string comboFile = aInFile.Replace(fileGrpName, "combo");

                if (File.Exists(comboFile))
                    File.Delete(comboFile);

                // read groups from all files
                foreach (string f in files)
                {

                    string[] arr = File.ReadAllLines(f);
                    screen.AppendText(String.Format("Read {0} lines from the file {1} \r\n", arr.Length, f));
                    File.AppendAllLines(comboFile, arr);
                }

                rtn = convertGeneric(comboFile, aOutFile);

                if (rtn != -1)
                {
                    screen.AppendText(String.Format("{0} groups loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText(Environment.NewLine);
                }


            }
            catch (Exception ex)
            {
                screen.AppendText("convertGroups: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;
        } // end of convertGroups function



        /// <summary>
        ///  could be multiple group files from portia
        /// </summary>
        /// <param name="aInFile"></param>
        /// <param name="aOutFile"></param>
        /// <returns></returns>
        public int convertGroups(String aInFile, String aOutFile)
        {
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            string fileGrpName = string.Empty;
            string dir = string.Empty;
            //var myList = new List<string>();

            try
            {
                screen.AppendText( "#########################" + Environment.NewLine); 
                screen.AppendText( "###   Groups Conversion   ###\r\n");
                screen.AppendText( "#########################" + Environment.NewLine);

                // get the incoming group file name without extension
                fileGrpName = Path.GetFileNameWithoutExtension(aInFile);
                dir = Path.GetDirectoryName(aInFile);

                // find all group files & combine them into one
                string[] files = Directory.GetFiles(dir, "*" + fileGrpName + "*", SearchOption.TopDirectoryOnly);
                string comboFile = aInFile.Replace(fileGrpName, "combo");

                if (File.Exists(comboFile))
                    File.Delete(comboFile);

                // read groups from all files
                foreach (string f in files)
                {

                    string[] arr = File.ReadAllLines(f);
                    screen.AppendText(String.Format("Read {0} lines from the file {1} \r\n",arr.Length , f));
                    File.AppendAllLines(comboFile , arr);
                }

                rtn = convertGeneric(comboFile, aOutFile);

                if (rtn != -1)
                {
                    screen.AppendText( String.Format("{0} groups loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText( Environment.NewLine);   
                }


            }
            catch (Exception ex)
            {
                screen.AppendText( "convertGroups: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;
        } // end of convertGroups function

               
        public int convertSecType(String aInFile, String aOutFile)
        {
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            var myList = new List<string>();
            string[] arr;
            int requiredCols = 76;
            string errLine = string.Empty;

            try {
                screen.AppendText("######################" + Environment.NewLine);
                screen.AppendText("###   Sec Type Conversion   ###\r\n");
                screen.AppendText("######################" + Environment.NewLine);

                StreamWriter sw = new StreamWriter(aOutFile);
                List<String> secTypeList = new List<String>();

                List<string> columns = new List<string>();
                using (var reader = new CsvFileReader(aInFile))
                {
                    
                    while (reader.ReadRow(columns))
                    {
                        Application.DoEvents();
                                               
                        arr = columns.ToArray();
                        if (arr[arr.Length - 1].Trim() != "SECTYPE")
                        {
                            // validate that the line has been split into columns = requiredColumns
                            if (arr.Length != requiredCols)
                            {
                                foreach (string col in columns)
                                {
                                    errLine += col + " ";
                                }
                                screen.AppendText("convertSecType ERROR: " + String.Format("Line {0} could not be split into {1} fields required for Moxy import\r\n", errLine, requiredCols.ToString()));
                                rtn = -1;
                                return rtn;
                            }

                            line = String.Empty;

                            // remove double quotes
                            for (int i = 0; i <= (arr.Length - 1); i++)
                            {
                                arr[i] = arr[i].Replace("\"", String.Empty);
                            }

                            // combine sec type in one column
                            arr[0] = arr[0].ToLower() + arr[1].ToLower();
                            arr[1] = String.Empty;
                            arr[arr.Length - 1] = "?"; //  last field

                            line = String.Empty;
                            for(int i=0; i<arr.Length;i++)
                            {
                                if(i!=arr.Length-1)
                                    line += arr[i] + "\t";
                                else
                                    line += arr[i] ;
                            }
                              

                           // shoud remove the last TAB character ???

                        }
                        else
                        {
                            line = arr[0];
                        }
                                               
                        //sw.WriteLine(line);
                        secTypeList.Add(line);
                        rtn += 1;
                        status.Text = String.Format("Sec Type # {0}", rtn.ToString());

                    }  // end of while

                } // end of using

                // remove douplicates
                secTypeList = secTypeList.Distinct().ToList();

                rtn = secTypeList.Count;
                // write sec types to output file
                foreach(String s in secTypeList)
                    sw.WriteLine(s);

                sw.Close();
                if (rtn != -1)
                {
                    screen.AppendText(String.Format("{0} sec types loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText(Environment.NewLine);
                }

            }
            catch (Exception ex)
            {
                screen.AppendText("convertSecType: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;
        }

            public int convertPrice(String aInFile, String aOutFile)
        {
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            var myList = new List<string>();
            string[] arr;
            int requiredCols = 5;
            string errLine = string.Empty; 

            try
            {
                screen.AppendText( "######################" + Environment.NewLine); 
                screen.AppendText( "###   Price Conversion     ###\r\n");
                screen.AppendText( "######################" + Environment.NewLine); 

                //rtn = convertGeneric(aInFile, aOutFile);
                 StreamWriter sw = new StreamWriter(aOutFile); 

                List<string> columns = new List<string>();
                using (var reader = new CsvFileReader(aInFile))
                {
                    while (reader.ReadRow(columns))
                    {
                        Application.DoEvents();  
                        
                        arr = columns.ToArray();
                        if (arr[arr.Length - 1].Trim() != "PRICE")
                        {
                            // validate that the line has been split into columns = requiredColumns
                            if (arr.Length != requiredCols)
                            {
                                foreach (string col in columns)
                                {
                                    errLine += col + " ";
                                }
                                screen.AppendText( "convertPrice ERROR: " + String.Format("Line {0} could not be split into {1} fields required for Moxy import\r\n", errLine, requiredCols.ToString()));
                                rtn = -1;
                                return rtn;
                            }

                            line = String.Empty;

                            // remove double quotes
                            for (int i = 0; i <= (arr.Length - 1); i++)
                            {
                                arr[i] = arr[i].Replace("\"", String.Empty);
                            }

                            // construct line in Moxy TSV format
                            line = arr[0] + arr[1].ToString().ToLower()   + "\t";                      // moxy security type
                            line += arr[2] + "\t" + arr[3] + "\t" + arr[4] + "\t";
                            line +=  "\t\t\t";                                                                             // terminate the line with three tabs so the line is in 8 columns Moxy format
                            
                        } 
                        else
                        {
                            line = arr[0]; 
                        }

                        // check for the values in first and second positions
                        if (checkMandatoryValues(line) == -1)
                        {
                            return -1;
                        }

                        sw.WriteLine(line);
                        rtn += 1;
                        status.Text = String.Format("Group # {0}", rtn.ToString()); 

                    }  // end of while
                } // end of using
                sw.Close();
                if (rtn != -1)
                {
                    screen.AppendText( String.Format("{0} prices loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText( Environment.NewLine);  
                }


            }
            catch (Exception ex)
            {
                screen.AppendText( "convertPrice: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;
        } // end of convertPrice function

        public int convertCustodian(String aInFile, String aOutFile)
        {
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            var myList = new List<string>();

            try
            {
                screen.AppendText( "######################" + Environment.NewLine);
                screen.AppendText( "###   Custodian Conversion   ###\r\n");
                screen.AppendText( "######################" + Environment.NewLine);

                // Open the file to read from. 
                string[] arr = File.ReadAllLines(aInFile);
                foreach (string s in arr)
                {
                    Application.DoEvents();
                    rtn += 1;
                    tmp = s.Replace("\"", "");
                    if (tmp != "CUST")
                    {
                        tmp = tmp.Replace(",", "\t");
                        // replace & with &amp; so Moxy exepts it.
                        //tmp = tmp.Replace("&", "&amp;");
                    }


                    // check for the values in first and second positions
                    if (checkMandatoryValues(tmp) == -1)
                    {
                        return -1;
                    }

                    // check for the values in first and second positions
                    if (checkMandatoryValues(tmp) == -1)
                    {
                        return -1;
                    }

                    myList.Add(tmp);
                    rtn += 1;
                    status.Text = String.Format("Row # {0}", rtn.ToString());
                }

                File.WriteAllLines(aOutFile, myList.ToArray());
                                             
                screen.AppendText( String.Format("{0} custodians loaded into file {1}\r\n", arr.Length.ToString(), aOutFile));
                screen.AppendText( Environment.NewLine);
 
            }
            catch (Exception ex)
            {
                screen.AppendText( "convertCustodian: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }

            return rtn;
        }

        public int convertCurrency(String aInFile, String aOutFile)
        {
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            var myList = new List<string>();

            try
            {
                screen.AppendText( "######################" + Environment.NewLine); 
                screen.AppendText( "###   Currency Conversion   ###\r\n");
                screen.AppendText( "######################" + Environment.NewLine);

                // Open the file to read from. 
                string[] arr = File.ReadAllLines(aInFile);

                foreach (string s in arr)
                {
                    Application.DoEvents();
                    rtn += 1;
                                       
                    tmp = s.Replace("\"", "");
                    tmp = tmp.Replace(",", "\t");

                    // adjustment for Chech Currency
                    //tmp = ReplaceFirst(tmp, "CZ", "CS");

                    // check for the values in first and second positions
                    if (checkMandatoryValues(tmp) == -1)
                    {
                        return -1;
                    }

                    myList.Add(tmp);
                    rtn += 1;
                    status.Text = String.Format("Row # {0}", rtn.ToString());
                }

                File.WriteAllLines(aOutFile, myList.ToArray());


                //rtn = convertGeneric(aInFile, aOutFile);

                //if (rtn != -1)
                //{
                    screen.AppendText( String.Format("{0} currencies loaded into file {1}\r\n", arr.Length.ToString(), aOutFile));
                    screen.AppendText( Environment.NewLine);
                //}

            }
            catch (Exception ex)
            {
                screen.AppendText( "convertCurrency: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }


            return rtn;
        } // end of convertCurrency function


        public Tuple<int, HashSet<string>> convertPortfolio18(String aInFile, String aOutFile)
        {

            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            var myList = new List<string>();
            char[] delimiterChars = { ',' };
            int requiredCols = 66;
            string[] arr;
            DataTable dt = new DataTable();
            string port = string.Empty;
            ArrayList duplicates = new ArrayList();
            HashSet<string> hsPortfolios = new HashSet<string>();
            try
            {
                // use datatable to find duplicate portfolios
               dt.Columns.Add("portfolio");
                writeHeaderOnScreen("     Portfolio Conversion     ");
                string reqColNo = ConfigurationManager.AppSettings["PortfolioImportRequiredCols"];
                StreamWriter sw = new StreamWriter(aOutFile, false, Encoding.GetEncoding("windows-1250"));

                List<string> columns = new List<string>();
                using (var reader = new CsvFileReader(aInFile))
                {
                    while (reader.ReadRow(columns))
                    {
                        arr = columns.ToArray();

                        // replace last elment of the array with "0"
                        if (arr[arr.Length - 1].Trim() != "PORTFOLIO")
                        {

                            arr[arr.Length - 1] = "0";
                            // validate that the line has been split into 69 columns
                            if (arr.Length != Int32.Parse(reqColNo))
                            {
                                screen.AppendText("convertPortfolio ERROR: " + String.Format("Line {0} could not be split into {1} fields required for Moxy import\r\n", line, requiredCols.ToString()));
                                rtn = -1;
                            }

                            line = String.Empty;
                            port = arr[0].ToString();
                            hsPortfolios.Add(port);

                            foreach (string col in arr)
                            {

                                tmp = col.Replace("\"", String.Empty);
                                line += tmp + '\t';

                            } // end FOR loop
                            //line = line.TrimEnd();
                        } // end of IF
                        else
                        {
                            line = arr[0];
                        }

                        // check for the values in first and second positions
                        if (checkMandatoryValues(line) == -1)
                        {
                            return new Tuple<int, HashSet<string>>(-1, null);
                        }


                        sw.WriteLine(line);
                        DataRow row = dt.NewRow();
                        row[0] = port;
                        dt.Rows.Add(row);


                        rtn += 1;

                    } // end WHILE loop
                } // end of usinng

                sw.Close();

                duplicates = FindDuplicateRows(dt, "portfolio");

                foreach (DataRow d in duplicates)
                {

                    screen.AppendText(String.Format("Duplicate portfolio in portfolio source file - {0}\r\n", d[0]));
                    screen.AppendText(Environment.NewLine);
                    rtn = -1;
                    screen.ScrollToCaret();
                    screen.Focus();
                }


                if (rtn != -1)
                {
                    screen.AppendText(String.Format("{0} portfolios loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText(Environment.NewLine);
                }



            } // end of try
            catch (Exception ex)
            {
                screen.AppendText("convertPortfolio: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }
            return new Tuple<int, HashSet<string>>(rtn, hsPortfolios);


        } // end of convertPortfolio function


        public Tuple<int, HashSet<string >>  convertPortfolio(String aInFile, String aOutFile)
        {
            
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            var myList = new List<string>();
            char[] delimiterChars = { ',' };
            int requiredCols = 66;
            string[] arr;
            DataTable dt = new DataTable();
            string port = string.Empty;
            ArrayList duplicates = new ArrayList();
            HashSet<string> hsPortfolios = new HashSet<string>(); 
            try
            {
                // use datatable to find duplicate portfolios
                dt.Clear();
                dt.Columns.Add("portfolio");

                printHeader("Portfolio Conversion");

                StreamWriter sw = new StreamWriter(aOutFile, false, Encoding.GetEncoding("windows-1250"));

                List<string> columns = new List<string>();
                using (var reader = new CsvFileReader(aInFile))
                {
                    while (reader.ReadRow(columns))
                    {
                        arr = columns.ToArray();

                        // replace last elment of the array with "0"
                        if (arr[arr.Length - 1].Trim() != "PORTFOLIO")
                        {

                            arr[arr.Length - 1] = "0";
                            // validate that the line has been split into 67 columns
                            if (arr.Length != requiredCols)
                            {
                                screen.AppendText("convertPortfolio ERROR: " + String.Format("Line {0} could not be split into {1} fields required for Moxy import\r\n", line, requiredCols.ToString()));
                                rtn = -1;
                            }

                            line = String.Empty;
                            port = arr[0].ToString();
                            hsPortfolios.Add(port);

                            foreach (string col in arr)
                            {

                                tmp = col.Replace("\"", String.Empty);
                                line += tmp + '\t';

                            } // end FOR loop
                            //line = line.TrimEnd();
                        } // end of IF
                        else
                        {
                            line = arr[0];
                        }

                        // check for the values in first and second positions
                        if (checkMandatoryValues(line) == -1)
                        {
                            return new Tuple<int, HashSet<string>>(-1, null);
                        }


                        sw.WriteLine(line);
                        DataRow row = dt.NewRow();
                        row[0] = port;
                        dt.Rows.Add(row);


                        rtn += 1;

                    } // end WHILE loop
                } // end of usinng

                sw.Close();

                duplicates = FindDuplicateRows(dt, "portfolio");

                foreach (DataRow d in duplicates)
                {

                    screen.AppendText(String.Format("Duplicate portfolio in portfolio source file - {0}\r\n", d[0]));
                    screen.AppendText(Environment.NewLine);
                    rtn = -1;
                    screen.ScrollToCaret();
                    screen.Focus();
                }


                if (rtn != -1)
                {
                    screen.AppendText(String.Format("{0} portfolios loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText(Environment.NewLine);
                }



            } // end of try
            catch (Exception ex)
            {
                screen.AppendText( "convertPortfolio: " + ex.Message + Environment.NewLine);
                rtn = -1;
            }
            return new Tuple<int, HashSet<string>>( rtn, hsPortfolios );


        } // end of convertPortfolio function

        private void printHeader(string headerText)
        {
            string midLine = "###   " + headerText + " ###";
            string topLine = string.Empty;
            string bottomLine = string.Empty ;

            for (int i=0; i< midLine.Length; i++)
            {
                topLine += "#";
                bottomLine += "#";
            }
                        
            screen.AppendText(topLine + Environment.NewLine);
            screen.AppendText("###   " +headerText +  " ###" + Environment.NewLine);
            screen.AppendText(bottomLine + Environment.NewLine);
        }

        public int convertSecurity(String aInFile, String aOutFile)
        {
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            var myList = new List<string>();
            char[] delimiterChars = { ',' };
            string[] arr;
            int max = 0;
            int runningMax = 0;
            int exch = 0;
            int requiredCols = 22;
            try
            {
                printHeader("Convert Security");

                DateTime dt = File.GetLastWriteTime(aInFile );
                StreamWriter sw = new StreamWriter(aOutFile); 

                //screen.AppendText( System.Reflection.MethodBase.GetCurrentMethod().Name;
                // create XML header
                line = @"<?xml version=""1.0""?>";
                sw.WriteLine(line);
                line = @"<AdventXML version=""3.0"">";
                sw.WriteLine(line);
                line = @"<AccountProvider name=""Portia"">";
                sw.WriteLine(line);
                line = @"<Imports/>";
                sw.WriteLine(line);
                line = @"<InterpretFXRates/>";
                sw.WriteLine(line);
                line = String.Format(@"<SCXList date=""{0}"">", dt.ToString("yyyyMMdd"));
                sw.WriteLine(line);
                List<string> columns = new List<string>();
                using (var reader = new CsvFileReader(aInFile))
                {
                    while (reader.ReadRow(columns))
                    {
                        // TODO: Do something with columns' values
                        // validate that the line has been split into number of required columns
                        arr = columns.ToArray();
                        if (arr[arr.Length - 1].Trim() != "SECURITY")
                        {
                            
                            // validate that the line has been split into number of required  columns
                            if (arr.Length != requiredCols)
                            {
                                screen.AppendText( "convertSecurity ERROR: " + String.Format("Line {0} could not be split into {1} fields required for Moxy import\r\n", line, requiredCols.ToString()));
                                return  -1;
                            }

                            line = String.Empty;

                            for(int i=0; i <= (arr.Length-1); i++ )
                            {

                                arr[i]  = arr[i].ToString().Replace("\"", String.Empty);
                                if (arr[i] == null || arr[i] == "null") { 
                                    arr[i] = string.Empty; 
                                }

                                // <SCX type="ad" iso="USD" symbol="204445209" fullname="COMPANIA DE ALUMBRADO ELEC-SAN SAL" cusip="204445209" country="us" exch="004" axyssecuserdef1id="100" axyssecuserdef2id="0" axyssecuserdef3id="3" astcode="e" shortassetclass="e" isin="" sedol="" issuer="us"/>
                                line = String.Format("<SCX type=\"{0}\" ", arr[0].Trim());
                                line += String.Format("iso=\"{0}\" ", arr[2].Trim());
                                line += String.Format("symbol=\"{0}\" ", arr[3].Trim());
                                                               

                                // replace & with &amp; so Moxy exepts it.
                                String fullName = arr[4].Trim().Replace("&", "&amp;") ;

                                line += String.Format("fullname=\"{0}\" ", fullName );
                                
                                try
                                {
                                  //  if (String.Format("fulname=\"{0}\" ", arr[4].Trim()).Length > max)
                                   // {
                                        // max length of sec name
                                        //max = String.Format("fullname=\"{0}\" ", arr[4].Trim()).Length;
                                        max = arr[4].Trim().Length;
                                    if (max > runningMax)
                                        runningMax = max;
                                       if (max  > 39 ) { screen.AppendText(this.GetType().FullName + ":  Security: " + arr[4] + " Length: " + arr[4].Trim().Length  + Environment.NewLine); }
                                   // }

                                }
                                catch (Exception e)
                                {
                                    screen.AppendText( this.GetType().FullName + ": " + e.Message + Environment.NewLine);
                                }
                             

                                if (arr[5] != string.Empty)
                                {
                                    line += String.Format("cusip=\"{0}\" ", arr[5]);
                                }
                                line += String.Format("country=\"{0}\" ", arr[6]);

                              

                                if (arr[7] != string.Empty)
                                {
                                    exch = Convert.ToInt32(arr[7]);
                                    line += String.Format("exch=\"{0}\" ", exch.ToString());
                                }
                                                      
                                if (arr[8] != string.Empty )
                                {
                                    line += String.Format("axyssecuserdef1id=\"{0}\" ", arr[8]);
                                }
                                if (arr[9] != string.Empty)
                                {
                                    line += String.Format("axyssecuserdef2id=\"{0}\" ", arr[9]);
                                }
                                if (arr[10] != string.Empty)
                                {
                                    line += String.Format("axyssecuserdef3id=\"{0}\" ", arr[10]);
                                }
                                                             
                                line += String.Format("astcode=\"{0}\" ", arr[11]);
                                line += String.Format("shortassetclass=\"{0}\" ", arr[12]);
                                if (arr[13] != string.Empty)
                                {
                                    line += String.Format("isin=\"{0}\" ", arr[13]);
                                }
                                if (arr[14] != string.Empty)
                                {
                                    line += String.Format("sedol=\"{0}\" ", arr[14]);
                                }
                                line += String.Format("issuer=\"{0}\" ", arr[15]);

                                if (arr[16] != string.Empty)
                                {
                                    line += String.Format("maturity=\"{0}\" ", arr[16]);
                                }

                                if (arr[17] != string.Empty)
                                {
                                    line += String.Format("sector=\"{0}\" ", arr[17]);
                                }

                                if (arr[18] != string.Empty)
                                {
                                    line += String.Format("indgrp=\"{0}\" ", arr[18]);
                                }

                                if (arr[19] != string.Empty)
                                {
                                    if (Double.Parse(arr[19]) > 0)
                                        line += String.Format("shareout=\"{0}\" ", arr[19]);
                                    else
                                        line += String.Format("shareout=\"{0}\" ", "0.001");
                                }

                                if (arr[20] != string.Empty)
                                {
                                    line += String.Format("fixenabled=\"{0}\" ", arr[20]);
                                }

                                if (arr[21] != string.Empty)
                                {
                                    line += String.Format("userdef1=\"{0}\" ", arr[21]);
                                }

                                //else
                                //{
                                    line += "/> ";
                                //}

                                //line =String.Format ( @"<SCX type="{0}" iso="{1}" symbol="{2}" fullname="{3}" cusip="{4}" country="{5}" exch="{6}" axyssecuserdef1id="{7}" axyssecuserdef2id="{8}" axyssecuserdef3id="{9}" astcode="{10}" shortassetclass="{11}" isin="{12}" sedol="{13}" issuer="{14}"/>\n", arr[0] ); 
                            } // end FOR loop

                            sw.WriteLine(line);
                            rtn += 1;
                        } // end of IF
                             

                    } // end WHILE loop
                } // end of usinng


                // create XML footer
                 line = @"</SCXList>";
                 sw.WriteLine(line);
                 line = @"</AccountProvider>";
                 sw.WriteLine(line);
                 line = @"</AdventXML>";
                 sw.WriteLine(line);
               
                sw.Close();
                if (rtn != -1)
                {
                    screen.AppendText( String.Format("{0} securities loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText( Environment.NewLine);
                    screen.AppendText( String.Format("Max length of sec name {0}\r\n", runningMax.ToString()));
                    screen.AppendText( Environment.NewLine);   
                    status.Text = "Ready";
                }

            }

            catch (Exception ex)
            {
                screen.AppendText( this.GetType().FullName + ": " + ex.Message + Environment.NewLine);
                rtn = -1;
            }
            return rtn;
        } // end of convertSecurity function

        public int convertTaxLots(String aInFile, String aOutFile, HashSet<string> aHSPortfolios)
        {
            int rtn = 0;
            string tmp = string.Empty;
            string line = string.Empty;
            var myList = new List<string>();
            char[] delimiterChars = { ',' };
            int requiredCols = 29;
            string[] arr;
            HashSet<string> hsPortsWithHld = new HashSet<string>();     // portfolios with holdings
            HashSet<string> hsPortsNoHld = new HashSet<string>();       // portfolios without holdings
            try
            {

                screen.AppendText( "##########################" + Environment.NewLine);
                screen.AppendText( "###   Tox Lots Conversion   ###\r\n");
                screen.AppendText( "#########################" + Environment.NewLine); 

                DateTime dt = File.GetLastWriteTime(aInFile);
                StreamWriter sw = new StreamWriter(aOutFile);

                //screen.AppendText( System.Reflection.MethodBase.GetCurrentMethod().Name;
                // create XML header
                line = @"<?xml version=""1.0""?>";
                sw.WriteLine(line);
                line = @"<AdventXML version=""3.0"">";
                sw.WriteLine(line);
                line = @"<AccountProvider name=""PortiaTaxLotFile"">";
                sw.WriteLine(line);
                line = @"<Imports/>";
                sw.WriteLine(line);
                line = @"<InterpretFXRates/>";
                sw.WriteLine(line);
                line = String.Format(@"<LTXList date=""{0}"">", dt.ToString("yyyyMMdd"));
                sw.WriteLine(line);
                List<string> columns = new List<string>();
                using (var reader = new CsvFileReader(aInFile))
                {
                    while (reader.ReadRow(columns))
                    {
                        
                        arr = columns.ToArray();
                        if (arr[arr.Length - 1].Trim() != "TAXLOT")
                        {

                            // validate that the line has been split into number of required columns
                            if (arr.Length != requiredCols)
                            {
                                screen.AppendText( "convertToxLots ERROR: " + String.Format("Line {0} could not be split into {1} fields required for Moxy import\r\n", line, requiredCols.ToString()));
                                return -1;
                            }

                            line = String.Empty;

                            for (int i = 0; i <= (arr.Length - 1); i++)
                            {
                                Application.DoEvents();  
                                arr[i] = arr[i].ToString().Replace("\"", String.Empty);
                                if (arr[i] == null || arr[i] == "null")
                                {
                                    arr[i] = string.Empty;
                                }
                                                                
                                line = String.Format("<LTX portfolio=\"{0}\" ", arr[1]);
                                hsPortsWithHld.Add(arr[1].ToString()); 
                                line += String.Format("type=\"{0}\" ", arr[2].ToLower() + arr[3].ToLower()  );
                                line += String.Format("symbol=\"{0}\" ", arr[4]);
                                line += String.Format("postype=\"{0}\" ", arr[5]);
                                line += String.Format("hldate=\"{0}\" ", arr[6]);
                                line += String.Format("ocdate=\"{0}\" ", arr[7]);
                                line += String.Format("quantity=\"{0}\" ", arr[8]);
                                line += String.Format("totalcost=\"{0}\" ", arr[9]);
                                line += String.Format("pledge=\"{0}\" ", "n");
                                line += String.Format("lotnum=\"{0}\" ", arr[12]);
                                line += String.Format("custodian=\"{0}\" ", "254");
                                line += String.Format("iszeromv=\"{0}\" ", "0");
                                line += String.Format("userdef1=\"{0}\" ", arr[16]);
                                line += String.Format("broker=\"{0}\"/>", arr[10]);
                                
                            } // end FOR loop

                            sw.WriteLine(line);
                            rtn += 1;
                            status.Text = String.Format("Tax Lot # {0}", rtn.ToString());
                                                                                 

                        } // end of IF
                                            

                    } // end WHILE loop
                } // end of usinng

                // find portfolios without holdings & create zero holdings
                hsPortsNoHld = aHSPortfolios; 
                hsPortsNoHld.ExceptWith(hsPortsWithHld);

                if (hsPortsNoHld.Count > 0 )
                {
                    foreach(string s in hsPortsNoHld )
                    {
                        line = string.Empty;
                        line = String.Format("<LTX portfolio=\"{0}\" ", s);
                        line += String.Format("type=\"{0}\" ", "caus");
                        line += String.Format("symbol=\"{0}\" ", "-USD CASH-");
                        line += String.Format("postype=\"{0}\" ", "0");
                        line += String.Format("hldate=\"{0}\" ", "");
                        line += String.Format("ocdate=\"{0}\" ", "");
                        line += String.Format("quantity=\"{0}\" ", "0");
                        line += String.Format("totalcost=\"{0}\" ", "0");
                        line += String.Format("pledge=\"{0}\" ", "n");
                        line += String.Format("lotnum=\"{0}\" ", "0");
                        line += String.Format("custodian=\"{0}\" ", "254");
                        line += String.Format("iszeromv=\"{0}\" ", "0");
                        line += String.Format("userdef1=\"{0}\" ", "No");
                        line += String.Format("broker=\"{0}\"/>", "");
                        sw.WriteLine(line);
                        rtn += 1;
                        status.Text = String.Format("Zero Holdings Tax Lot # {0}", rtn.ToString());
                    }

                }

                // create XML footer
                line = @"</LTXList>";
                sw.WriteLine(line);
                line = @"</AccountProvider>";
                sw.WriteLine(line);
                line = @"</AdventXML>";
                sw.WriteLine(line);

                sw.Close();
                if (rtn != -1)
                {
                    screen.AppendText( String.Format("{0} tax lots loaded into file {1}\r\n", (rtn - 1).ToString(), aOutFile));
                    screen.AppendText( Environment.NewLine);   
                    status.Text = "Ready";
                }

            }

            catch (Exception ex)
            {
                screen.AppendText( this.GetType().FullName + ": " + ex.Message + Environment.NewLine);
                rtn = -1;
            }
            return rtn;
        } // end of convertTaxLots function

        public string ReplaceFirst(string text, string search, string replace)
        {
            // 
            // replaces onlt the first occurence of the search in the text
            //
            int pos = text.IndexOf(search);
            if (pos < 0)
            {
                return text;
            }
            return text.Substring(0, pos) + replace + text.Substring(pos + search.Length);
        }

        public ArrayList FindDuplicateRows(DataTable dTable, string colName)
        {
            Hashtable hTable = new Hashtable();
            ArrayList duplicateList = new ArrayList();

            //Add list of all the unique item value to hashtable, which stores combination of key, value pair.
            //And add duplicate item value in arraylist.
            foreach (DataRow drow in dTable.Rows)
            {
                if (hTable.Contains(drow[colName]))
                    duplicateList.Add(drow);
                else
                    hTable.Add(drow[colName], string.Empty);
            }

  
            //array list contains dups.
            return duplicateList;
        }

        public void writeHeaderOnScreen(String message)
        {
            String midLine = "### " + message + " ###\r\n";
            string coverLine =new  String('#', midLine.Length) + Environment.NewLine;

            screen.AppendText(coverLine);
            screen.AppendText(midLine);
            screen.AppendText(coverLine);
        }

        public String replaceAllInsideQuotesCommasWithTildas(String inputStr)
        {
            String rtn = null;
            bool openQuote = false;
            try
            {
                foreach (char c in inputStr )
                {
                    if (c == '"')
                    {
                        if (openQuote == false)
                            openQuote = true;
                        else
                            openQuote = false;
                    }

                    if (c == ',' && openQuote)
                    {
                        rtn += '~';
                    }
                    else
                        rtn += c;

                }// eol
                
            }
             catch (Exception ex)
            {
                screen.AppendText("convertPrice: " + ex.Message + Environment.NewLine);
            }
            return rtn;
        }// eof
        
       /// <summary>
       /// isValidColNumber() - checks if the import file has required by Moxy specs
       ///                                      column number
       ///                                         
       /// </summary>
       /// <param name="file"></param>
       /// <param name="reqColNo"></param>
       /// <returns>true/false</returns>
        bool isValidColNumber(string file, int reqColNo)
        {
            bool rtn = false;
            try
            {
                // Open the file to read from. 
                string[] arr = File.ReadAllLines(file);

                if (arr == null || arr.Length == 0)
                    throw new Exception("isValidColNumber Error: Input file " + file + " is empty" );

                if (arr.Length > 1)
                {
                    // split the first non header string to items
                    String[] indArr = arr[1].Split(',');
                    if (indArr.Length == reqColNo)
                        rtn = true;
                    else
                    {
                        screen.AppendText("isValidColNumber Error: " +  String.Format(" Column number in broker file does not match. Required columns: {0}, provided: {1}  ", reqColNo, indArr.Length));
                    }
                }
                
             }
             catch (Exception ex)
            {
                screen.AppendText("isValidColNumber Error: " + ex.Message + Environment.NewLine);
            }

            return rtn;
        }
        
                
    } // end of class
} // end of namespace
