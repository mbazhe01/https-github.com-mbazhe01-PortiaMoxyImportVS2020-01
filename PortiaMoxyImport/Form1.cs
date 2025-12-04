using PortiaMoxyImport.Entities;
using PortiaMoxyImport.Forms;
using PortiaMoxyImport.Services;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;  
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PortiaMoxyImport
{

    //
    //  PortiaMoxyImport application:
    //                                                              1. create holdings files from Portia to import into Moxy
    //                                                              2. convert Moxy trades to import into Portia
    //                                                              3. convert FX connect FX trades to import into Portia
    //
    //  TO DO:   
    //                                                              1. Error counting & reporting
    //                                                              2. Moxy Trades for AIM
    



    public partial class Form1 : Form
    {
        string sqlitedb = "moxyimport.sqlite";
        const int RNDNUM = 8; // round to 8 decimal places
        const string fileNotFound = "File {0} not found please run Moxy Export";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //
            // TEST UNIT 01: testing conversion functions
            //
            String inFile;
            String outFile;
            
            PortiaMoxyManager pm = new PortiaMoxyManager(ref  rtbScreen, ref lblStatus );

            inFile = @"J:/Temp/PortiaMoxyImport/FromPortia/holiday.csv";
            outFile = @"J:/Temp/PortiaMoxyImport/MoxyFiles/holiday.tsv";

            if (pm.convertHoliday(inFile, outFile ) == -1) {
                return; 
            }

            inFile = @"J:/Temp/PortiaMoxyImport/FromPortia/groups.csv";
            outFile = @"J:/Temp/PortiaMoxyImport/MoxyFiles/groups.tsv";

            if (pm.convertGroups(inFile, outFile) == -1)
            {
                return;
            }

            inFile = @"J:/Temp/PortiaMoxyImport/FromPortia/price.csv";
            outFile = @"J:/Temp/PortiaMoxyImport/MoxyFiles/price.tsv";
                     

            if (pm.convertPrice(inFile, outFile) == -1)
            {
                return;
            }

            inFile = @"J:/Temp/PortiaMoxyImport/FromPortia/currency.csv";
            outFile = @"J:/Temp/PortiaMoxyImport/MoxyFiles/currency.tsv ";

            if (pm.convertCurrency(inFile, outFile) == -1)
            {
                return;
            }

            inFile = @"J:/Temp/PortiaMoxyImport/FromPortia/portfolio.csv";
            outFile = @"J:/Temp/PortiaMoxyImport/MoxyFiles/portfolio.tsv ";


            //if (pm.convertPortfolio(inFile, outFile) == -1)
            //{
            //    return;
            //}

            inFile = @"J:/Temp/PortiaMoxyImport/FromPortia/security.csv";
            outFile = @"J:/Temp/PortiaMoxyImport/MoxyFiles/secinfo.scx ";


            if (pm.convertSecurity(inFile, outFile) == -1)
            {
                return;
            }


            inFile = @"J:/Temp/PortiaMoxyImport/FromPortia/taxlot.csv";
            outFile = @"J:/Temp/PortiaMoxyImport/MoxyFiles/taxlot1.ltx ";


            //if (pm.convertTaxLots(inFile, outFile) == -1)
            //{
            //    return;
            //}



            // scroll to the end of text box
            tbScreen.SelectionStart = tbScreen.TextLength;
            tbScreen.ScrollToCaret();
        }

        /// <summary>
        /// Portia Holdings for Moxy
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
           

            //SQLiteDatabase db;
            string fType= string.Empty;     // file tyypes: holiday, groups, price, currency, portfolio, security, taxlot, broker
            string inPath = string.Empty;   // input file path  
            string outPath = string.Empty; // output file path  
            HashSet<string> hsPortfolios = null; // to hold all portfolios coming from Portia
            try
            {
                tbScreen.Clear();  

                PortiaMoxyManager pm = new PortiaMoxyManager(ref  rtbScreen, ref lblStatus);
                MoxyDatabase md = new MoxyDatabase(Util.getAppConfigVal("moxyconstr"), rtbScreen);

                DataTable inFiles = md.getSrcFiles(Util.getAppConfigVal("getPortiaSrcFilesSP"));
                DataTable outFiles = md.getSrcFiles(Util.getAppConfigVal("getMoxyImportFilesSP"));

                // get rid of sqlite database -> use Moxy 1/24/2019 MB

                // input files
                //db = new SQLiteDatabase(sqlitedb);
                //DataTable inFiles;
                //String query = "select ID \"id\", VALUE \"value\"";
                //query += "from infiles;";
                //inFiles = db.GetDataTable(query);
                
               // output files 
                //DataTable outFiles;
                //query = "select ID \"id\", VALUE \"value\"";
                //query += "from outfiles;";
                //outFiles = db.GetDataTable(query);

                // loop through each input file
                foreach (DataRow r in inFiles.Rows)
                {
                    fType=r["id"].ToString();
                    inPath=r["value"].ToString();
                    tbScreen.AppendText( String.Format("File Type: {0} Source Path: {1} ", fType, inPath) + Environment.NewLine);

                    string s = fType ;
                    DataRow foundRow =  outFiles.Rows.Find(s);

                    if (foundRow != null)
                    {
                        tbScreen.AppendText( String.Format("File Type: {0} Destination Path: {1} ", foundRow[0] , foundRow[1]) + Environment.NewLine);
                        outPath = foundRow[1].ToString() ;
                        Application.DoEvents();
                        // File conversion from Portia to Moxy
                        switch (fType)
                        {
                            case "holiday":
                                if (pm.convertHoliday(inPath, outPath) == -1)
                                {
                                    return;
                                }
                                break;
                            case "groups":
                                if (pm.convertGroups(inPath, outPath) == -1)
                                {
                                    return;
                                }
                                break;
                            case "price":
                                if (pm.convertPrice(inPath, outPath) == -1)
                                {
                                    return;
                                }
                                break;
                            case "currency":
                                 if (pm.convertCurrency(inPath, outPath) == -1)
                                {
                                    return;
                                }
                                break;
                            case "portfolio":

                                Tuple<int, HashSet<string>> rt = pm.convertPortfolio(inPath, outPath);
                                   if (rt.Item1 == -1) { return; }
                                   else { hsPortfolios = rt.Item2; }
                                    break;
                            case "security":
                                    if (pm.convertSecurity(inPath, outPath) == -1)
                                    {
                                        return;
                                    }
                                    break;
                            case "taxlot":
                                    if (pm.convertTaxLots(inPath, outPath, hsPortfolios) == -1)
                                    {
                                        return;
                                    }
                                    break;
                            case "custodian":
                                    if (pm.convertCustodian(inPath, outPath) == -1)
                                    {
                                        return;
                                    }
                                    break;
                            case "broker":
                                    if (pm.convertBrokers(inPath, outPath) == -1)
                                    {
                                        return;
                                    }
                                    break;
                            case "sector":
                                if (pm.convertSectors(inPath, outPath) == -1)
                                {
                                    return;
                                }
                                break;
                            case "industry":
                                if (pm.convertIndustry(inPath, outPath) == -1)
                                {
                                    return;
                                }
                                break;
                            case "sectype":
                                if (pm.convertSecType(inPath, outPath) == -1)
                                {
                                    return;
                                }
                                break;
                            default:
                                tbScreen.AppendText( "Default case" + Environment.NewLine );
                                break;
                        }



                        // scroll to the end of text box
                        tbScreen.SelectionStart = tbScreen.TextLength;
                        tbScreen.ScrollToCaret();
                        lblStatus.Text = "Ready";

                    }
                    else
                    {
                        tbScreen.AppendText( String.Format("A row with the primary key of {0} could not be found in {1} ", s, "outfiles")); 
                        return; 
                    }
                }

                // scroll to the end of text box
                tbScreen.SelectionStart = tbScreen.TextLength;
                tbScreen.ScrollToCaret();
            }
            catch (Exception fail)
            {
                String error = "The following error has occurred:\n\n";
                error += fail.Message.ToString() + "\n\n";
                MessageBox.Show(error);
                Globals.WriteErrorLog(fail.ToString());
                this.Close();
            }

        }

        private void btnPortiaFiles_Click(object sender, EventArgs e)
        {
            SQLiteDatabase db;
            string connStr = string.Empty;
            string sql = "usp_getmoxyexport";
            DateTime asOfDate= DateTime.Today; 
            //int rtn;
            string outFolder = string.Empty;

            try {

                tbScreen.Clear();  

                // Moxy Connection
                db = new SQLiteDatabase(sqlitedb);
                DataTable moxyConn;
                String query = "select ID \"id\", VALUE \"value\"";
                query += " from moxy where id=\"connstr\";";
                moxyConn  = db.GetDataTable(query);

                foreach (DataRow r in moxyConn.Rows)
                {
                    connStr = r["value"].ToString();
                    tbScreen.AppendText( String.Format("Moxy Connection: {0}" , connStr ) + Environment.NewLine);
                }

                // get Moxy output folder
                query = "select ID \"id\", VALUE \"value\"";
                query += " from moxy where id=\"outFolder\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    outFolder = r["value"].ToString();
                    tbScreen.AppendText( String.Format("Moxy Export Output Folder: {0}", outFolder) + Environment.NewLine);
                }

                // ask user for a date
                DateTime today = DateTime.Now;

                DateTime  input = DateTime.Now;
                string userinput = input.ToString("MM/dd/yyyy");
                ShowInputDialog(ref userinput );
                if (DateTime.TryParse(userinput, out asOfDate))
                {
                    // it's a recognized as a DateTime                  
                }
                else
                {
                    // it's not recognized as a DateTime
                    String error = "The following error has occurred:\n\n";
                    error += userinput + " is not a valid date" + "\n\n";
                    MessageBox.Show(error);
                    return;
                }

             
                // get trades from Moxy
                 SqlConnection conn = new SqlConnection(connStr);
                SqlDataAdapter da = new SqlDataAdapter();
                da.SelectCommand = new SqlCommand(sql, conn);
                da.SelectCommand.CommandType = CommandType.StoredProcedure;

                da.SelectCommand.Parameters.Add("@asofdate", SqlDbType.DateTime ).Value = asOfDate;   

                
                DataSet ds = new DataSet();
                da.Fill(ds, "result_name");

                DataTable dt = ds.Tables["result_name"];

                SaveToText(dt, outFolder + "MoxyExport" + asOfDate.ToString("yyyyMMdd") + ".tsv");
               
                conn.Close();  

            }
            catch (FormatException ex)
            {
                //String error = "The following error has occurred:\n\n";
                //error += userinput + " is not a valid date" + "\n\n";
                MessageBox.Show(ex.Message );
                Globals.WriteErrorLog(ex.ToString());
                this.Close();
            }
            catch (Exception fail)
            {
                String error = "The following error has occurred:\n\n";
                error += fail.Message.ToString() + "\n\n";
                MessageBox.Show(error);
                Globals.WriteErrorLog(fail.ToString());
                this.Close();
            }
          
        }
               
        public int SaveToText(DataTable dt, String filePath)
        {
            int rtn=0;
            String ln = string.Empty;  
            int cnt = 0;

            try {
                   StreamWriter sw = new StreamWriter(filePath);
                    // write column header
                    foreach (DataColumn  col in dt.Columns)
                    {
                         ln += col.ColumnName + "\t";
                    }
                    sw.WriteLine(ln);
                    foreach (DataRow row in dt.Rows )
                    {
                         ln =string.Empty;
                        foreach (DataColumn col in dt.Columns )
                        {
                            ln += row[col].ToString() + "\t";
                        }
                        sw.WriteLine(ln);
                        cnt += 1;
                    }
                    sw.Close(); 
                     tbScreen.AppendText("\n\rRows saved to file " + filePath + " : " + cnt.ToString());
            }
              catch (Exception fail)
            {
                String error = "The following error has occurred:\n\n";
                error += fail.Message.ToString() + "\n\n";
                MessageBox.Show(error);
                Globals.WriteErrorLog(fail.ToString());
                this.Close();
                rtn = -1;
            }

            return rtn;
        }

        private DialogResult ShowInputDialog(ref string input)
        {
            System.Drawing.Size size = new System.Drawing.Size(200, 70);
            Form inputBox = new Form();

            int formWidth = this.ClientSize.Width; 

            inputBox.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            inputBox.ClientSize = size;
            inputBox.Text = "Export Date";

            System.Windows.Forms.TextBox textBox = new TextBox();
            textBox.Size = new System.Drawing.Size(size.Width - 10, 23);
            textBox.Location = new System.Drawing.Point(5, 5);
            textBox.Text = input;
            inputBox.Controls.Add(textBox);

            Button okButton = new Button();
            okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
            okButton.Name = "okButton";
            okButton.Size = new System.Drawing.Size(75, 23);
            okButton.Text = "&OK";
            okButton.Location = new System.Drawing.Point(size.Width - 80 - 80, 39);
            inputBox.Controls.Add(okButton);

            Button cancelButton = new Button();
            cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            cancelButton.Name = "cancelButton";
            cancelButton.Size = new System.Drawing.Size(75, 23);
            cancelButton.Text = "&Cancel";
            cancelButton.Location = new System.Drawing.Point(size.Width - 80, 39);
            inputBox.Controls.Add(cancelButton);

            inputBox.AcceptButton = okButton;
            inputBox.CancelButton = cancelButton;

            inputBox.StartPosition = FormStartPosition.CenterParent;  


            DialogResult result = inputBox.ShowDialog();
            input = textBox.Text;
            return result;
        }

        public DateTime PreviousWorkDay(DateTime date, string connStr)
        {
            do
            {
                date = date.AddDays(-1);
            }
            while (IsHoliday(date, connStr) || IsWeekend(date));

            return date;
        }

        private bool IsWeekend(DateTime date)
        {
            bool rtn = date.DayOfWeek == DayOfWeek.Saturday ||
                   date.DayOfWeek == DayOfWeek.Sunday;
            return rtn;
        }

        private bool IsHoliday(DateTime date, string connStr)
        {
            bool rtn = false;
            try
            {
                SqlConnection sqlConnection1 = new SqlConnection(connStr);
                SqlCommand cmd = new SqlCommand();
                SqlDataReader reader;

                cmd.CommandText = "SELECT 1 FROM MoxyHoliday";
                cmd.CommandText += " WHERE CalendarId = 1 AND Holiday = '" + date.ToString()  + "'";
                cmd.CommandType = CommandType.Text;
                cmd.Connection = sqlConnection1;

                sqlConnection1.Open();

                reader = cmd.ExecuteReader();
                // Data is accessible through the DataReader object here.
                while (reader.Read())
                {
                    if ((int) reader[0] == 1)
                    {
                        rtn = true;
                    }
                }

                reader.Close(); 
                sqlConnection1.Close();
            }
            catch (Exception fail)
            {
                String error = "The following error has occurred:\n\n";
                error += fail.Message.ToString() + "\n\n";
                MessageBox.Show(error);
                Globals.WriteErrorLog(fail.ToString());
                this.Close();
            }

            return rtn;

        }

        private void buttonMoxyToAIM_Click(object sender, EventArgs e)
        {
            //SQLiteDatabase db;
            string outFolder = string.Empty;
            string axysPath = string.Empty;
            string srcFile = string.Empty;                       // the folder where Moxy Export saves moxyaxys.trn file 
            string fileName = string.Empty;                      // source file name only, No path  
            string dbConn = string.Empty;                        // Moxy database connection
            string dbConnPortia = string.Empty;                  // Portia database connection
            string dbConnBilling = string.Empty;                 // BillingWarehouseProtrak database connection
            string tradingCurrencyStoredProc = string.Empty;     // stored procedure to retrieve protfolio's trading currency
            string reportingCurrencyStoredProc = string.Empty;   // stored procedure to retreive portfolio's reporting currency
            string lastCrossRateStoredProc = string.Empty;
            string postToAxys = string.Empty;                    // Y or N; indicates if it's necessary to post trades to Axys
            //int rtn = 0;
            string tradeCur = string.Empty;                      // portfolio trading currency as defined in Protrak
            string crossRate = string.Empty;                     // trades cross rate for non-us based portfolios
            int count = 0;
            string newLine = string.Empty;
            string securityCur = string.Empty;                   // the currency of the traded security

            Globals.errCnt = 0;                                                         // reset error counter

            tbScreen.Clear();
                      
            if (getAppMetaData(ref outFolder, ref axysPath, ref srcFile, ref dbConn, ref postToAxys, ref tradingCurrencyStoredProc, ref dbConnPortia, ref lastCrossRateStoredProc, ref reportingCurrencyStoredProc) == -1) { return; }

            //
            // check if the source file exists moxyaxys.trn
            //
            if (!File.Exists(srcFile)) { MessageBox.Show(String.Format("File {0} not found please run Moxy Export", srcFile)); return; }
            //
            // make a copy of binary source file
            //
            String newSrcFile = Path.GetFileNameWithoutExtension(srcFile) + "_" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("hhmmss") + ".trn";
            File.Copy(srcFile, outFolder + newSrcFile, true);

            // 
            // execute imex to export source file
            //
            runImexExport(outFolder, axysPath, srcFile);
            fileName = Path.GetFileName(srcFile);

            //
            // rename the file to make it CSV and dated
            //              
            //
            // after the test remove time stamp
            //
            String newFile = Path.GetFileNameWithoutExtension(outFolder + fileName) + "_" + DateTime.Now.ToString("yyyyMMdd") + "_"+ DateTime.Now.ToString("hhmmss") + ".csv";

            if (File.Exists(outFolder + fileName))
            {

                try
                {
                    File.Copy(outFolder + fileName, outFolder + newFile, true); // this is an initial csv file that we update down the code to fit AIM specs
                    tbScreen.AppendText("Finished export of " + srcFile + Environment.NewLine);
                    tbScreen.AppendText("Check for the output in: " + outFolder + Environment.NewLine);
                    tbScreen.AppendText("File: " + newFile);


                    MoxyDatabase md = new MoxyDatabase(dbConn, rtbScreen);
                    PortiaDatabase pd = new PortiaDatabase(dbConnPortia, rtbScreen, tradingCurrencyStoredProc, lastCrossRateStoredProc, "");

                    if (unfoldSpecificLots(outFolder + newFile) == -1)
                    {
                        return;
                    }

                    string[] lines = File.ReadAllLines(outFolder + newFile);
                    List<string> newLines = new List<string>();

                    // process Moxy trades file line by line
                    foreach (string line in lines)
                    {

                        Application.DoEvents();

                        // this is a comment line - ignore
                        if (line.IndexOf(";,;,") != -1 || line.IndexOf(",;,") != -1 || line.IndexOf(";;") !=-1) { continue; }
                                             
                        count += 1;
                        lblStatus.Text = String.Format("Processing trade: {0}", count);
                        ///////////////////////////////////////////////////////////////////////////////////////////////
                        //                                                                                           //
                        // analyze & format trades for AIM                                                           //    
                        //                                                                                           //
                        ///////////////////////////////////////////////////////////////////////////////////////////////
                        Trade trade = new Trade(rtbScreen, line, dbConn , dbConnPortia, tradingCurrencyStoredProc , lastCrossRateStoredProc, reportingCurrencyStoredProc, ""  );
                        trade.convert();
                        
                        newLine = String.Join(",", trade.items);
                        if (!trade.doNotInclude)
                        {
                            newLines.Add(newLine);
                        }
                        else
                        {

                            tbScreen.AppendText(String.Format("\r\n-!-!-!-> Excluded cash trade - port & sec currency are the same. : {0}", newLine));
                        }
                        

                    } // end of foreach

                    File.WriteAllLines(outFolder + newFile, newLines);
                    //
                    // when count > 0 the output file is not empty
                    //      1. copy source file (moxyaxys.trn) to Axys Folder where post32.exe can see it
                    //      2. run post32.exe 
                    //

                    // && postToAxys.ToUpper().Equals ('Y')

                    if (count > 0 && postToAxys.ToUpper().Equals('Y'.ToString()))
                    {
                        File.Copy(outFolder + fileName, axysPath + fileName, true);
                        ProcessStartInfo PostProc = new ProcessStartInfo(axysPath + "post32.exe");
                        //PostProc.WorkingDirectory = Path.GetDirectoryName(srcFile);
                       
                        Process p2;
                        PostProc.Arguments = " -fmoxyaxys";
                        PostProc.UseShellExecute = false;
                        p2 = Process.Start(PostProc);
                        while (p2.HasExited == false)
                        {
                            Application.DoEvents();
                        }

                    }


                    // delete moxyaxys.trn
                    File.Delete(srcFile);  

                    
                }
                catch (Exception exp)
                {
                    Globals.saveErr(exp.Message);
                    Globals.WriteErrorLog(exp.ToString());
                    MessageBox.Show(exp.Message);
                    this.Close();
                    return;
                }

            }
            else
            {
                Globals.saveErr("--->Failed to create file: " + newFile);
                return;
            }

            tbScreen.AppendText(String.Format("\r\n\r\nNumber of trades: {0}", count));
            tbScreen.AppendText(String.Format("\r\n\r\nNumber of errors: {0}", Globals.errCnt.ToString()));
            lblStatus.Text = "Ready";
            Globals.init(); 


            //---------------------------TEST----------------//
            //try
            //{
            //    Trade t = new Trade(tbScreen);
            //    t.convert(); 
            //}
            //catch (Exception ex)
            //{
            //    tbScreen.AppendText(ex.Message); 
            //}


            

        }

        //
        // use imex to convert trn file created by Moxy Export procedure in Moxy
        // to csv file
        //
        private void btnPortiaFilesFromImex_Click(object sender, EventArgs e)
        {
           
            //SQLiteDatabase db;
            string outFolder = string.Empty;
            string axysPath = string.Empty; 
            string srcFile = string.Empty;                                       // the folder where Moxy Export saves moxyaxys.trn file 
            string fileName = string.Empty;                                   // source file name only, No path  
            string dbConn = string.Empty;                                      // Moxy database connection
            string dbConnPortia = string.Empty;                            // Portia database connection 
            string tradingCurrencyStoredProc = string.Empty;       // stored procedure to retrieve protfolio's trading currency
            string lastCrossRateStoredProc = string.Empty;          // stored procedure to get last avialable cross rate between two currencies
            string postToAxys = string.Empty;                               // Y or N; indicates if it's necessary to post trades to Axys
            int rtn = 0;
            string tradeCur = string.Empty;                                     // portfolio trading currency as defined in Protrak
            string crossRate = string.Empty;                                    // trades cross rate for non-us based portfolios
            string securityCur = string.Empty;                                  // security currency
            string conversionInstruction = string.Empty;
            string reportingCurrencyStoredProc = string.Empty;
            try
            {
                tbScreen.Clear();
                if (getAppMetaData(ref outFolder, ref axysPath, ref srcFile, ref dbConn, ref postToAxys, ref tradingCurrencyStoredProc, ref dbConnPortia, ref lastCrossRateStoredProc, ref reportingCurrencyStoredProc) == -1)
                {
                    return;
                }
                  
                                
                //
                // check if the source file exists moxyaxys.trn
                //
                if (!File.Exists(srcFile)) 
                { 
                       MessageBox.Show(String.Format ("File {0} not found please rum Moxy Export", srcFile) );
                       return ;
                 }
                //
                // make a copy of binary source file
                //
                String newSrcFile = Path.GetFileNameWithoutExtension(srcFile) + "_" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("hhmmss") + ".trn";
                File.Copy(srcFile, outFolder + newSrcFile, true);
           

                // 
                // execute imex to export source file
                //
                runImexExport(outFolder, axysPath, srcFile);

                fileName = Path.GetFileName(srcFile);
        
                //
                // rename the file to make it CSV and dated
                //              
                 //
                // after the test remove time stamp
                //
                String newFile = Path.GetFileNameWithoutExtension(outFolder + fileName) + "_" + DateTime.Now.ToString("yyyyMMdd") + "_" +  DateTime.Now.ToString("hhmmss")  + ".csv";
              
                if (File.Exists(outFolder + fileName))
                {
                    try
                    {
                        
                        File.Copy(outFolder + fileName, outFolder + newFile, true ); // this is an initial csv file that we update down the code to fit AIM specs
                        tbScreen.AppendText( "Finished export of " + srcFile + Environment.NewLine);
                        tbScreen.AppendText( "Check for the output in: " + outFolder + Environment.NewLine);
                        tbScreen.AppendText( "File: " + newFile);

                                           
                        MoxyDatabase md = new MoxyDatabase(dbConn, rtbScreen );
                        string[] lines = File.ReadAllLines(outFolder + newFile);
                        List<string> newLines = new List<string>() ;
                        int count = 0;
                        string newLine;

                        // process Moxy file line by line
                        foreach (string line in lines)
                        {
                            Application.DoEvents();
                            
                            if (line.IndexOf(";,;,") != -1)
                            {
                                // this is a comment line - ignore
                                continue;
                            }
                           
                            string[] items = line.Split(',');

                            if (items[0].Equals("24905"))
                            {
                                //rtn = rtn;
                            }

                            //
                            // for cash transaction:
                           //                      1. convert spot sells to buys
                            //                  
                            if (items[4].Equals("$cash") && items[1].Equals("sl") )
                            {
                                //
                                // This is Spot Sell:
                                //              1. swap currencies
                                //              2. replace sl with by
                                //              3. flip FX rate 
                                //              4. flip qty and total amount
                                //
                                string tmp = items[3] ;                                                               
                                items[3] = items[11] ;
                                items[11] = tmp;
                                items[1] = "by";
                                double newFxrate = 1/Convert.ToDouble( (items[13]) );
                                items[13] =( newFxrate ).ToString ("0.########")  ;

                                tmp = items[17];
                                items[17] = items[8];
                                items[8] = tmp;

                            }

                            // Apply sell rules for equity sell
                            if (!items[4].Equals("$cash") && (items[1].Equals("sl") || items[1].Equals("SL")))
                            {
                                items[9] = getSellingRule(items[9], items[0]);
                             } // end of selling rule

                            //   1. convert source symbol to Portia format like -CAD CASH-
                            if (items[4].Equals("$cash"))
                            {
                                string newSymbol = string.Empty;
                                rtn = md.convertSymbolToPortiaCash(ref items[3], ref newSymbol);
                                items[4] = newSymbol;
                            }

                            // convert destination symbol to Portia format
                            if (items[12].Equals("$cash"))
                            {
                                string newSymbol = string.Empty;
                                rtn = md.convertSymbolToPortiaCash(ref items[11], ref newSymbol);
                                items[12] = newSymbol;
                            }
                           
                            rtn = md.getTradingCurrency(tradingCurrencyStoredProc, items[0], ref tradeCur);
                            rtn = md.getCrossRate(int.Parse(items[39]), items[5], tradeCur , items[3], items[0], ref crossRate, ref conversionInstruction);
                            //crossRate = null;
                            if (!tradeCur.Equals("USD") && string.IsNullOrEmpty(crossRate))
                            {
                                double num;
                                crossRate = items[13];   // when there's no cross rate for non us based portfolio -> use trade date fx rate
                                if (Double.TryParse(crossRate, out num))
                                {
                                    crossRate = items[13];
                               }
                            } // end of if


                            if (!String.IsNullOrEmpty(crossRate))
                            {
                                //
                                // insert cross rate into the file
                                //
                                Double crossRateNum;
                                items[36] = crossRate;

                                if (Double.TryParse(crossRate, out crossRateNum))
                                {
                                    // get portfolio trading currency
                                    //string tradeCur= string.Empty;
                                    //rtn = md.getTradingCurrency(tradingCurrencyStoredProc, items[0], ref tradeCur);
                                    if (!String.IsNullOrEmpty(tradeCur))
                                    {
                                        //
                                        //  1. replace sorce or destination type & symbol if USD with trading currency
                                        //  2. recalculate quantity or trading amount by applying the cross rate
                                        //

                                        //  source symbol test for USD
                                        if (items[4].IndexOf ("-USD CASH-") != -1 && items[11].Substring(0,2).Equals("ca") ) {
                                            string tmp = items[4].Replace("USD", tradeCur);
                                            double number;
                                            items[4] = tmp;
                                            // recalc quantity with cross rate
                                            if (Double.TryParse(items[8], out number))
                                            {
                                                //items[8] = (number *crossRateNum).ToString("0.##");
                                                //items[17] = number.ToString("0.##");
                                            }

                                          
                                          
                                        } // end of if : source symbol test
                                        // destination symbol test for USD
                                        if (items[12].IndexOf("-USD CASH-") != -1 && items[3].Substring(0,2).Equals("ca"))
                                        {
                                            string tmp = items[12].Replace("USD", tradeCur);
                                            double number;
                                            items[12] = tmp;
                                            // recalc trade amount
                                            if (Double.TryParse(items[8], out number))
                                            {

                                                //items[17] = (number * crossRateNum).ToString("0.##");
                                                //items[8] = (number / crossRateNum).ToString("0.##");
                                            }

                                          

                                        } // end of if : destination symbol test 
                                    }
                                }

                                else
                                    Console.WriteLine("{0} is not a valid cross rate.", crossRate);

                                newLine = String.Join(",", items);
                            }
                            else
                            {
                                newLine = String.Join(",", items);
                            }
                            newLines.Add(newLine); 
                            count++;
                        } // end of For Each Loop

                        File.WriteAllLines(outFolder + newFile, newLines );
                        
                        //
                        // when count > 0 the output file is not empty
                        //      1. copy source file (moxyaxys.trn) to Axys Folder where post32.exe can see it
                        //      2. run post32.exe 
                        //

                        // && postToAxys.ToUpper().Equals ('Y')

                        if (count > 0 && postToAxys.ToUpper().Equals('Y'.ToString() ))
                        {
                            File.Copy(outFolder + fileName, axysPath  + fileName, true);
                            ProcessStartInfo PostProc = new ProcessStartInfo(axysPath + "post32.exe");
                            PostProc.WorkingDirectory = Path.GetDirectoryName(srcFile);
                            Process p2;
                            PostProc.Arguments = " -fmoxyaxys";
                            PostProc.UseShellExecute = false;                           
                            p2 = Process.Start(PostProc);
                            while (p2.HasExited == false)
                            {
                                Application.DoEvents();
                            }

                        }


                        // delete moxyaxys.trn
                        //File.Delete(srcFile);  

                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show(exp.Message);
                        Globals.WriteErrorLog(exp.ToString());
                        this.Close();
                        return;
                    }
                }
                else
                {
                    tbScreen.AppendText( "--->Failed to create file: " + newFile);
                    return;
                }

                //
                // delete Moxy source file
                //
               
                    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                Globals.WriteErrorLog(ex.ToString());
                this.Close();
            }

        }

        //
        // Function runImexExport:  convert Axys or Moxy .trn files to .csv
        //
        private static void runImexExport(string outFolder, string axysPath, string srcFile)
        {
            ProcessStartInfo ImexProc = new ProcessStartInfo(axysPath + "imex32.exe");
            Process p;
            ImexProc.Arguments = " -e -f" + srcFile + " -tcsv -u -c -d" + outFolder;
            p = Process.Start(ImexProc);
            while (p.HasExited == false)
            {
                Application.DoEvents();
            }
        }

        private static string getSellingRule(string switchCase, string portfolio)
        {
            string rtn = string.Empty;

//            if (String.IsNullOrEmpty(switchCase))
//                rtn = string.Empty;


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
                case "a":                   // allocate
                    rtn = "9";
                    break;
                case "l":                   // LIFO
                    rtn = "2";
                    break;
                default:
                    // specific lot
                    rtn = "0";
                    break;
            }
            return rtn;
        }

        private int getAppMetadata(ref string evareFileID, ref string evareFileLocation, ref string reconConnStr) {
            try
            {
                SQLiteDatabase db;
                db = new SQLiteDatabase(sqlitedb);
                DataTable evareConn;

                // get evare file id 
                String query = "select ID \"id\", VALUE \"value\"";
                query += " from recon where id=\"evareFileID\";";
                evareConn = db.GetDataTable(query);
                foreach (DataRow r in evareConn.Rows)
                {
                    evareFileID = r["value"].ToString();

                }

                if (ReportMetaData("recon", "evareFileID", evareFileID, GetCurrentMethod()) == -1) { return -1; }

                // get evare file location 
                query = "select ID \"id\", VALUE \"value\"";
                query += " from recon where id=\"evareFileLocation\";";
                evareConn = db.GetDataTable(query);
                foreach (DataRow r in evareConn.Rows)
                {
                    evareFileLocation = r["value"].ToString();

                }

                if (ReportMetaData("recon", "evareFileLocation", evareFileLocation, GetCurrentMethod()) == -1) { return -1; }

                // get recon connection  
                query = "select ID \"id\", VALUE \"value\"";
                query += " from recon where id=\"connStr\";";
                evareConn = db.GetDataTable(query);
                foreach (DataRow r in evareConn.Rows)
                {
                    reconConnStr = r["value"].ToString();

                }

                if (ReportMetaData("recon", "connStrr", reconConnStr, GetCurrentMethod()) == -1) { return -1; }



                return 0;
            }
            catch (Exception ex)
            {
                Globals.errCnt += 1;
                tbScreen.AppendText("\r\n" + GetCurrentMethod() + ": " + ex.Message);
                Globals.WriteErrorLog(ex.ToString());
                return -1;
            }

        }

        private MetaData getAppMetaDataMoxy()
        {
            try
            {
                string moxycon = Util.getAppConfigVal("moxy24constr");
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
                    mData = new MetaData(reader[0].ToString()   , reader[1].ToString() , reader[2].ToString(),
                                                            reader[3].ToString(), reader[4].ToString(), reader[5].ToString(),
                                                            reader[6].ToString(), moxycon, reader[7].ToString(), reader[8].ToString() ,
                                                            reader[9].ToString());

                }

                reader.Close();
                sqlConnection1.Close();

                return mData;
            }
            catch(Exception ex)
            {
                Globals.errCnt += 1;
                tbScreen.AppendText("\r\n" + GetCurrentMethod() + ": " + ex.Message);
                Globals.WriteErrorLog(ex.ToString());
                throw new Exception("\r\n" + GetCurrentMethod() + ": " + ex.Message);             
            }
        }// eof

        private int getAppMetaData(ref string outFolder, ref string axysPath, ref string srcFile, ref string dbConn, ref string postToAxys, ref string tradingCurrencyStoredProc, ref string dbConnPortia, ref string lastCrossRateStoredProc, ref string reportingCurrencyStoredProc)
        {          
                    
            try
            { 
                // SQLLITE
                SQLiteDatabase db;
                db = new SQLiteDatabase(sqlitedb);
                DataTable moxyConn;

                // get Moxy output folder
                String query = "select ID \"id\", VALUE \"value\"";
                query += " from moxy where id=\"outFolder\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    outFolder = r["value"].ToString();
                  
                }

                if (ReportMetaData("moxy", "outFolder", outFolder, GetCurrentMethod()) == -1) { return -1; }          

                // get Axys location
                query = "select ID \"id\", VALUE \"value\"";
                query += " from moxy where id=\"axyspath\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    axysPath = r["value"].ToString();
                 }

                if (ReportMetaData("moxy", "axyspath", axysPath, GetCurrentMethod()) == -1) { return -1; }

                           
                //
                // get Moxy connection
                //
                query = "select ID \"id\", VALUE \"value\"";
                query += " from moxy where id=\"connstr\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    dbConn = r["value"].ToString();
                   
                }

                if (ReportMetaData("moxy", "connstr", dbConn, GetCurrentMethod()) == -1) { return -1; }

             
                //
                // get Portia connection
                //
                query = "select ID \"id\", VALUE \"value\"";
                query += " from portia where id=\"connStr\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    dbConnPortia = r["value"].ToString();
                    
                }

                if (ReportMetaData("portia", "connStr", dbConnPortia , GetCurrentMethod()) == -1) { return -1; }
                            
                //
                // get trn file source folder
                //
              
                query = "select ID \"id\", VALUE \"value\"";
                query += " from moxy where id=\"srcFile\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    srcFile = r["value"].ToString();                  
                }

                if (ReportMetaData("moxy", "srcFile", srcFile, GetCurrentMethod()) == -1) { return -1; }

                //
                // get postToAxys flag
                //
                query = "select ID \"id\", VALUE \"value\"";
                query += " from moxy where id=\"postToAxys\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    postToAxys = r["value"].ToString();
                   
                }

                if (ReportMetaData("moxy", "postToAxys", postToAxys, GetCurrentMethod()) == -1) { return -1; }
                
                //
                // get trading currency stored proc
                //
                query = "select ID \"id\", VALUE \"value\"";
                query += " from portia where id=\"tradingCurrencyStoredProc\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    tradingCurrencyStoredProc = r["value"].ToString();
                    
                }

                if (ReportMetaData("portia", "tradingCurrencyStoredProc", tradingCurrencyStoredProc, GetCurrentMethod()) == -1) { return -1; }

                //
                // get reporting currency stored proc
                //
                query = "select ID \"id\", VALUE \"value\"";
                query += " from portia where id=\"reportingCurrencyStoredProc\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    reportingCurrencyStoredProc = r["value"].ToString();
                }

                if (ReportMetaData("portia", "reportingCurrencyStoredProc", reportingCurrencyStoredProc, GetCurrentMethod()) == -1) { return -1; }
                
                //
                // get last cross rate stored proc
                //
                query = "select ID \"id\", VALUE \"value\"";
                query += " from portia where id=\"lastCrossRateStoredProc\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    lastCrossRateStoredProc = r["value"].ToString();

                }

                if (ReportMetaData("portia", "lastCrossRateStoredProc", lastCrossRateStoredProc, GetCurrentMethod()) == -1) { return -1; }


                return 0;
            }
            catch (Exception ex)
            {
                Globals.errCnt += 1;
                tbScreen.AppendText( "\r\n" +GetCurrentMethod() +": " + ex.Message);
                Globals.WriteErrorLog(ex.ToString());
                return -1;
            }

           
        } // end of getAppMetaData()


        /// <summary>
        ///     ReportMetaData function: displays on the screen the result of querying
        ///                                                   application meta data from SqLite database
        /// </summary>
        /// <param name="tableName">the name of the table contaning meatdata.</param>
        /// <param name="id">id column in the table</param>
        /// <param name="value">value column the table</param>
        /// <returns>0/-1</returns>
        private int ReportMetaData(string tableName, string id, string metaResult, string callingFunction)
        {
            int rtn = 0;
            string errMsg;
            if (String.IsNullOrEmpty(metaResult) || String.IsNullOrWhiteSpace(metaResult))
            {
                errMsg = String.Format("Function {0} : Meta data has not been retreived from table: {1}, id column parm: {2}. Check SqLite database.", callingFunction, tableName, id );
                tbScreen.AppendText( Globals.saveErr(errMsg) + Environment.NewLine);
                rtn = -1;
            }
            else
            {
                tbScreen.AppendText( String.Format("Application parm for : {0} is {1}", id, metaResult ) + Environment.NewLine);
            }
            return rtn ;
        }


        //
        // Convert FX Connect Trades to AIM format
        //
        private void btn_FXConnectTrades_Click(object sender, EventArgs e)
        {

            string outFolder = string.Empty;
            string axysPath = string.Empty;
            string dbConn = string.Empty;                                      // Moxy database connection
            string dbConnPortia = string.Empty;                            // Portia database connection 
            string tradingCurrencyStoredProc = string.Empty;
            string reportingCurrencyStoredProc = string.Empty;
            string lastCrossRateStoredProc = string.Empty; 
            string aimLocation = null;
            string postToAxys = null;



            string fNameUSD = @"\\tweedy_files\Advent\topostus.trn";
            string fNameAUD = @"\\tweedy_files\Advent\topostau.trn";
            string fNameCAD = @"\\tweedy_files\Advent\topostca.trn";
            string fNameEUR = @"\\tweedy_files\Advent\toposteu.trn";
            string fNameCHF = @"\\tweedy_files\Advent\topostch.trn";
            string fNameNZD = @"\\tweedy_files\Advent\topostnz.trn";
            string fNameGBP = @"\\tweedy_files\Advent\topostgb.trn";

            string fnameAIM = null;     // fx connect file in AIM format

            
            bool usdFlag = false;    // indicates that there are USD trades to be imported
            bool audFlag = false;
            bool cadFlag = false;    // indicates that there are CAD trades to be imported
            bool eurFlag = false;    // indicates that there are EUR trades to be imported
            bool chfFlag = false;    // indicates that there are CHF trades to be imported
            bool nzdFlag = false;    // indicates that there are NZD trades to be imported
             bool gbpFlag = false;

            string fName = null;
            string portfolio = null;
            string portCodeToUse = null;
            string baseCurrency = null;
            string blotterPath = "";
            // defines the user blotter
            string audBlotterpath = null;
            string cadBlotterpath = null;
            string eurBlotterpath = null;
            string chfBlotterpath = null;
            string nzdBlotterpath = null;
            string gbpBlotterpath = null;
            StreamWriter fwUSD = default(StreamWriter);
            StreamWriter fwAUD = default(StreamWriter);
            StreamWriter fwCAD = default(StreamWriter);
            StreamWriter fwEUR = default(StreamWriter);
            StreamWriter fwCHF = default(StreamWriter);
            StreamWriter fwNZD = default(StreamWriter);
            StreamWriter fwGBP = default(StreamWriter);

             StreamWriter fwAIM = default(StreamWriter);

            string srcFile = null;
            
            if (getAppMetaData(ref outFolder, ref axysPath, ref srcFile, ref dbConn, ref postToAxys, ref tradingCurrencyStoredProc, ref dbConnPortia, ref lastCrossRateStoredProc, ref reportingCurrencyStoredProc) == -1) { return; }
            
            //srcFile = "H:\\FXCON\\topost1.trn";
            fName = "H:\\FXCON\\topost.trn";
                      

            //axysPath = "H:\\Axys3\\";

            audBlotterpath = axysPath + "AUDX\\";
            cadBlotterpath = axysPath +"CAD\\";
            eurBlotterpath = axysPath + "EUR\\";
            chfBlotterpath = axysPath + "CHF\\";
            nzdBlotterpath = axysPath + "NZD\\";
            gbpBlotterpath = axysPath + "GBP\\";
                
                tbScreen.Clear();
                SQLiteDatabase db = new SQLiteDatabase(sqlitedb);
                DataTable moxyConn;
               
                //
                // get AIM path
                //
              
                string query = "select ID \"id\", VALUE \"value\"";
                query += " from moxy where id=\"outFolder\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    aimLocation = r["value"].ToString();
                    tbScreen.AppendText( String.Format("AIM Location: {0}", aimLocation) + Environment.NewLine);
                }
                fnameAIM = aimLocation + "FXCon_" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("hhmmss") + ".csv"; ;

                blotterPath = this.getUserBlotter();
                tbScreen.AppendText(String.Format("User Blotter: {0}", blotterPath) + Environment.NewLine);
                       
                //
                // get FX Con source file
                //
                query = "select ID \"id\", VALUE \"value\"";
                query += " from fxconnect where id=\"FXConSrcFile\";";
                moxyConn = db.GetDataTable(query);
                foreach (DataRow r in moxyConn.Rows)
                {
                    srcFile   = r["value"].ToString();
                    tbScreen.AppendText( String.Format("FX Connect source file: {0}", srcFile) + Environment.NewLine);
                }
            
        //DataRow row = default(DataRow);
        DataRow[] rows = null;
        SqlConnection Conn = new SqlConnection(dbConn );
        SqlCommand Cmd = new SqlCommand("usp_GetPortMap", Conn);
        SqlDataAdapter DA = new SqlDataAdapter();
        DataSet DSet = new DataSet();
        Cmd.CommandType = CommandType.StoredProcedure;
        SqlParameter RetValue = Cmd.Parameters.Add("RetValue", SqlDbType.Int);
        RetValue.Direction = ParameterDirection.ReturnValue;
        SqlParameter portcode = Cmd.Parameters.Add("@portcode", SqlDbType.VarChar);
        portcode.Direction = ParameterDirection.Input;
        DA.SelectCommand = Cmd;

    // read through topost.trn and put trades in to
    // appropriate topostXX.trn file based on the base currency
    try {
        String line = null;
        // 
        // get FC trades first - they contain broker info
        //
        FCTrades fctrades = new FCTrades( rtbScreen);
        Stream stream1 = new FileStream(srcFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        StreamReader sr1 = new StreamReader(stream1);

           do {
            line = sr1.ReadLine();
            if (String.IsNullOrEmpty(line))
                continue;
            if (line.IndexOf("cash") == -1)
            {
                fctrades.addTrade(line);
            }

           } while (!(line == null));


        sr1.Close();


        Stream stream = new FileStream(srcFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        StreamReader sr = new StreamReader(stream);
        //StreamReader sr = new StreamReader(srcFile );
        
        
        
        fwUSD = File.CreateText(fNameUSD);
        fwAUD = File.CreateText(fNameAUD);
        fwCAD = File.CreateText(fNameCAD);
        fwEUR = File.CreateText(fNameEUR);
        
        fwCHF = File.CreateText(fNameCHF);
        fwNZD = File.CreateText(fNameNZD);
        fwGBP = File.CreateText(fNameGBP);

        fwAIM =  File.CreateText(fnameAIM);

        do {
            line = sr.ReadLine();
            if (String.IsNullOrEmpty(line))
                continue; 

            // first 5 charecters in each line are portfolio codes
            //portfolio = VisualBasic.Left(line, 5);
            portfolio = line.Substring(0, 5);
            portcode.Value = portfolio;
            Conn.Open();
            DA.Fill(DSet, "maps");

            portCodeToUse = portfolio;
            baseCurrency = "us";

            DataTable t = DSet.Tables["maps"];

            rows = t.Select();
            foreach (DataRow row1 in rows) {
                portCodeToUse =convertToPortiaCode ( row1["PortCodeToUse"].ToString());
                baseCurrency = row1["BaseCurrency"].ToString().Trim() ;
            }

            // write a trade to appropriate blotter file based on
            // portfolio's based currency
            switch (baseCurrency) {

                case "au":
                    // replace port code with port code to use from the map table
                    line = portCodeToUse + line.Remove(0, 5);
                    fwAUD.WriteLine(line);
                    cadFlag = true;
                    break;

                case "ca":
                    // replace port code with port code to use from the map table
                    line = portCodeToUse + line.Remove(0, 5);
                    fwCAD.WriteLine(line);
                    cadFlag = true;
                    break;
                case "eu":
                    line = portCodeToUse + line.Remove(0, 5);
                    fwEUR.WriteLine(line);
                    eurFlag = true;
                    break;
                case "ch":
                    line = portCodeToUse + line.Remove(0, 5);
                    fwCHF.WriteLine(line);
                    chfFlag = true;
                    break;
                case "nz":
                    line = portCodeToUse + line.Remove(0, 5);
                    fwNZD.WriteLine(line);
                    nzdFlag = true;
                    break;
                case "gb":
                    line = portCodeToUse + line.Remove(0, 5);
                    fwGBP.WriteLine(line);
                    gbpFlag = true;
                    break;
                default:
                    fwUSD.WriteLine(line);
                    usdFlag = true;
                    break;
            } // end of switch


            ///////////////////////////////////////////////////////////////////////////////////////////////
            //                                                                                                                  //
            // analyze & format FX trade for AIM                                       //    
            //                                                                                                                  //
            ///////////////////////////////////////////////////////////////////////////////////////////////
            TradeFX trade = new TradeFX(rtbScreen, line, dbConn, dbConnPortia, tradingCurrencyStoredProc, lastCrossRateStoredProc, string.Empty, fctrades, "");
            if (!trade.doNotInclude)
            {  
                trade.convert();
                string tmp = String.Join(",", trade.items);
                fwAIM.WriteLine( String.Join(",", trade.items));
            }
            

            DSet.Clear();
            Conn.Close();
        } while (!(line == null));
        sr.Close();

        fwUSD.Close();
        fwCAD.Close();
        fwEUR.Close();
        fwCHF.Close();
        fwNZD.Close();
        fwGBP.Close();



        fwAIM.Close();

        tbScreen.Text += "\r\n\r\n Created file for AIM: " +fnameAIM;


    } catch (Exception ex) {
        tbScreen.Text  += "\r\n" + ex.Message;
        Globals.WriteErrorLog(ex.ToString());
    }


    if (!postToAxys.Equals("Y".ToString()))
    {

        return;
    }
    else
    {
        DialogResult dialogResult = MessageBox.Show("Would you like to put FX Connect trades in Axys blotters?", "Request", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (dialogResult == DialogResult.Yes)
        {
            //do something
        }
        else if (dialogResult == DialogResult.No)
        {
            //do something else
            return;
        }
    }

    // import into USD Axys blotter
    ProcessStartInfo ImexProc = new ProcessStartInfo(axysPath + "imex32.exe");
    Process p = default(Process);
    File.Copy(fNameUSD, fName, true);
    ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + blotterPath;

    if (usdFlag == true && postToAxys.ToUpper().Equals('Y'.ToString()) ) {

        try {
            if (this.validateTRNFile(fNameUSD, tbScreen )) {
                
               
                p = Process.Start(ImexProc);
                while (!p.HasExited) {
                    // wait for the process to finish
                    Application.DoEvents();
                }
                tbScreen.AppendText( "\r\nFinished import of " + fNameUSD);
                
            }

        } catch (Exception ex) {
            tbScreen.AppendText( "\r\n" + ex.Message);
            Globals.WriteErrorLog(ex.ToString());
        }
    }

    // import into AUD Axys blotter
    ImexProc.WorkingDirectory = "H:\\Axys3\\USERS\\audxuser";
    File.Copy(fNameAUD, fName, true);
    ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + audBlotterpath;

    if (audFlag == true && postToAxys.ToUpper().Equals('Y'.ToString()))
    {

        try
        {
            if (this.validateTRNFile(fNameCAD, tbScreen))
            {
               
                p = Process.Start(ImexProc);
                while (!p.HasExited)
                {
                    // wait for the process to finish
                    Application.DoEvents();
                }
                tbScreen.AppendText(  "\r\nFinished import of " + fNameCAD);
            }

        }
        catch (Exception ex)
        {
            tbScreen.AppendText( "\r\n" + ex.Message);
            Globals.WriteErrorLog(ex.ToString());
        }
    }

    // import into CAD Axys blotter
    ImexProc.WorkingDirectory = "H:\\Axys3\\USERS\\caduser";
    File.Copy(fNameCAD, fName, true);
    ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + cadBlotterpath;

    if (cadFlag == true && postToAxys.ToUpper().Equals('Y'.ToString()))
    {

        try {
            if (this.validateTRNFile(fNameCAD, tbScreen )) {
                
                p = Process.Start(ImexProc);
                while (!p.HasExited) {
                    // wait for the process to finish
                    Application.DoEvents();
                }
                tbScreen.AppendText( "\r\nFinished import of " + fNameCAD);
            }

        } catch (Exception ex) {
            tbScreen.AppendText( "\r\n" + ex.Message);
            Globals.WriteErrorLog(ex.ToString());
        }
    }

    // import into EUR Axys blotter
    ImexProc.WorkingDirectory = "H:\\Axys3\\USERS\\euruser";
    File.Copy(fNameEUR, fName, true);
    ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + eurBlotterpath;

    if (eurFlag == true && postToAxys.ToUpper().Equals('Y'.ToString()))
    {

        try {
            if (this.validateTRNFile(fNameEUR, tbScreen)) {
                

                p = Process.Start(ImexProc);
                while (!p.HasExited) {
                    // wait for the process to finish
                    Application.DoEvents();
                }
                tbScreen.AppendText( "Finished import of " + fNameEUR);
            }

        } catch (Exception ex) {
            tbScreen.AppendText( "\r\n" + ex.Message);
            Globals.WriteErrorLog(ex.ToString());
        }
    }

    // import into CHF Axys blotter
    ImexProc.WorkingDirectory = "H:\\Axys3\\USERS\\chfuser";
    File.Copy(fNameCHF, fName, true);
    ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + chfBlotterpath;

    if (chfFlag == true && postToAxys.ToUpper().Equals('Y'.ToString()))
    {

        try {
            if (this.validateTRNFile(fNameCHF, tbScreen)) {
                
                p = Process.Start(ImexProc);
                while (!p.HasExited) {
                    // wait for the process to finish
                    Application.DoEvents();
                }
                tbScreen.AppendText(  "\r\nFinished import of " + fNameCHF);
            }

        } catch (Exception ex) {
            tbScreen.AppendText( "\r\n" + ex.Message);
            Globals.WriteErrorLog(ex.ToString());
        }
    }

    // import into NZD Axys blotter
    ImexProc.WorkingDirectory = "H:\\Axys3\\USERS\\NZDUSER";
    File.Copy(fNameNZD, fName, true);
    ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + nzdBlotterpath;

    if (nzdFlag == true && postToAxys.ToUpper().Equals('Y'.ToString()))
    {
        try {
            if (this.validateTRNFile(fNameNZD, tbScreen)) {
                
                p = Process.Start(ImexProc);
                while (!p.HasExited) {
                    // wait for the process to finish
                    Application.DoEvents();
                }
                tbScreen.AppendText( "\r\nFinished import of " + fNameNZD);
            }


        } catch (Exception ex) {
            tbScreen.AppendText( "\r\n" + ex.Message);
            Globals.WriteErrorLog(ex.ToString());
        }
    }

    // import into GBP Axys blotter
    ImexProc.WorkingDirectory = "H:\\Axys3\\USERS\\GBPUSER";
    File.Copy(fNameGBP, fName, true);
    ImexProc.Arguments = " -i -f" + fName + " -tcsv -u -c -d" + gbpBlotterpath;

    if (gbpFlag == true && postToAxys.ToUpper().Equals('Y'.ToString()))
    {
        try {
            if (this.validateTRNFile(fNameGBP, tbScreen)) {
                
                p = Process.Start(ImexProc);
                while (!p.HasExited) {
                    // wait for the process to finish
                    Application.DoEvents();
                }
                tbScreen.AppendText( "\r\nFinished import of " + fNameGBP);
            }

        } catch (Exception ex) {
            tbScreen.AppendText( "\r\n" + ex.Message);
            Globals.WriteErrorLog(ex.ToString());
        }
    }

        }


        //
        //  validateTRNFile function
        // 
        private  bool validateTRNFile(string fName, TextBox screen)
        {
            // validates TRN files before import
            bool rtn = true;
            //System.DateTime lastWrite = default(System.DateTime);
            
            try
            {
                if (!File.Exists(fName))
                {
                    screen.AppendText( "\r\nFile " + fName + " not found");
                    return false;
                }

                //lastWrite = File.GetLastWriteTime(fName).Date
                //If CDate(lastWrite) <> Today Then
                //screen.AppendText( vbCrLf + "File " + fName + " is not current"
                //Return False
                //End If

                FileInfo fi = new FileInfo(fName);

                if (fi.Length == 0)
                {
                    screen.AppendText( "\r\nFile " + fName + " is empty");
                    return false;
                }

            }
            catch (Exception ex)
            {
                screen.AppendText( "\r\n" + ex.Message);
                Globals.WriteErrorLog(ex.ToString());
            }
            
            return rtn;
        } // end of validateTRNFile()
        
        //
        // getUserBlotter function - gets an Axys blotter for current Windows user
        //
        private string getUserBlotter()
        {
            string blotterPath = string.Empty ;

            // assign the blotter path based on the windows user
            // blotterPath = "H:\Axys3\USERS\JEANNEPR\" ' default
            System.Security.Principal.WindowsIdentity idWindows = System.Security.Principal.WindowsIdentity.GetCurrent();
            string winUser = idWindows.Name.ToLower();
            switch (winUser)
            {

                case "tweedy\\mikeba":
                    //blotterPath = "H:\\Axys3\\USERS\\MIKEBA\\";
                    blotterPath = @"\\tweedy_files\Advent\Axys3\users\mikeba";
                    break;

                case "tweedy\\jeannepr":
                    blotterPath = @"\\tweedy_files\Advent\Axys3\users\JEANNEPR";
                    break;
                case ("tweedy\\annmariema"):
                    blotterPath = @"\\tweedy_files\Advent\Axys3\users\ANNMARIE";

                    break;
            }

            if (blotterPath.Length == 0)
            {
                tbScreen.AppendText( "\r\n-->-->Function getUserBlotter:  US blotter path undefined... ");
                return string.Empty ;
            }
            else
            {
                tbScreen.AppendText( "\r\nUS blotter: " + blotterPath + "\r\n");
                tbScreen.AppendText( "\r\nWindows User: " + idWindows.Name);
            }


            return blotterPath;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //SQLiteDatabase db;
            string outFolder = string.Empty;
            string axysPath = string.Empty;
            string srcFile = string.Empty;                                       // the folder where Moxy Export saves moxyaxys.trn file 
            string fileName = string.Empty;                                   // source file name only, No path  
            string dbConn = string.Empty;                                      // Moxy database connection
            string dbConnPortia = string.Empty;                            // Portia database connection 
            string tradingCurrencyStoredProc = string.Empty;    // stored procedure to retrieve protfolio's trading currency
            string reportingCurrencyStoredProc = string.Empty;
            string lastCrossRateStoredProc = string.Empty;
            string postToAxys = string.Empty;                               // Y or N; indicates if it's necessary to post trades to Axys
            int rtn = 0;
            string tradeCur = string.Empty;                                     // portfolio trading currency as defined in Protrak
            string crossRate = string.Empty;                                    // trades cross rate for non-us based portfolios
            int count = 0;
            string newLine = string.Empty;
            string securityCur = string.Empty;
            string conversionInstruction = null;
            Globals.errCnt = 0;                                                         // reset error counter

            tbScreen.Clear();
            if (getAppMetaData(ref outFolder, ref axysPath, ref srcFile, ref dbConn, ref postToAxys, ref tradingCurrencyStoredProc, ref dbConnPortia,ref lastCrossRateStoredProc, ref reportingCurrencyStoredProc) == -1) { return; }

            //
            // check if the source file exists moxyaxys.trn
            //
            if (!File.Exists(srcFile)) { MessageBox.Show(String.Format("File {0} not found please rum Moxy Export", srcFile)); return;}
            //
            // make a copy of binary source file
            //
            String newSrcFile = Path.GetFileNameWithoutExtension(srcFile) + "_" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("hhmmss") + ".trn";
            File.Copy(srcFile, outFolder + newSrcFile, true);

            // 
            // execute imex to export source file
            //
            runImexExport(outFolder, axysPath, srcFile);

            fileName = Path.GetFileName(srcFile);

            //
            // rename the file to make it CSV and dated
            //              
            //
            // after the test remove time stamp
            //
            String newFile = Path.GetFileNameWithoutExtension(outFolder + fileName) + "_" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("hhmmss") + ".csv";

            if (File.Exists(outFolder + fileName))
            {

                try
                {
                    File.Copy(outFolder + fileName, outFolder + newFile, true); // this is an initial csv file that we update down the code to fit AIM specs
                    tbScreen.AppendText( "Finished export of " + srcFile + Environment.NewLine);
                    tbScreen.AppendText( "Check for the output in: " + outFolder + Environment.NewLine);
                    tbScreen.AppendText ("File: " + newFile);


                    MoxyDatabase md = new MoxyDatabase(dbConn, rtbScreen);
                    PortiaDatabase pd = new PortiaDatabase(dbConnPortia, rtbScreen, tradingCurrencyStoredProc, lastCrossRateStoredProc, "");
                    string[] lines = File.ReadAllLines(outFolder + newFile);
                    List<string> newLines = new List<string>();                 

                     // process Moxy trades file line by line
                    foreach (string line in lines)
                    {
                      
                        Application.DoEvents();

                        // this is a comment line - ignore
                        if (line.IndexOf(";,;,") != -1 && line.IndexOf(",;,") != -1) { continue; }

                        string[] items = line.Split(',');
                        tbScreen.AppendText("\r\n -- ");
                        count += 1;
                        lblStatus.Text = String.Format("Processing trade: {0}", count);
                        ///////////////////////////////////////////////////////////////////////////////////////////////
                        //                                                                                                                  //
                        // analyze & format trades for AIM                                                          //    
                        //                                                                                                                  //
                        ///////////////////////////////////////////////////////////////////////////////////////////////
                       if (items[0].Equals("24807")) {
                            rtn = 0;
                        }
                        rtn = md.getISOCurrency(items[3], ref securityCur);
                        rtn = pd.getTradingCurrency(tradingCurrencyStoredProc, items[0], ref tradeCur);
                        crossRate = string.Empty; // clear prev cross rate

                        if (!tradeCur.Equals("USD")  && String.IsNullOrEmpty (items[36])  ) {
                            if (md.getCrossRate(int.Parse(items[39]), items[5], tradeCur, items[3], items[0], ref crossRate, ref conversionInstruction) == -1) { }
                               setCrossRate(crossRate, ref  items);
                        }
                     
                       clearSettleFX(ref items);
                        
                        if (items[4].Equals("$cash"))
                        ////////////////
                        // cash trade
                        ///////////////
                        {
                            //rtn = pd.getTradingCurrency(tradingCurrencyStoredProc, items[0], ref tradeCur);
                            //rtn = md.getISOCurrency(items[3], ref securityCur);

                            // do not include cash trades when portfolio currency matches security currency
                            if (tradeCur.Equals(securityCur )) {  continue; }

                            // replace Moxy src & dest symbols with Portia's
                            //   1. convert source symbol to Portia format like -CAD CASH-
                            if (items[4].Equals("$cash"))
                            {
                                //string newSymbol = string.Empty;
                                //rtn = md.convertSymbolToPortiaCash(items[3], ref newSymbol);
                                //items[4] = newSymbol;
                                rtn = md.convertSymbolToPortiaCash(ref items[3], ref items[4] );
                            }

                            // convert destination symbol to Portia format
                            if (items[12].Equals("$cash"))
                            {
                                //string newSymbol = string.Empty;
                                //rtn = md.convertSymbolToPortiaCash(items[11], ref newSymbol);
                                //items[12] = newSymbol;
                                rtn = md.convertSymbolToPortiaCash(ref items[11], ref items[12] );
                            }


                            if (!String.IsNullOrEmpty(crossRate))
                            // a cash trade for NON US based portfolio
                            {
                                if ((items[1].Equals("sl") || items[1].Equals("SL")))
                                    // a spot sell for non us based portfolio
                                    setNonUSCashSell(ref items, tradeCur);
                                else
                                    // a spot buy for non us based portfolio
                                    setNonUSCashBuy(ref items);
                             }
                            else
                                // a cash trade for US based portfolio
                            {
                                if ((items[1].Equals("sl") || items[1].Equals("SL")))
                                      setUSCashSell(ref items);
                                else
                                    setUSCashBuy(ref items);                             
                            }

                        }
                        else
                        ///////////////////
                        // equity trade
                        ///////////////////
                        {

                            if ((items[1].Equals("sl") || items[1].Equals("SL")))
                                setEquitySell(ref items);
                            else
                                setEquityBuy(ref items);
                          
                        }

                        newLine = String.Join(",", items);
                        newLines.Add(newLine);
                       
                    } // end of foreach

                    File.WriteAllLines(outFolder + newFile, newLines);
                    //
                    // when count > 0 the output file is not empty
                    //      1. copy source file (moxyaxys.trn) to Axys Folder where post32.exe can see it
                    //      2. run post32.exe 
                    //

                    // && postToAxys.ToUpper().Equals ('Y')

                    if (count > 0 && postToAxys.ToUpper().Equals('Y'.ToString()))
                    {
                        File.Copy(outFolder + fileName, axysPath + fileName, true);
                        ProcessStartInfo PostProc = new ProcessStartInfo(axysPath + "post32.exe");
                        PostProc.WorkingDirectory = Path.GetDirectoryName(srcFile);
                        Process p2;
                        PostProc.Arguments = " -fmoxyaxys";
                        PostProc.UseShellExecute = false;
                        p2 = Process.Start(PostProc);
                        while (p2.HasExited == false)
                        {
                            Application.DoEvents();
                        }

                    }


                    // delete moxyaxys.trn
                    //File.Delete(srcFile);  
                   
                   

                } 
                catch (Exception exp)
                {
                    Globals.saveErr(exp.Message);
                    Globals.WriteErrorLog(exp.ToString());
                    MessageBox.Show(exp.Message);
                    this.Close();
                    return;
                }

            }
            else
            {
                Globals.saveErr("--->Failed to create file: " + newFile);
                return;
            }
            
            tbScreen.AppendText(String.Format("\r\n\r\nNumber of trades: {0}", count ));
            tbScreen.AppendText( String.Format("\r\n\r\nNumber of errors: {0}", Globals.errCnt.ToString() ));
            lblStatus.Text = "Ready";
            Globals.init(); 

        } // end of Moxy To AIM conversion


        private String convertToPortiaCode(String portCode)
        {
            String rtn = String.Empty;

            rtn = portCode;

            if (!String.IsNullOrEmpty(portCode))
            {
                if(rtn.IndexOf("fc" ) != -1  )  {
                    rtn = rtn.Replace("fc", String.Empty);
                }

                if (rtn.Substring(0, 1).Equals('7'.ToString())  )
                {
                    StringBuilder sb = new StringBuilder(rtn );
                    sb[0] = '2';
                    rtn = sb.ToString();
                }
            }

            return rtn;
        }   
 
        private void setCrossRate(string crossRate, ref string[] items)
        {

            try
            {
                if (!String.IsNullOrEmpty(crossRate) && items.Length > 0)
                {
                    //
                    // insert cross rate 
                    //
                    items[36] = crossRate;
                }
            }
            catch (Exception ex)
            {
                tbScreen.AppendText( Globals.saveErr( "\r\n" +GetCurrentMethod()+ ": " + ex.Message + "\r\n"));
                Globals.WriteErrorLog(ex.ToString());
            }

         
        } // end of setCrossRate()

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
        ///     setEquityBuy() - sets necessary fields used by AIM for equity buy for US based portfolios
        /// </summary>
        /// <param name="items">Moxy Trade represented by string array of items.</param>
        private void setEquityBuy(ref string[] items)
        {
            try
            {
                // sec2base --> settleFX = tradeFX
                items[14] = items[13]; 

                // sec2cbal
                items[15] = "1";

                // sec2port
                items[13] = items[36];
            }
            catch (Exception ex)
            {
                tbScreen.AppendText(Globals.saveErr( "\r\nFunction setEquityBuy Exception:" + ex.Message + "\r\n"));
                Globals.WriteErrorLog(ex.ToString());
            }
        } // end of setEquityBuy()

        /// <summary>
        ///     setEquitySell() - sets necessary fields used by AIM for equity sell for US portfolios
        /// </summary>
        /// <param name="items">Moxy Trade represented by string array of items.</param>
        private void setEquitySell(ref string[] items)
        {
           

            try
            {
                // covert Moxy selling rule to Portia's
                items[9] = getSellingRule(items[9], items[0]);

                // sec2base --> settleFX = tradeFX
                items[14] = items[13];

                // sec2cbal
                items[15] = "1";

                // sec2port
                items[13] = items[36];

             
            }
            catch (Exception ex)
            {
                tbScreen.AppendText(Globals.saveErr( "\r\nFunction setEquitySell Exception:" + ex.Message + "\r\n"));
                Globals.WriteErrorLog(ex.ToString());
            }

        } // end of setEquitySell()

        /// <summary>
        ///     setNonUSCashBuy() - sets necessary fields used by AIM for spot buy for Non US based portfolios
        /// </summary>
        /// <param name="items">Moxy Trade represented by string array of items.</param>
        private void setNonUSCashBuy(ref string[] items)
        {
            try
            {
                // trade amount
                 Double crossRateNum, qtyNum=0, tradeAmt=0;
                 String  crossRate = items[36];
                 String qty = items[8];

                  if (Double.TryParse(crossRate, out crossRateNum))
                  {
                      if (Double.TryParse(qty, out qtyNum)) {
                          tradeAmt = Math.Round (qtyNum / crossRateNum, RNDNUM );
                          items[17] = tradeAmt.ToString();   
                      }
                      else {
                           tbScreen.AppendText(Globals.saveErr( "\r\n Function setNonUSCashBuy: Qty is unavailable for the trade-->  " +String.Join(",", items ) ));
                      }
                  }
                  else
                  {
                      tbScreen.AppendText(Globals.saveErr( "\r\n Function setNonUSCashBuy: Cross Rate is unavailable for the trade-->  " +String.Join(",", items ) ));
                  }


                // sec2Base
                  String byCur = items[3].Substring(2,2);
                  if (byCur == "us")
                  {
                      items[14] = "1";
                  }
                  else
                  {
                      items[14] = String.Empty;
                  }
                
                // Sec2Cbal
                  items[15] = Math.Round ((qtyNum / tradeAmt), RNDNUM ).ToString();              

                // Sec2Port
                  items[13] = crossRate;

            }
            catch (Exception ex)
            {
                tbScreen.AppendText( Globals.saveErr("\r\n Function setNonUSCashBuy: " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }
        } // end of setNonEquityBuy()


        /// <summary>
        ///     setNonUSCashSell() - sets necessary fields used by AIM for equity sell for Non US based portfolios
        /// </summary>
        /// <param name="items">Moxy Trade represented by string array of items.</param>
        private void setNonUSCashSell(ref string[] items, string tradingCur)
        {

            try
            {

                // swap src dest symbols
                string tmp = items[3];
                string tmp2 = items[4];
                items[3] = items[11];
                items[4] = items[12];
                items[11] = tmp;
                items[12] = tmp2;

                // replace sell with buy with trading currency
                items[1] = "by";
                items[4] = String.Format("-{0} CASH-", tradingCur);
                
                // qty & tradeAmt
                Double crossRateNum, qtyNum = 0;
                String crossRate = items[36];
                String qty = items[8];

                // when cross rate is unavailable on the trade try to get it from Portia
                if (String.IsNullOrEmpty(crossRate))
                {
                   
                }

                if (Double.TryParse(crossRate, out crossRateNum))
                {
                    if (Double.TryParse(qty, out qtyNum))
                    {
                        items[8] = (qtyNum /(1/ crossRateNum)).ToString("0.##");
                        items[17] = qtyNum.ToString();          //  trade amount becomes orig qty 
                    }
                    else
                    {
                        tbScreen.AppendText(Globals.saveErr( "\r\n" + GetCurrentMethod() + " : Qty is unavailable for the trade-->  " + String.Join(",", items)));
                    }
                }
                else
                {
                    tbScreen.AppendText(Globals.saveErr( "\r\n" + GetCurrentMethod() + " : Cross Rate is unavailable for the trade-->  " + String.Join(",", items)));
                }

                // sec2Base
                String fxRate = items[13];
                Double fxRateNum = 0;

                if (Double.TryParse(fxRate, out fxRateNum))
                {
                    items[14] =Math.Round (  (crossRateNum * fxRateNum), RNDNUM ).ToString();
                }
                else
                {
                    tbScreen.AppendText(Globals.saveErr( "\r\n" + GetCurrentMethod() + " : FX Rate is unavailable for the trade-->  " + String.Join(",", items)));
                }

                // sec2cbal
                items[15] = Math.Round (  (Double.Parse(items[8]) / Double.Parse(items[17])), RNDNUM ).ToString();

                // sec2port
                items[13] =Math.Round ( (1 / crossRateNum), RNDNUM ).ToString();
            }
            catch (Exception ex)
            {
                tbScreen.AppendText( Globals.saveErr ("\r\n" + GetCurrentMethod()+ " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

        } // end of setNonEquitySell()

        /// <summary>
        ///     setUSCashBuy() - sets necessary fields used by AIM for equity buy for Non US based portfolios
        /// </summary>
        /// <param name="items">Moxy Trade represented by string array of items.</param>
        private void setUSCashBuy(ref string[] items)
        {
            try
            {
                // sec2Base
                items[14] = items[13];
                
                // sec2Cbal
                Double qtyNum=0, tradeAmt=0;
                if (Double.TryParse(items[8], out qtyNum) && Double.TryParse(items[17], out tradeAmt) && !String.IsNullOrEmpty(items[17])  )
                {
                    items[15] =Math.Round ((qtyNum/tradeAmt),RNDNUM).ToString();
                }
                else
                {
                    tbScreen.AppendText( Globals.saveErr("\r\n" + GetCurrentMethod() + " : Qty or Trade Amount is unavailable for the trade-->  " + String.Join(",", items)));
                }


                // sec2Port
                items[14] = items[14];

            }
            catch (Exception ex)
            {
                tbScreen.AppendText( Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

        } // end of setUSCashBuy()


        /// <summary>
        ///     setUSCashSell() - sets necessary fields used by AIM for equity sell for Non US based portfolios
        /// </summary>
        /// <param name="items">Moxy Trade represented by string array of items.</param>
        private void setUSCashSell(ref string[] items)
        {

            try
            {

                // swap src dest symbols
                string tmp = items[3];
                items[3] = items[11];
                items[11] = tmp;

                tmp = items[4];
                items[4] = items[12];
                items[12] = tmp;


                // replace sell with buy
                items[1] = "by";

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
                    tbScreen.AppendText(Globals.saveErr( "\r\n" + GetCurrentMethod() + " : Qty is unavailable for the trade-->  " + String.Join(",", items)));
                }

                // sec2Base
                String fxRate = items[13];
                Double fxRateNum = 0;

                if (Double.TryParse(fxRate, out fxRateNum) && fxRateNum != 0)
                {
                    items[14] = Math.Round ((1/ fxRateNum), RNDNUM).ToString();
                }
                else
                {
                    tbScreen.AppendText( Globals.saveErr("\r\n" + GetCurrentMethod() + " : FX Rate is unavailable for the trade-->  " + String.Join(",", items)));
                }


                // sec2CBal
                Double tradeAmt = 0;
                if (Double.TryParse(items[17], out tradeAmt) && !String.IsNullOrEmpty(items[17]))
                {
                    items[15] =Math.Round (  (qtyNum / tradeAmt), RNDNUM).ToString();
                }
                else
                {
                    tbScreen.AppendText(Globals.saveErr( "\r\n" + GetCurrentMethod() + " : Trade Amount is unavailable for the trade-->  " + String.Join(",", items)));
                }


                // sec2Port
                items[14] = items[14];
            }
            catch (Exception ex)
            {
                tbScreen.AppendText( Globals.saveErr( "\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

        } // end of setUSCashSell()

        public string GetCurrentMethod()
        {
            StackTrace st = new StackTrace();
            StackFrame sf = st.GetFrame(1);

            return sf.GetMethod().Name;
        }

     
        /// <summary>
        ///     unfoldSpecificLots - replaces specific lot sells with trades with lot numbers
        /// </summary>
        /// <param name="file">a file name with Moxy trades </param>
        /// <returns></returns>
        private int unfoldSpecificLots(string file) {
            int rtn = 0;

            try
            {
                string[] lines = File.ReadAllLines(file);
                List<string> newLines = new List<string>();
                MetaData mData = getAppMetaDataMoxy();
                PortiaDatabase pd = new PortiaDatabase(mData.portiaConStr, rtbScreen, mData.tradingCurrencyStoredProc, mData.lastCrossRateProc, mData.sellRuleStoredProc);

                if (lines != null && lines.Length != 0)
                {
                    // process Moxy trades file line by line
                    for (int i=0; i<lines.Length  ; i++)
                    {
                        // this is a comment line - ignore
                        if (lines[i].IndexOf(";,;,") != -1) { continue; }

                        string[] items = lines[i].Split(',');
                        
                        // clear tax lot # destination
                        items[31] = string.Empty;

                        // equity sell
                        if (!items[4].Equals("$cash") && (items[1].Equals("sl") || items[1].Equals("SL")))
                        {

                            //
                            // check if the selling rule has been set by moxy
                            //
                            string sellingRule = getSellingRule(items, pd);

                            //
                            // specific lots sold -> replace this trade with specific lots trades
                            // the lines after this will have 
                            //
                            if (sellingRule.Equals("0"))
                            {
                                // preserve orig sell
                                string[] origSell=(string[]) items.Clone() ;
                                Double otherFeeSum = 0;
                                Double tradeAmtSum = 0;
                                Double secFeeSum = 0;
                                Double commissionSum = 0;
                                Array.Copy(items, origSell, 0 ); 

                                i++;
                                //&& (i + 1) < lines.Length
                                while (i < lines.Length && lines[i].IndexOf("LOT:QTY") != -1 )
                                {

                                    string[] lot = lines[i].Split(',');
                                    string lotNum = string.Empty;
                                    string qty = string.Empty;
                                    if (extractLotQty(lot[2], ref lotNum, ref qty) != -1)
                                    {
                                        //
                                        // create specific lot trade here
                                        // lot number goes to dest. lot location [31] 
                                        //
                                        string[] lotTrade = null;
                                        if (createLotSell(origSell, qty, lotNum, ref lotTrade) != -1)
                                        {
                                            newLines.Add(String.Join(",", lotTrade));
                                            otherFeeSum += Double.Parse(lotTrade[26]);
                                            tradeAmtSum += Double.Parse(lotTrade[17]);
                                            secFeeSum += Double.Parse(lotTrade[22]);
                                            commissionSum += Double.Parse(lotTrade[23]);
                                        }
                                        else
                                        {
                                            // failed to create lot sell
                                            tbScreen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " :  Failed to create a lot sell."));
                                        }// end of else

                                    }// end of if
                                    i++;
                                } // end of while

                                // add rounding difference to the last lot
                                Double roundDiffOtherFee =Math.Round( Double.Parse(origSell[26]) - otherFeeSum, Globals.RNDNUM2);
                                addRoundingDiffToLastLot(26, newLines, roundDiffOtherFee);
                                
                                //if(roundDiffOtherFee != 0)
                                //{
                                //    string[] lastLot = newLines[newLines.Count - 1].Split(',');
                                //    // new other fee
                                //    lastLot[26] = (Double.Parse(lastLot[26]) + roundDiffOtherFee).ToString();
                                //    // remove last new line and add corrected
                                //    newLines.RemoveAt(newLines.Count - 1);
                                //    newLines.Add(String.Join(",", lastLot));
                                //}
                                Double roundDiffTradeAmt = Math.Round(Double.Parse(origSell[17]) - tradeAmtSum, Globals.RNDNUM2);
                                addRoundingDiffToLastLot(17, newLines, roundDiffTradeAmt);
                                
                                //if (roundDiffTradeAmt != 0)
                                //{
                                //    string[] lastLot = newLines[newLines.Count - 1].Split(',');
                                //    // new trade amt
                                //    lastLot[17] = (Double.Parse(lastLot[17]) + roundDiffTradeAmt).ToString();
                                //    // remove last new line and add corrected
                                //    newLines.RemoveAt(newLines.Count - 1);
                                //    newLines.Add(String.Join(",", lastLot));
                                //}
                                Double roundDiffSecFee = Math.Round(Double.Parse(origSell[22]) - secFeeSum, Globals.RNDNUM2);
                                addRoundingDiffToLastLot(22, newLines, roundDiffSecFee);

                                //if (roundDiffSecFee != 0)
                                //{
                                //    string[] lastLot = newLines[newLines.Count - 1].Split(',');
                                //    // new trade amt
                                //    lastLot[22] = (Double.Parse(lastLot[22]) + roundDiffSecFee).ToString();
                                //    // remove last new line and add corrected
                                //    newLines.RemoveAt(newLines.Count - 1);
                                //    newLines.Add(String.Join(",", lastLot));
                                //}
                                Double roundDiffCommission = Math.Round(Double.Parse(origSell[23]) - commissionSum, Globals.RNDNUM2);
                                addRoundingDiffToLastLot(23, newLines, roundDiffCommission);
                                
                                //if (roundDiffCommission != 0)
                                //{
                                //    string[] lastLot = newLines[newLines.Count - 1].Split(',');
                                //    // new trade amt
                                //    lastLot[23] = (Double.Parse(lastLot[23]) + roundDiffCommission).ToString();
                                //    // remove last new line and add corrected
                                //    newLines.RemoveAt(newLines.Count - 1);
                                //    newLines.Add(String.Join(",", lastLot));
                                //}

                                i--;
                            }
                            else
                            {
                                newLines.Add(lines[i]);
                            }

                        }
                        else
                        {
                            newLines.Add(lines[i]);
                        }


                    }// end of for
                    File.WriteAllLines(file, newLines);
                } // end of if
                else
                {
                    // empty file
                    tbScreen.AppendText(file + " is empty."); 
                    rtn = -1;
                }
            } // end of try
            catch (Exception ex)
            {
                tbScreen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }
        
            return rtn;
        } // end of uySpecificLots()

        /// <summary>
        ///     extractLotQty() - extract lot number and lot qty 
        /// </summary>
        /// <param name="lotQtyField">string contaning LOT:QTY key:value</param>
        /// <returns>0/-1</returns>
        private int extractLotQty(string lotQtyField, ref string lotNum, ref string  qty) {
            int rtn = 0;
            try {
                // argument might look like this:  LOT:QTY 1:8600.00000000
                string s = lotQtyField.Replace  ("LOT:QTY ", string.Empty ) ;
                string[] values = s.Split(':');
                lotNum = values[0];
                qty = values[1];

            }
            catch (Exception ex)
            {
                rtn = -1;
                tbScreen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

            return rtn;
        } // end of extractLotQty

        /// <summary>
        ///     using original specific sell trade create a sell trade for a lot
        /// </summary>
        /// <param name="origSel">original sell trade</param>
        /// <param name="lotQty">lot qty</param>
        /// <param name="lotNum">lot number</param>
        /// <param name="lotSell">new lot trade</param>
        /// <returns>0/-1</returns>
        private int createLotSell(string[] origSel, string lotQty, string lotNum, ref string[] lotSell) {
            int rtn =0;
            try
            {
                Double number;
                lotSell =(string[])origSel.Clone() ;
                lotSell[8] = lotQty;          
                lotSell [31] = lotNum;
                Double newTradeAmt =Math.Round (( Double.Parse(origSel[17]) / Double.Parse(origSel[8]))*  Double.Parse(lotQty), Globals.RNDNUM2) ;
                lotSell[17] = newTradeAmt.ToString();
                lotSell[31] = lotNum;

                // sec fee
                if (Double.TryParse(origSel[22] , out number)){
                    Double newSECFee = Math.Round((Double.Parse(origSel[22]) / Double.Parse(origSel[8])) * Double.Parse(lotQty ), Globals.RNDNUM2);
                    lotSell[22] = newSECFee.ToString();  
                }
                // commission
                if (Double.TryParse(origSel[23], out number))
                {
                    Double newCommission = Math.Round((Double.Parse(origSel[23]) / Double.Parse(origSel[8])) * Double.Parse(lotQty ), Globals.RNDNUM2);
                    lotSell[23] = newCommission.ToString();  
                }
                // other fee
                if (Double.TryParse(origSel[26], out number))
                {
                    Double newOtherFee = Math.Round((Double.Parse(origSel[26]) / Double.Parse(origSel[8])) * Double.Parse(lotQty ), Globals.RNDNUM2);
                    lotSell[26] = newOtherFee.ToString();
                }
                
            }
            catch (Exception ex)
            {
                rtn = -1;
                tbScreen.AppendText(Globals.saveErr("\r\n" + GetCurrentMethod() + " : " + ex.Message));
                Globals.WriteErrorLog(ex.ToString());
            }

            return rtn;
        }

    

        /// <summary>
        ///     Convert Evare daily files to AIM
        ///                     1.	Evare files location: N:\Evare\Download
        ///                     2.	File name consists of Bank name, file type + yyyyMMdd.csv
        ///                     3.	Find all the files updated as of specified date with word ‘Tran’ in the file name.
        ///                     4.	Read each file and convert transactions to AIM format.         
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_Evare_Click(object sender, EventArgs e)
        {
            string evareFileLocation = string.Empty;
            string evareFileID = string.Empty;
            string evareConnStr = string.Empty;
            ArrayList files = new ArrayList();

            if ( getAppMetadata(ref evareFileID, ref evareFileLocation , ref evareConnStr ) == -1) { return; }

            EvareFile ef = new EvareFile(tbScreen, evareConnStr, evareFileID, evareFileLocation);
            files = ef.getFiles();
            foreach (String f in files)
            {
                ef.convertToAIM(f);
            }

        }
              
        private void button3_Click(object sender, EventArgs e)
        {
            Globals.errCnt = 0;                                                         // reset error counter
            tbScreen.Clear();

            MetaData mData = getAppMetaDataMoxy();

        }

        // Moxy -> AIM trades
        private void button3_Click_1(object sender, EventArgs e)
        {
          
            Globals.errCnt = 0;                                                         // reset error counter
            tbScreen.Clear();
            string fileName = null;
            int count = 0;

            MetaData mData = getAppMetaDataMoxy();
           
            //
            // check if the source file exists moxyaxys.trn
            //
            if (!File.Exists(mData.srcFile)) { MessageBox.Show(String.Format(fileNotFound , mData.srcFile)); return; }
            //
            // make a copy of binary source file
            //
            String newSrcFile = Path.GetFileNameWithoutExtension(mData.srcFile) + Util.DateTimeStamp() + ".trn";
            File.Copy(mData.srcFile, mData.outFolder + newSrcFile, true);
            // 
            // execute imex to export source file
            //
            runImexExport(mData.outFolder, mData.axysPath, mData.srcFile);
            fileName = Path.GetFileName(mData.srcFile);
            //
            // rename the file to make it CSV and dated
            //              
            //
            // after the test remove time stamp
            //
            String newFile = Path.GetFileNameWithoutExtension(mData.outFolder + fileName) + Util.DateTimeStamp () + ".csv";
            if (!File.Exists(mData.outFolder + fileName)) { Globals.saveErr("--->Failed to create file: " + newFile); return; }

            try
            {
                File.Copy(mData.outFolder + fileName, mData.outFolder + newFile, true); // this is an initial csv file that we update down the code to fit AIM specs
                tbScreen.AppendText("Finished export of " +mData.srcFile + Environment.NewLine);
                tbScreen.AppendText("Check for the output in: " + mData.outFolder + Environment.NewLine);
                tbScreen.AppendText("File: " + newFile);
                
                MoxyDatabase md = new MoxyDatabase(Util.getAppConfigVal("moxyconstr"), rtbScreen);
                PortiaDatabase pd = new PortiaDatabase(mData.portiaConStr , rtbScreen, mData.tradingCurrencyStoredProc, mData.lastCrossRateProc, mData.sellRuleStoredProc);
                if (unfoldSpecificLots(mData.outFolder + newFile) == -1)
                {
                    return;
                }

                string[] lines = File.ReadAllLines(mData.outFolder + newFile);
                List<string> newLines = new List<string>();
                 
                //foreach(string line in lines)
                //{
                   
                //    // this is a comment line - ignore
                //    if (line.IndexOf(";,;,") != -1 || line.IndexOf(",;,") != -1 || line.IndexOf(";;") != -1) { continue; }

                //    Application.DoEvents();
                //    count += 1;
                //    lblStatus.Text = String.Format("Processing trade: {0}", count);
                  
                //    string convertedLine = convertMoxyTrade(line, mData);
                //    if (convertedLine != null)
                //        newLines.Add(convertedLine);
                //}

                for (int i = 0; i < lines.Length; i++) // Loop through List with for
                {
                    // this is a comment line - ignore
                    if (lines[i].IndexOf(";,;,") != -1 || lines[i].IndexOf(",;,") != -1 || lines[i].IndexOf(";;") != -1) { continue; }
                    // if Auto Spot line - ignore
                    if((i+1)< lines.Length && lines[i+1].IndexOf("Auto Spot")!=-1)
                    {
                        continue;
                    }
                    if(lines[i ].IndexOf("Auto Spot") != -1) { continue; }

                    // process line
                    Application.DoEvents();
                    count += 1;
                    lblStatus.Text = String.Format("Processing trade: {0}", count);

                    string convertedLine = convertMoxyTrade(lines[i], mData);
                    if (convertedLine != null)
                        newLines.Add(convertedLine);

                }

                File.WriteAllLines(mData.outFolder + newFile, newLines);

                // send a copy of the file to Portia 12 conversion folder
                String fileCopyOutFolder = Util.getAppConfigVal("moxyAIMOutFolder");
                File.WriteAllLines(fileCopyOutFolder + newFile, newLines);

                // delete trn file
                File.Delete(mData.srcFile);
                // report errors
                tbScreen.AppendText(String.Format("\r\n\r\nFile copy: {0}", fileCopyOutFolder + newFile ));
                tbScreen.AppendText(String.Format("\r\n\r\nNumber of trades: {0}", count));
                tbScreen.AppendText(String.Format("\r\n\r\nNumber of errors: {0}", Globals.errCnt.ToString()));
                lblStatus.Text = "Ready";
                Globals.init();

            }
            catch (Exception exp)
            {
                Globals.saveErr(exp.Message);
                Globals.WriteErrorLog(exp.ToString());
                MessageBox.Show(exp.Message);
                this.Close();
                return;
            }


        }// eof

        /// <summary>
        /// analyze and format trades for AIM
        /// </summary>
        /// <param name="line"></param>
        /// <returns></returns>
        private string convertMoxyTrade(string line, MetaData mData)
        {
            string tradeLine;
            try
            {
                Trade trade = new Trade(rtbScreen, line, mData.moxyConStr , mData.portiaConStr , 
                                                        mData.tradingCurrencyStoredProc, mData.lastCrossRateProc, mData.reportingCurrencyStoredProc, mData.sellRuleStoredProc);
                trade.convert();

                tradeLine = String.Join(",", trade.items);
                if (trade.doNotInclude)
                {
                    tradeLine = null;
                    tbScreen.AppendText(String.Format("\r\n-!-!-!-> Excluded cash trade - port & sec currency are the same. : {0}", tradeLine));
                }


            }
            catch (Exception ex)
            {
                throw new Exception("convertMoxyTrade Func: " + ex.Message);
            }

            return tradeLine;
        }

        /// <summary>
        /// converts portia holdings files to Moxy format ready to import to Moxy
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_PortiaToMoxy_Click(object sender, EventArgs e)
        {
            string fType = string.Empty;     // file tyypes: holiday, groups, price, currency, portfolio, security, taxlot, broker
            string inPath = string.Empty;   // input file path  
            string outPath = string.Empty; // output file path  
            //HashSet<string> hsPortfolios = null; // to hold all portfolios coming from Portia
            try
            {
                tbScreen.Clear();

                PortiaMoxyManager pm = new PortiaMoxyManager(ref rtbScreen, ref lblStatus);
                // scroll to the end of text box
                tbScreen.SelectionStart = tbScreen.TextLength;
                tbScreen.ScrollToCaret();
            }
            catch (Exception fail)
            {
                String error = "The following error has occurred:\n\n";
                error += fail.Message.ToString() + "\n\n";
                MessageBox.Show(error);
                Globals.WriteErrorLog(fail.ToString());
                this.Close();
            }
        }// eof

        private void btn_FXTRades_AIM_New_Click(object sender, EventArgs e)
        {
            Globals.errCnt = 0;                                                         // reset error counter
            tbScreen.Clear();
                     
            try
            {
                MetaData mData = getAppMetaDataMoxy();
                tbScreen.Clear();

                // read through topost1.trn file to get fx trades
                FCTrades fctrades = getFCTrades(tbScreen, mData);

                convertFXToAIM(rtbScreen, mData, fctrades);

            }
            catch (Exception exp)
            {
                processException(exp);
             }


        }// end of event

        private void processException(Exception ex)
        {
            Globals.saveErr(ex.Message);
            Globals.WriteErrorLog(ex.ToString());
            MessageBox.Show(ex.Message);
        }


        FCTrades getFCTrades(TextBox screen, MetaData mData)
        {
            FCTrades fctrades = new FCTrades(rtbScreen);
            String line = null;
            try
            {
                Stream stream1 = new FileStream(mData.fxconSrcFile, FileMode.Open, FileAccess.Read,
                                                                        FileShare.ReadWrite);
                StreamReader sr1 = new StreamReader(stream1);
                do
                {
                    line = sr1.ReadLine();
                    if (String.IsNullOrEmpty(line)) continue;
                    if (line.IndexOf("cash") == -1)
                        fctrades.addTrade(line);
                }
                while (!(line == null));
                sr1.Close();

            }
             catch (Exception fail)
            {
                processException(fail);
             }
            return fctrades;
        }// eof


        private void convertFXToAIM(RichTextBox screen, MetaData mData, FCTrades fctrades)
        {
            try
            {
                int cnt = 0;
                Stream stream = new FileStream(mData.fxconSrcFile, FileMode.Open, FileAccess.Read,
                                                                        FileShare.ReadWrite);
                StreamReader sr = new StreamReader(stream);
                String file = Path.GetFileNameWithoutExtension(mData.outFolder + mData.aimFile) + Util.DateTimeStamp() 
                            + Path.GetExtension(mData.outFolder + mData.aimFile) ;
                String fnameAIM = mData.outFolder + file;
                String fnameAIMCopy = Util.getAppConfigVal("moxyAIMOutFolder") + file;
                String line = null;
                String portfolio = null;
                StreamWriter fwAIM = File.CreateText(fnameAIM);

                // read trades from trn file and convert them to AIM format
                do
                {
                   
                    line = sr.ReadLine();
                    if (String.IsNullOrEmpty(line)) continue;

                    // first five chars are portfolio code
                    portfolio = line.Substring(0, 5);

                    // analyze & fromat FX Connect trade for AIM
                    TradeFX trade = new TradeFX(screen, line, mData.moxyConStr, mData.portiaConStr,
                                                                        mData.tradingCurrencyStoredProc, mData.lastCrossRateProc,
                                                                        mData.reportingCurrencyStoredProc, fctrades, mData.sellRuleStoredProc);
                    if (!trade.doNotInclude)
                    {
                        trade.convert();
                        fwAIM.WriteLine(String.Join(",", trade.items));
                        cnt++;
                    }

                }
                while (!(line==null));
                sr.Close();
                fwAIM.Close();
                // make a copy of the file for portia 12

                File.Copy(fnameAIM, fnameAIMCopy);

                screen.AppendText("\r\n\r\n Created file for AIM: " +fnameAIM);
                screen.AppendText("\r\n\r\n Created copy of file for AIM: " + fnameAIMCopy);
                screen.AppendText("\r\n Trades in the file : " + cnt);

            }
              catch (Exception fail)
            {
                processException(fail);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "Portia Moxy 24 Import v." + Util.getAppConfigVal("Version") +
                              " " + Util.getAppConfigVal("moxy24constr").Replace("Integrated Security=SSPI", "");
        }

        string getSellingRule(string[] items, PortiaDatabase pd)
        {
            string sellingRule = string.Empty;

            if (String.IsNullOrEmpty(items[9]))
            {
                // no rule is set --> get the selling rule from portia
                sellingRule = pd.getSellingRule(items[0]);
            }
            else
            {
                sellingRule = getSellingRule(items[9], items[0]);
            }

            return sellingRule;
        }

        void addRoundingDiffToLastLot(int arrayInd, List<string> newLines, Double roundDiff)
        {
            if (newLines.Count > 0 &&  roundDiff != 0)
            {
                string[] lastLot = newLines[newLines.Count - 1].Split(',');
                // new trade amt
                lastLot[arrayInd] = (Double.Parse(lastLot[arrayInd]) + roundDiff).ToString();
                // remove last new line and add corrected
                newLines.RemoveAt(newLines.Count - 1);
                newLines.Add(String.Join(",", lastLot));
            }
        }

        private void buttonMoxyAIM_Click(object sender, EventArgs e)
        {
            Globals.errCnt = 0;           // reset error counter
            tbScreen.Clear();
                   
            DateTime asOfDate;

            try
            {
                MetaData mData = getAppMetaDataMoxy();

                // get date - by default is today
                FormSelectDate frmSelect = new FormSelectDate();

                if (frmSelect.ShowDialog(this) == DialogResult.OK)
                {
                    MoxyDatabase md = new MoxyDatabase(Util.getAppConfigVal("moxyconstr"), rtbScreen);
                    PortiaDatabase pd = new PortiaDatabase(mData.portiaConStr, rtbScreen, mData.tradingCurrencyStoredProc, mData.lastCrossRateProc, mData.sellRuleStoredProc);
                    asOfDate = frmSelect.getSelectedDate();

                    List<TrnLine>  trades = md.getMoxyExport(asOfDate);
                     tbScreen.AppendText(String.Format("\r\n\r\nNumber of trades: {0}", trades.Count));

                }
            } catch(Exception ex)
            {
                tbScreen.AppendText(Globals.saveErr(GetCurrentMethod() + ":" + ex.Message + "\r\n"));
            }

        }//eof

        /// <summary>
        /// Portia holdings to Moxy 24
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_PortiaToMoxy24_Click(object sender, EventArgs e)
        {

            string fType = string.Empty;     // file types: holiday, groups, price, currency, portfolio, security, taxlot, broker
            string inPath = string.Empty;   // input file path  
            string outPath = string.Empty; // output file path  
            HashSet<string> hsPortfolios = null; // to hold all portfolios coming from Portia
            try
            {
                PortiaMoxyManager pm = new PortiaMoxyManager(ref rtbScreen, ref lblStatus);
                MoxyDatabase md = new MoxyDatabase(Util.getAppConfigVal("moxy24constr"), rtbScreen);

                DataTable inFiles = md.getSrcFiles(Util.getAppConfigVal("getPortiaSrcFilesSP"));
                DataTable outFiles = md.getSrcFiles(Util.getAppConfigVal("getMoxyImportFilesSP"));

                // loop through each input file
                foreach (DataRow r in inFiles.Rows)
                {
                    fType = r["id"].ToString();
                    inPath = r["value"].ToString();
                    rtbScreen.AppendText(String.Format("File Type: {0} Source Path: {1} ", fType, inPath) + Environment.NewLine);

                    DataRow foundRow = outFiles.Rows.Find(fType);
                    if (foundRow == null)
                    {
                        ShowError(rtbScreen, String.Format("A row with the primary key of {0} could not be found in {1} ", fType, "outfiles"));
                       
                        return;
                    }

                    rtbScreen.AppendText(String.Format("File Type: {0} Destination Path: {1} ", foundRow[0], foundRow[1]) + Environment.NewLine);
                    outPath = foundRow[1].ToString();
                    Application.DoEvents();
                   
                   pm.convertForMoxy(fType, inPath, outPath, hsPortfolios);

                  ScrollToBottom();

                }

                }
                 catch (Exception fail)
            {
                HandleError(fail, closeForm: true);
            }
        }// eof

        private void ScrollToBottom()
        {
            tbScreen.SelectionStart = tbScreen.TextLength;
            tbScreen.ScrollToCaret();
            lblStatus.Text = "Ready";

        }

        private void HandleError(Exception ex, bool closeForm = false)
        {
            MessageBox.Show($"The following error has occurred:\n\n{ex.Message}",
                            "Error",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);

            Globals.WriteErrorLog(ex.ToString());

            if (closeForm)
                this.Close();
        }

        private void ShowError(RichTextBox rtb, string errText)
        {
            rtb.SelectionColor = Color.Red; // Set the desired color
            rtb.AppendText(errText +"\n"); // Add the text
            rtb.SelectionColor = rtb.ForeColor; // Reset to default color

        }

        private async void btnNTFXTradesDownload_ClickAsync(object sender, EventArgs e)
        {
            try {
                rtbScreen.Clear();
                // ask the date for the trades to download
                // SHOW POPUP DATE PICKER
                var dlg = new FormSelectDate();
                
                    var result = dlg.ShowDialog(this);

                    if (result != DialogResult.OK)
                    {
                        rtbScreen.AppendText("Trade date selection cancelled.\r\n");
                        return;
                    }

                    // User selected a date
                    DateTime tradeDate = dlg.getSelectedDate();
                    string dateSuffix = tradeDate.ToString("yyyyMMdd");
                    string filePattern = "*Tweedy*" + dateSuffix + ".csv";

                    var config = new SftpConfig
                    {
                        Host = "data-transfer.lumint.com",
                        Port = 22,
                        Username = "reporting_tweedy",
                        PrivateKeyPath = @"C:\keys\tb_nt_id_rsa",
                        PrivateKeyPassphrase = "ur+|u9enlv\"z/bQ;!48.",
                        Password = "",
                        RemoteDirectory = "/",
                        FilePattern = filePattern,
                        LocalDirectory = @"C:\Temp"
                    };

                    IDownloadNTFXTrades downloader = new NTSFTPDownloader(config);
                    var localFilePath = await downloader.DownloadFileAsync();

                    rtbScreen.AppendText(
                        $"Downloaded NT FX trades for {tradeDate:MM/dd/yyyy} → {localFilePath}\r\n");
                             
                                
                IGetNTFXTradesFromFile tradeReader = new NTFXTradesReader(localFilePath);
                List<NTFXTradeDTO> trades = await tradeReader.GetTradesFromFileAsync();

                if(trades.Count == 0)
                {
                    rtbScreen.AppendText($"No trades found in the downloaded file{localFilePath}.\n");
                    return;
                }

                rtbScreen.AppendText($"Read {trades.Count} trades from the file.\n");

                var outputFilePath = Path.Combine(
                    Util.getAppConfigVal("moxyAIMOutFolder"),
                    $"FXCon_{DateTime.Now:yyyyMMdd_HHmmss}.csv");

                string flipCurrencies = Util.getAppConfigVal("FlipRateCurrencies");
                List<string> flipCurrencyList = flipCurrencies.Split(',').Select(c => c.Trim().ToUpper()).ToList();

                IConvertNTFXTradesToAIM converter = new NTFXTradesConverter(trades, outputFilePath, flipCurrencyList);
                //converter.Convert();
                converter.ConvertWithAdjuster();

                //TO DO: check rounding in the output file


                //rtbScreen.AppendText($"Converted trades to AIM format and saved to: {outputFilePath}\n");
                var fileUri = new Uri(outputFilePath).AbsoluteUri;

                rtbScreen.AppendText($"Converted trades to AIM format and saved to:\n");
                rtbScreen.AppendText(outputFilePath + "\n");

                rtbScreen.AppendText($"Link to open the file:\n");
                rtbScreen.AppendText(fileUri + Environment.NewLine);
                rtbScreen.AppendText("\n");


            }
            catch (Exception ex)
            {
                HandleError(ex, closeForm: false);
                ShowError(rtbScreen, ex.Message);
            }
        }

        private void rtbScreen_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(e.LinkText);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not open link: " + ex.Message);
            }
        }

        private static void AppendHyperlink(RichTextBox box, string displayText, string url)
        {
            string escapedUrl = EscapeForRtf(url);
            string escapedText = EscapeForRtf(displayText);

            // RTF hyperfield fragment
            string rtf =
                @"{\field{\*\fldinst HYPERLINK """ + escapedUrl +
                @"""}{\fldrslt " + escapedText + "}}";

            box.SelectedRtf = rtf;
        }

        private static string EscapeForRtf(string text)
        {
            if (text == null) return string.Empty;

            return text
                .Replace(@"\", @"\\")
                .Replace("{", @"\{")
                .Replace("}", @"\}");
        }

    } // end of class
}  // end of namespace
