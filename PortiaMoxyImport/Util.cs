using PortiaMoxyImport.Entities;
using PortiaMoxyImport.Redesign;
using PortiaMoxyImport.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PortiaMoxyImport
{
    class Util
    {
        private const String cannotFindConfigItem = "-- <item> not found in config file. ???";
        private const String fileNotFound = "File <file> not found. ???";

        /// <summary>
        /// returns a value from config file by key
        /// </summary>
        /// <param name="name">key in the config file</param>
        /// <returns></returns>
        public static String getAppConfigVal(string name)
        {
            MethodBase mbase = new StackTrace().GetFrame(0).GetMethod();
            string methodName = mbase.DeclaringType.Name + "." + mbase.Name;

            try
            {
                var connection = System.Configuration.ConfigurationManager.AppSettings [name];
                return connection.ToString();
            }
            catch (NullReferenceException e)
            {
                string errMsg = e.Message;
                throw new Exception(methodName + cannotFindConfigItem.Replace("<item>", name));

            }
            catch (Exception ex)
            {
                throw new Exception(String.Format(methodName + ": " + ex.Message));

            }

        }// eof

        public static string DateTimeStamp()
        {
            return "_" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("hhmmss");
        }// eof

         


        public static void CombineFirstTwoColumns(string inputFilePath, string outputFilePath, string delimiter = ",")
        {
            var outputLines = new List<string>();

            using (var reader = new StreamReader(inputFilePath))
            {
                string line;
                int rowIndex = 0;

                while ((line = reader.ReadLine()) != null)
                {
                    var columns = SplitCsvLine(line);

                    if (columns.Count < 2)
                    {
                        outputLines.Add(line); // leave untouched if fewer than 2 columns
                        continue;
                    }

                    if (rowIndex == 0)
                    {
                        // Header line, keep as-is or modify if needed
                        outputLines.Add(line);
                    }
                    else
                    {
                        // Combine column 0 and 1
                        string combined = columns[0].Replace("\"", "") +  columns[1].Replace("\"", "");
                        columns.RemoveAt(0); // remove column 0
                        columns[0] = combined.ToLower(); // replace former column 1

                        outputLines.Add(string.Join(",", columns));
                    }

                    rowIndex++;
                }
            }

            File.WriteAllLines(outputFilePath, outputLines, Encoding.UTF8);
        }

        private static List<string> SplitCsvLine(string line)
        {
            var result = new List<string>();
            var sb = new StringBuilder();
            bool inQuotes = false;

            foreach (char c in line)
            {
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    sb.Append(c); // keep quotes if needed
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }

            result.Add(sb.ToString());
            return result;
        }

        public static string AddCombinedToFileName(string filePath)
        {
            string directory = Path.GetDirectoryName(filePath) ?? "";
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filePath);
            string extension = Path.GetExtension(filePath);

            string newFileName = $"{fileNameWithoutExtension}_combined{extension}";
            string newFilePath = Path.Combine(directory, newFileName);

            return newFilePath;
        }

        public static bool ContainsEmptyStrings(List<string> list)
        {
            return list.Any(item => string.IsNullOrEmpty(item));
        }

        /// <summary>
        /// Find and validate sourtce files for pending forwards, return the file paths of pending_forwards.csv
        /// </summary>
        /// <param name="srcFolde"></param>
        /// <returns></returns>
        /// <exception cref="FileNotFoundException"></exception>
        public static PendingForwardsResult LoadSourceFiles(string tb10FolderPath, string tb20FolderPath, DateTime reportDate)
        {
            try
            {

                DateTime confirmDate = AddBusinessDays(reportDate, 2);

                string datePrefix = string.Format("{0}.{1}.{2}",
                    confirmDate.Month,
                    confirmDate.Day,
                    confirmDate.ToString("yy"));

                if (!Directory.Exists(tb10FolderPath))
                    return PendingForwardsResult.Failure($"TB10 folder not found: {tb10FolderPath}");

                if (!Directory.Exists(tb20FolderPath))
                    return PendingForwardsResult.Failure($"TB20 folder not found: {tb20FolderPath}");

               

                string tb10Path = Path.Combine(tb10FolderPath, datePrefix + " PendingForwards_curr_TB10.pdf");
                string tb20Path = Path.Combine(tb20FolderPath, datePrefix + " PendingForwards_curr_TB20.pdf");

                // TEMP: dump raw lines and return
                //PendingForwardsParser.DumpRawLines(tb10Path, @"C:\Temp\tb10_raw_lines.txt");
                //MessageBox.Show("Done! Check C:\\Temp\\tb10_raw_lines.txt");
                
                if (!File.Exists(tb10Path))
                    return PendingForwardsResult.Failure($"TB10 file not found: {tb10Path}");

                if (!File.Exists(tb20Path))
                    return PendingForwardsResult.Failure($"TB20 file not found: {tb20Path}");

                DataTable tb10Data = PendingForwardsParser.ParsePdf(tb10Path, reportDate);
                DataTable tb20Data = PendingForwardsParser.ParsePdf(tb20Path, reportDate);

                if (tb10Data.Rows.Count == 0)
                    return PendingForwardsResult.Failure($"TB10 file parsed successfully but contained no trade records: {tb10Path}");

                if (tb20Data.Rows.Count == 0)
                    return PendingForwardsResult.Failure($"TB20 file parsed successfully but contained no trade records: {tb20Path}");

                return PendingForwardsResult.Ok(tb10Path, tb20Path, tb10Data, tb20Data);
            }
            catch (Exception ex)
            {
                return PendingForwardsResult.Failure($"Unexpected error loading source files: {ex.Message}");
            }
        }


        public static DateTime AddBusinessDays(DateTime date, int days)
        {
            int added = 0;
            DateTime result = date;

            while (added < days)
            {
                result = result.AddDays(1);
                if (result.DayOfWeek != DayOfWeek.Saturday && result.DayOfWeek != DayOfWeek.Sunday)
                    added++;
            }

            return result;
        }

        public static void DumpDataTable(DataTable dt, IConversionReporter reporter)
        {
            if (dt == null)
            {
                reporter.Error("DataTable is null.");
                return;
            }

            reporter.Info($"DataTable: {dt.TableName} — {dt.Columns.Count} columns, {dt.Rows.Count} rows");

            // Column names
            List<string> colNames = new List<string>();
            foreach (DataColumn col in dt.Columns)
                colNames.Add(col.ColumnName);
            reporter.Info("Columns: " + string.Join(", ", colNames));

            // Rows
            int rowNum = 0;
            foreach (DataRow row in dt.Rows)
            {
                List<string> vals = new List<string>();
                foreach (DataColumn col in dt.Columns)
                    vals.Add(string.Format("{0}={1}", col.ColumnName, row[col] == DBNull.Value ? "NULL" : row[col].ToString()));
                reporter.Info(string.Format("Row {0}: {1}", rowNum++, string.Join(" | ", vals)));
            }
        }

    }//eoc


}
