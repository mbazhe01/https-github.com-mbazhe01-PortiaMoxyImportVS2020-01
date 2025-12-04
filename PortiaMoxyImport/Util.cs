using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

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

    }


}
