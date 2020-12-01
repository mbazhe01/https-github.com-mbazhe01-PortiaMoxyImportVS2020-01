using System;
using System.Collections.Generic;
using System.Diagnostics;
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


    }
}
