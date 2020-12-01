using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using System.Xml; 
using System.Data.SQLite; 


//
// NOT USED
//
namespace PortiaMoxyImport
{
    class SQLiteManager
    {
        string SQLITEDB = @"Data Source= ;Version=3;";
        public TextBox screen; // for communication with UI
        //public Label status;      // for showing progress of the task
        public SQLiteManager(ref TextBox aScreen)
        {
            screen = aScreen;
            
        }

           public int getSQLiteValue( String table, String sectionVal, String idVal, ref String keyVal )
           {
                // This function reads the value from settings table in SQLite DB
                // This is a replacement of INI file
               int rtn = 0;
               String sql = string.Empty;

               try {
                   SQLiteConnection m_dbConnection;
                    m_dbConnection = new SQLiteConnection(SQLITEDB);
                    m_dbConnection.Open();
                    sql = @"select value from " + table + " where section= '" + sectionVal + "'  AND id ='" + idVal + "' ";
                    SQLiteCommand   Command;
                    Command = new SQLiteCommand(sql, m_dbConnection);
                    SQLiteDataReader reader;
                    reader = Command.ExecuteReader();
                    while (reader.Read())
                    {
                         //keyVal = reader(0).ToString();
                    } // end of while loop
              

               }

               catch (Exception ex)
                {
                    screen.AppendText( "getSQLiteValue: " + ex.Message + Environment.NewLine);
                    rtn = -1;
                }

                return rtn;
           }


        
    }
}
