using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Finisar.SQLite;

namespace helper
{
   public class Database
    {
    
        public void AddRecord(string english,string arabic)
        {
            SQLiteConnection sqlite_conn;
            SQLiteCommand sqlite_cmd;
           
            sqlite_conn = new SQLiteConnection("Data Source=Database.db;Version=3;New=False;Compress=True;");
                       
                sqlite_conn.Open();
          
           
            sqlite_cmd = sqlite_conn.CreateCommand();
         sqlite_cmd.CommandText = "CREATE TABLE IF NOT EXISTS `Translation` (`ID`	INTEGER PRIMARY KEY AUTOINCREMENT,`English`	TEXT,`Arabic`	TEXT); ";

            // Now lets execute the SQL ;D
           sqlite_cmd.ExecuteNonQuery();
            sqlite_cmd.CommandText = "INSERT INTO Translation (English, Arabic) VALUES ('"+english+"','"+ arabic+"');";
          

            sqlite_cmd.ExecuteNonQuery();
           sqlite_conn.Close();
        }
        public string getRecord(string english)
        {
            string arabic="";
            SQLiteConnection sqlite_conn;
            SQLiteCommand sqlite_cmd;
            SQLiteDataReader sqlite_datareader;

            sqlite_conn = new SQLiteConnection("Data Source=Database.db;Version=3;New=False;Compress=True;");

            sqlite_conn.Open();

          
            sqlite_cmd = sqlite_conn.CreateCommand();
            //   sqlite_cmd.CommandText = "CREATE TABLE IF NOT EXISTS `Translation` (`ID`	INTEGER PRIMARY KEY AUTOINCREMENT,`English`	TEXT,`Arabic`	TEXT); ";

            // Now lets execute the SQL ;D
            //    sqlite_cmd.ExecuteNonQuery();

            sqlite_cmd.CommandText = "SELECT Arabic FROM Translation WHERE English LIKE '"+english+"'";
            if (Convert.ToString(sqlite_cmd.ExecuteScalar()) == string.Empty)
            {
                arabic = english;
            }
            else
            {
                arabic=Convert.ToString(sqlite_cmd.ExecuteScalar());


            }
            
            // Now the SQLiteCommand object can give us a DataReader-Object:
            sqlite_datareader = sqlite_cmd.ExecuteReader();
            
            
            sqlite_conn.Close();



            return arabic;
        }
    }
}
