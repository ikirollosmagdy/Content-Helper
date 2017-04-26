﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Finisar.SQLite;

namespace helper
{
   public class Database
    {
        static bool IsEgypt = Properties.Settings.Default.IsEgypt;

        public void AddRecord(string english,string arabic)
        {
            SQLiteConnection sqlite_conn;
            SQLiteCommand sqlite_cmd;
            //Database_UAE
            if (IsEgypt)
            {
                sqlite_conn = new SQLiteConnection("Data Source=Database_Egypt.db;Version=3;New=False;Compress=True;");
            }
            else
            {
                sqlite_conn = new SQLiteConnection("Data Source=Database_UAE.db;Version=3;New=False;Compress=True;");
            }
                       
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

            if (IsEgypt)
            {
                sqlite_conn = new SQLiteConnection("Data Source=Database_Egypt.db;Version=3;New=False;Compress=True;");
            }
            else
            {
                sqlite_conn = new SQLiteConnection("Data Source=Database_UAE.db;Version=3;New=False;Compress=True;");
            }

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

        public string getEAN(string title)
        {
            List<string> eans = new List<string>();
            string EAN = "";
            SQLiteConnection sqlite_conn;
            SQLiteCommand sqlite_cmd;
            SQLiteDataReader sqlite_datareader;

            sqlite_conn = new SQLiteConnection("Data Source=EAN.db;Version=3;New=False;Compress=True;");

            sqlite_conn.Open();


            sqlite_cmd = sqlite_conn.CreateCommand();
           

            sqlite_cmd.CommandText = "SELECT EAN FROM Suggested WHERE Title LIKE '" + title + "%'";
            // Looking for Ean by Values

            sqlite_datareader = sqlite_cmd.ExecuteReader();

            // The SQLiteDataReader allows us to run through the result lines:
            while (sqlite_datareader.Read()) // Read() returns true if there is still a result line to read
            {
                eans.Add(sqlite_datareader["EAN"].ToString());
                
            }
            
                for(int x = 0; x <eans.Count; x++)
                {

                EAN = EAN  + eans[x] + ", ";
                
                }
            


            sqlite_conn.Close();



            return EAN;

        }
    }
}
