using System;
using Oracle.ManagedDataAccess.Client;
using MySql.Data.MySqlClient;

namespace RoboReestrService
{
    class SQLConnect
    {

        public string senderId;
        public string aggregator;
        public OracleConnection oracleConnection;
        public MySqlConnection mysqlConnection;


        public SQLConnect(string oraclehost, string oracleport, string oracledatabase, string oracleuser, string oraclepass, string mysqlhost, string mysqlport, 
            string mysqldatabase, string mysqluser, string mysqlpass, string sender, string agg) 
        {
            senderId = sender;
            aggregator = agg;
            try
            {
                oracleConnection = new OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" + oraclehost + ")(PORT=" + oracleport +
                "))(CONNECT_DATA=(SERVICE_NAME=" + oracledatabase + ")));User Id=" + oracleuser + ";Password=" + oraclepass + ";");
                mysqlConnection = new MySqlConnection("Server=" + mysqlhost + ";Port=" + mysqlport + ";Database=" + mysqldatabase + ";Uid=" + mysqluser + ";Pwd=" 
                    + mysqlpass + ";");
            }
            catch (Exception ex) 
            {
                Logger.Log.Error("SQLCONNECT ERROR " + ex.ToString());
            }
            
        }

    }
}
