using MySql.Data.MySqlClient;
using System.Configuration;


namespace ITInventory
{
    class DBConnection
    {
       
        static MySqlConnection databaseConnection = null;
        public static MySqlConnection GetDBConnection()
        {
            if (databaseConnection == null)
            {
                string connectionString = ConfigurationManager.ConnectionStrings["myDatabaseConnection"].ConnectionString;
                databaseConnection = new MySqlConnection(connectionString);
            }
            return databaseConnection;
        }

    }
}
