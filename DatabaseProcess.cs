using MySql.Data.MySqlClient;

namespace ITInventory
{
    public class DatabaseProcess
    {

        public void UpdateComputerList()
        {

            GetMSSoftwareInfo();
            MySqlConnection conn = DBConnection.GetDBConnection();
            try
            {
                // Call Data on Form1 Class                
                Form1 callFuctionForm1 = new Form1();
                callFuctionForm1.CallFunction();
                // Call Data on Monitor Class
                var monitorsByMonitorIDs = MonitorInfo.GetNamesByMonitorIds();
                // Call Data on Anti-Virus infoemation Class
                var antivirusID = SoftwareInfo.getAntiVirusinfo();
                // Retrieve details of WIFI or LAN.
                string ethIP = Form1.ethIP;
                string ethmac = Form1.ethMac;
                string ethType = Form1.ethType;
                //Detect Wifi
                if (Form1.ethType == null) {
                    //ethType = Form1.wifiType;
                    ethType = "WiFi";
                   ethIP = Form1.wifiIP;
                   ethmac = Form1.wifiMac;
                }
                string Query = "CALL AddComputerDetails('" + Form1.computerName + "','AssetNo','" + Form1.OS + "','"
                                                         + Form1.OSType + "','" + Form1.Manufacturer + "','" + Form1.comModel + "','"
                                                         + Form1.SerialNumber + "','" + Form1.cpuName + "','" + Form1.Ram + "','"
                                                         + ethType + "','" + ethIP + "','" + ethmac + "','" + Form1.currentDomain + "','"
                                                         + Form1.userGivenName + "','" + Form1.userSurname + "','" + Form1.userLogin + "','"
                                                         + monitorsByMonitorIDs[0] + "','" + monitorsByMonitorIDs[1] + "','" + monitorsByMonitorIDs[2] + "','"
                                                         + antivirusID[0][0] + "','" + antivirusID[0][1] + "','" + antivirusID[0][2] + "','" + antivirusID[0][3] + "')";
                MySqlCommand MySqlCmd = new MySqlCommand(Query, conn);
                
                conn.Open();
                MySqlDataReader myreader = MySqlCmd.ExecuteReader();
                conn.Close();
            }
            catch
            {
                conn.Close();
            }
        }
        //Get Computer Asset from Database to display on Form1
        public static string GetComputerAssetNo()
        {
            string ComputerAssetNo = null;
            MySqlConnection conn = DBConnection.GetDBConnection();

            try
            {
                Form1 getComputername = new Form1();
                string getComname = Form1.computerName;

                string Query = "call SelectComputerAsset ('" + getComname + "')";
                MySqlCommand mySQLcmd = new MySqlCommand(Query, conn);



                conn.Open();
                MySqlDataReader mySQLReader = mySQLcmd.ExecuteReader();
                while (mySQLReader.Read())
                {
                    ComputerAssetNo = mySQLReader.GetString("AssetNo");
                }

                conn.Close();
            }
            catch
            {
                conn.Close();

            }

            return ComputerAssetNo;
        }
        public static void GetMSSoftwareInfo()
        {
            MySqlConnection conn = DBConnection.GetDBConnection();

            var msSoftwareList = SoftwareInfo.getMSOfficeInfo();

            foreach (var m in msSoftwareList)
            {

                try
                {
                    Form1 callFuctionForm1 = new Form1();
                    callFuctionForm1.CallFunction();
                    Form1 getComputername = new Form1();
                    string getComname = Form1.computerName;
                    string Query = "CALL AddMSSoftwareInfo ('" + m[1] + "','" + m[0] + "','" + m[2] + "','" + getComname + "','" + m[3] + "')";
                    MySqlCommand mySQLcmd = new MySqlCommand(Query, conn);

                    conn.Open();
                    MySqlDataReader mySQLReader = mySQLcmd.ExecuteReader();
                    conn.Close();

                }
                catch
                {
                    conn.Close();
                }
            }
        }


    }

}
