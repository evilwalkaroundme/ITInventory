using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Management;

namespace ITInventory
{
    public static class SoftwareInfo
    {
        public static string UnknowSoftware => "UnknowSoftware";
        public static DateTime timeCreated;
        public static List<string[]> getMSOfficeInfo()
        {
            ManagementObjectCollection msSoftware = null;
            List<string[]> result = new List<string[]>();
            String[,] temp;
            string SoftwareString;

            try
            {

                string filePath64x = @"C:\Program Files\Microsoft Office\Templates";
                string filepath86x = @"C:\Program Files (x86)\Microsoft Office\Templates";
                if (Directory.Exists(filePath64x))
                {
                    timeCreated = File.GetCreationTime(@"C:\Program Files\Microsoft Office\Templates");
                }
                else if (Directory.Exists(filepath86x))
                {
                    timeCreated = File.GetCreationTime(@"C:\Program Files (x86)\Microsoft Office\Templates");

                }
                else
                {
                    //MessageBox.Show("Not found");
                }

                ManagementObjectSearcher LicenseSearcher =
                           new ManagementObjectSearcher("root\\CIMV2",
                           "SELECT Description ,PartialProductKey,ProductKeyId2,LicenseFamily " +
                           "FROM SoftwareLicensingProduct " +
                           "WHERE ProductKeyId <> null AND Description like '%Office%'");

                msSoftware = LicenseSearcher.Get();
                temp = new String[msSoftware.Count, 3];

                if (msSoftware.Count > 0)
                {
                    int i = 0;
                    foreach (var item in msSoftware)
                    {

                        SoftwareString = item["Description"].ToString().ToLower();

                        if (SoftwareString.Contains("office"))
                        {

                            string col1 = temp[i, 0] = item["PartialProductKey"].ToString();
                            string col2 = temp[i, 1] = item["LicenseFamily"].ToString();
                            string col3 = temp[i, 2] = item["ProductKeyID2"].ToString();

                            i++;
                            result.Add(new[] { col1, col2, col3, timeCreated.ToString("yyyy-MM-dd h:mm:ss") });

                        }
                    }
                }
                else
                {
                    string sVersion = string.Empty;
                    string sLicenseFamily = "N/A";
                    string sProductKeyID = "N/A";
                    Microsoft.Office.Interop.Word.Application appVersion = new Microsoft.Office.Interop.Word.Application();
                    appVersion.Visible = false;
                    switch (appVersion.Version.ToString())
                    {
                        case "11.0":
                            sVersion = "Microsoft Offic 2003";
                            break;
                        case "12.0":
                            sVersion = "Microsoft Offic 2007";
                            break;
                        case "14.0":
                            sVersion = "Microsoft Offic 2010";
                            break;
                        default:
                            sVersion = "Too Old!";
                            break;
                    }
                    result.Add(new[] { sLicenseFamily, sVersion, sProductKeyID, timeCreated.ToString("yyyy-MM-dd h:mm:ss") });
                }
            }
            catch (Exception LSOexception)
            {
                Console.WriteLine(LSOexception.Message);
            }
            return result;
        }
        public static List<string[]> getAntiVirusinfo()
        {
            List<string[]> result = new List<string[]>();

            try
            {
                using (RegistryKey regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\WOW6432Node\\Symantec\\Symantec Endpoint Protection\\CurrentVersion"))
                using (RegistryKey opstateregKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\WOW6432Node\\Symantec\\Symantec Endpoint Protection\\CurrentVersion\\public-opstate"))
                using (RegistryKey SyLinkregKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\\WOW6432Node\\Symantec\\Symantec Endpoint Protection\\SMC\\SYLINK\\SyLink"))


                    if (regKey != null)
                    {

                        Object oProductName = regKey.GetValue("PRODUCTNAME");
                        Object oPRODUCTVERSION = regKey.GetValue("PRODUCTVERSION");
                        Object oLatestVirusDefsDate = opstateregKey.GetValue("LatestVirusDefsDate");
                        object oLastConnectedTime = SyLinkregKey.GetValue("LastConnectedTime");
                        //Version strPRODUCTNAME = new Version(oProductName as String);
                        result.Add(new[] { oProductName.ToString(), oPRODUCTVERSION.ToString(), oLatestVirusDefsDate.ToString(), oLastConnectedTime.ToString().Substring(0, 10) });


                    }
            }
            catch
            {

            }

            return result;
        }

    }
}
