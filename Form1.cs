using Microsoft.Win32;
using MySql.Data.MySqlClient;
using System;
using System.Management;
using System.Net.NetworkInformation;
using System.Windows.Forms;
using OpenHardwareMonitor.Hardware;
using System.Linq;

namespace ITInventory
{
    public partial class Form1 : Form

    {
        public static string userGivenName;
        public static string userSurname;
        public static string userLogin;
        public static string cpuName;
        public static double Ram;
        public static string OS;
        public static string OSName;
        public static string OSDisplayVersion;
        public static string OSCurrentBuild;
        public static string OSType;
        public static string currentDomain;
        public static string computerName;
        public static string comModel;
        public static string SerialNumber;
        public static string Manufacturer;
        public static string ethIP;
        public static string ethType;
        public static string ethMac;
        public static string wifiIP;
        public static string wifiType;
        public static string wifiMac;

        public Form1()
        {
            InitializeComponent();
        }

        public void CallFunction()
        {
            GetMonitorDetails();
            GetComputerDetials();
            GetUserInfo();
            CheckTimetoUpdate();
        }
        /*public void GetBrowserURL()
        {

        }*/
        public void Form1_Load(object sender, EventArgs e)
        {
            CheckConn();
            GetMonitorDetails();
            GetComputerDetials();
            GetUserInfo();
            GetSoftwareInfo();
            //GetBrowserURL();

            Timer timer = new Timer
            {
                Interval = 1000 // 1 secs
            };

            timer.Tick += new EventHandler(Timer1_Tick);
            timer.Start();
        }
        private void GetMonitorDetails()
        {
            var monitorsByMonitorIds = MonitorInfo.GetNamesByMonitorIds();
            // foreach (var monitorByMonitorId in monitorsByMonitorIds)
            MonitorManufacturerlbl.Text = monitorsByMonitorIds[0];
            MonitorModellbl.Text = monitorsByMonitorIds[1];
            MonitorSerialNolbl.Text = monitorsByMonitorIds[2];
        }
        public void GetComputerDetials()
        {

            cpuName = null;
            string MemorySize = null;
            OSType = null;
            currentDomain = System.Net.NetworkInformation.IPGlobalProperties.GetIPGlobalProperties().DomainName;
            computerName = System.Environment.MachineName;

            using (var view64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine,
                                            RegistryView.Registry64))
            {
                using (var cls32 = view64.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion", false))
                {
                    var windowsName = view64.GetValue(@"ProductName");
                    OSName = cls32.GetValue("ProductName").ToString();
                    OSDisplayVersion = cls32.GetValue("DisplayVersion").ToString();
                    OSCurrentBuild = cls32.GetValue("CurrentBuild").ToString();
                    OS = OSName + " " + OSDisplayVersion + " " + OSCurrentBuild;
                }

            }

            ManagementClass mcCPU = new ManagementClass("Win32_Processor");
            ManagementObjectCollection mocCPU = mcCPU.GetInstances();

            foreach (ManagementObject mo in mocCPU.Cast<ManagementObject>())
            {
                cpuName = mo.Properties["Name"].Value.ToString();
                break;
            }

           
            ManagementClass mcMemmory = new ManagementClass("win32_physicalmemory");
            ManagementObjectCollection mocMemmory = mcMemmory.GetInstances();
            foreach (ManagementObject mo in mocMemmory)
            {
                MemorySize = mo.Properties["Capacity"].Value.ToString();
                break;
            }
            double doubleVal = Convert.ToDouble(MemorySize);
            Ram = doubleVal / 1024 / 1024 / 1024;

            ManagementClass mcOsType = new ManagementClass("Win32_OperatingSystem");
            ManagementObjectCollection mocOsType = mcOsType.GetInstances();
            foreach (ManagementObject mo in mocOsType)
            {
                OSType = mo.Properties["OSArchitecture"].Value.ToString();
            }
            ManagementClass mcComSN = new ManagementClass("Win32_Bios");
            ManagementObjectCollection mocComSN = mcComSN.GetInstances();
            foreach (ManagementObject mo in mocComSN)
            {
                SerialNumber = mo.Properties["SerialNumber"].Value.ToString();
                Manufacturer = mo.Properties["Manufacturer"].Value.ToString();
            }
            ManagementClass mcWin32_ComputerSystem = new ManagementClass("Win32_ComputerSystem");
            ManagementObjectCollection mocWin32_ComputerSystem = mcWin32_ComputerSystem.GetInstances();
            foreach (ManagementObject mo in mocWin32_ComputerSystem)
            {
                //label10.Text = mo.Properties["SerialNumber"].Value.ToString();
                comModel = mo.Properties["Model"].Value.ToString();
            }

            foreach (NetworkInterface ni in NetworkInterface.GetAllNetworkInterfaces())
            {
                if (ni.NetworkInterfaceType == NetworkInterfaceType.Ethernet)
                {
                    if (ni.Name.StartsWith("vEthernet") == false && ni.OperationalStatus == OperationalStatus.Up && ni.NetworkInterfaceType != NetworkInterfaceType.Loopback && ni.NetworkInterfaceType != NetworkInterfaceType.Tunnel)
                    {
                        foreach (UnicastIPAddressInformation ip in ni.GetIPProperties().UnicastAddresses)
                        {
                            if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                            {
                                ethIP = ip.Address.ToString();
                                ethType = ni.NetworkInterfaceType.ToString();
                                ethMac = ni.GetPhysicalAddress().ToString();
                            }
                        }
                    }
                }
                else if (ni.NetworkInterfaceType == NetworkInterfaceType.Wireless80211)
                {
                    if (ni.Name.StartsWith("vEthernet") == false && ni.OperationalStatus == OperationalStatus.Up && ni.NetworkInterfaceType != NetworkInterfaceType.Loopback && ni.NetworkInterfaceType != NetworkInterfaceType.Tunnel)
                    {
                        foreach (UnicastIPAddressInformation ip in ni.GetIPProperties().UnicastAddresses)
                        {
                            if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                            {
                                wifiIP = ip.Address.ToString();
                                wifiType = ni.NetworkInterfaceType.ToString();
                                wifiMac = ni.GetPhysicalAddress().ToString();
                            }
                        }
                    }
                }
            }


            //Start Get Computer Asset NO. from Active directory
            // Get Data from AD Attribuit

            var comName = DatabaseProcess.GetComputerAssetNo();
            label40.Text = comName;
            label1.Text = cpuName;
            label2.Text = Ram + " GB";
            label5.Text = OS;
            label4.Text = computerName;
            label6.Text = OSType;
            label11.Text = currentDomain;
            label10.Text = SerialNumber;
            label12.Text = Manufacturer;
            label8.Text = ethIP;
            label7.Text = ethType;
            label26.Text = ethMac;
            label27.Text = wifiIP;
            label29.Text = wifiType;
            label28.Text = wifiMac;
            PCModellbl.Text = comModel;

        }
        public void GetUserInfo()
        {
            string userFullname = System.DirectoryServices.AccountManagement.UserPrincipal.Current.DisplayName;
            userGivenName = System.DirectoryServices.AccountManagement.UserPrincipal.Current.GivenName;
            userSurname = System.DirectoryServices.AccountManagement.UserPrincipal.Current.Surname;

            userLogin = Environment.UserName;
            //un = userFullname;
            label3.Text = userLogin;
            label9.Text = userFullname;
        }
        private void GetSoftwareInfo()
        {
            var getAntiVirusinfo = SoftwareInfo.getAntiVirusinfo();
            var gsoftwareInfo = SoftwareInfo.getMSOfficeInfo();
            foreach (var antiVirus in getAntiVirusinfo)
            {
                dataGridView2.Rows.Insert(0, antiVirus[0], antiVirus[1], antiVirus[2], antiVirus[3]);
            }
            foreach (var m in gsoftwareInfo)
            {
                dataGridView1.Rows.Insert(0, m[1], m[0], m[2], m[3]);
            }
        }
        /*private void CheckTimetoUpdate()
        {
            var start = DateTime.Now;
            var oldDate = DateTime.Parse("16/07/2021 09:11:00");

            if ((start - oldDate).TotalMinutes >= 1)
            {
                MessageBox.Show("Please update");
            }
        }*/
        protected void CheckTimetoUpdate()
        {

            // string strAccountId = "NDTH-PC-007";

            /*  DirectoryEntry de = new DirectoryEntry();
              de.Path = "LDAP://DC=nichidai,DC=co,DC=th";
              de.AuthenticationType = AuthenticationTypes.Secure;

              DirectorySearcher deSearch = new DirectorySearcher();

              deSearch.SearchRoot = de;
              deSearch.Filter = "(&(objectClass=computer) (cn=" + computerName + "))";

              SearchResult result = deSearch.FindOne();

              if (result != null)
              {

                  DirectoryEntry deUser = new DirectoryEntry(result.Path);
                  if (deUser.Properties["nDTComputerAsset"].Value != null)
                  {
                      string testss = deUser.Properties["nDTComputerAsset"].Value.ToString();
                      //MessageBox.Show(testss);
                      label40.Text = testss;
                  }
                  else {
                      MessageBox.Show("Not found");
                  }
                  deUser.Close();
              }
            */
        }
        /* Private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
         {
             //Show();
             this.WindowState = FormWindowState.Normal;
             this.ShowInTaskbar = true;
             notifyIcon1.Visible = false;
             this.Visible = true;
         }*/
        public void ExitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            TaskTrayApplicationContext nt = new TaskTrayApplicationContext();
            nt.ExitProgram(sender, e);
            //Environment.Exit(0);
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            TaskTrayApplicationContext nt = new TaskTrayApplicationContext();
            nt.ClickCloseMinimize(e);
            //Environment.Exit(0);
        }
        private void CheckConn()
        {
            MySqlConnection conn = DBConnection.GetDBConnection();
            try
            {
                conn.Open();
                toolStripStatusLabel2.Text = "The Database connected.";
                conn.Close();
            }
            catch
            {
                //this.toolStripStatusLabel2.Text = "Failed to connect to database.";
                toolStripStatusLabel2.Text = "Failed to connect to database.";
                conn.Close();
            }
        }
        private void Timer1_Tick(object sender, EventArgs e)
        {
            string CurrentTime = DateTime.Now.ToString();
            toolStripStatusLabel1.Text = CurrentTime;

        }

    }
}
