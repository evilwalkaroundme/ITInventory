using System;
using System.Collections.Generic;
using System.Management;
using System.Text;

namespace ITInventory
{
    public static class MonitorInfo
    {
        public static string UnknownMonitor => "UnknownMonitor";
        public static string[] GetNamesByMonitorIds()
        {

            var result = new List<string>();
            ManagementObjectSearcher searcher = null;
            ManagementObjectCollection monitors = null;

            try
            {
                var scope = new ManagementScope("\\\\.\\ROOT\\WMI");
                var query = new ObjectQuery("SELECT * FROM WmiMonitorID");
                searcher = new ManagementObjectSearcher(scope, query);

                monitors = searcher.Get();
                if (monitors.Count > 0)
                {
                    foreach (var monitor in monitors)
                    {
                        string ManufacturerName = monitor["ManufacturerName"].AsString();
                        result.Add(!string.IsNullOrEmpty(ManufacturerName) && !ManufacturerName.Contains("PnP") ? ManufacturerName : UnknownMonitor);

                        string userFriendlyName = monitor["UserFriendlyName"].AsString();
                        result.Add(!string.IsNullOrEmpty(userFriendlyName) && !userFriendlyName.Contains("PnP") ? userFriendlyName : UnknownMonitor);

                        string SerialNumberID = monitor["SerialNumberID"].AsString();
                        result.Add(!string.IsNullOrEmpty(SerialNumberID) && !SerialNumberID.Contains("PnP") ? SerialNumberID : UnknownMonitor);
                    }
                }
            }
            catch (Exception)
            {
                // ignored
            }
            finally
            {
                monitors?.Dispose();
                searcher?.Dispose();
            }

            return result.ToArray();
        }
        private static string AsString(this object obj)
        {
            switch (obj)
            {
                case null:
                case IReadOnlyList<ushort> decArray when decArray.Count == 0 || decArray[0] == 0:
                    return string.Empty;
                case IReadOnlyList<ushort> decArray:
                    {
                        var sb = new StringBuilder();
                        foreach (ushort val in decArray)
                        {
                            if (val == 0)
                                break;

                            // ASCII codes only
                            if (val >= 32 && val <= 127)
                                sb.Append((char)val);
                        }

                        return sb.ToString().Trim();
                    }
                default:
                    return string.Empty;
            }
        }
    }
}
