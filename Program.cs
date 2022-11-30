using System;
using System.Diagnostics;
using System.Windows.Forms;


namespace ITInventory
{
    static class Program
    {
        ///test
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Update information to database every 2Hr.
            var startTimeSpan = TimeSpan.Zero;
            var periodTimeSpan = TimeSpan.FromMinutes(120);
            var timer = new System.Threading.Timer((e) =>
            {
                DatabaseProcess dbp = new DatabaseProcess();
                dbp.UpdateComputerList();
            }, null, startTimeSpan, periodTimeSpan);
            //Update information to database every 2Hr.

            Process aProcess = Process.GetCurrentProcess();
            string aProcName = aProcess.ProcessName;


            if (Process.GetProcessesByName(aProcName).Length > 1)
            {
                MessageBox.Show("The application is already running.");
                Application.ExitThread();
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                //DatabaseProcess dbp = new DatabaseProcess();
                //dbp.UpdateComputerList();
                Application.Run(new TaskTrayApplicationContext());

            }
        }

    }
}
