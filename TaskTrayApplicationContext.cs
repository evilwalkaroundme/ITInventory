using System;
using System.Windows.Forms;

namespace ITInventory
{
    public class TaskTrayApplicationContext : ApplicationContext
    {
        NotifyIcon notifyIcon = new NotifyIcon();
        Form1 WindowInfo = new Form1();


        public TaskTrayApplicationContext()
        {
            MenuItem configMenuItem = new MenuItem("Configuration", new EventHandler(ShowConfig));
            MenuItem exitMenuItem = new MenuItem("Exit", new EventHandler(ExitProgram));

            notifyIcon.Icon = ITInventory.Properties.Resources.check;
            notifyIcon.ContextMenu = new ContextMenu(new MenuItem[] { configMenuItem, exitMenuItem });
            notifyIcon.Visible = true;

        }
        void ShowConfig(object sender, EventArgs e)
        {
            WindowInfo.Visible = true;
            notifyIcon.Visible = false;
        }

        public void ExitProgram(object sender, EventArgs e)
        {
            // We must manually tidy up and remove the icon before we exit.
            // Otherwise it will be left behind until the user mouses over.
            notifyIcon.Visible = false;
            Environment.Exit(0);
        }
        public void MinimizeWindows(EventArgs e)
        {
            notifyIcon.Visible = false;
        }

        public void ClickCloseMinimize(EventArgs e)
        {
            WindowInfo.Visible = false;
            notifyIcon.Visible = true;
        }
    }
}
