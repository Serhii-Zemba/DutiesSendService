using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DutiesSendService
{
    public class TrayIcon
    {
        public void AddIcon()
        {

            var icon = new NotifyIcon();
            icon.ContextMenu = new ContextMenu();
            icon.ContextMenu.MenuItems.Add("Выход", CloseProgram);
            icon.Icon = new Icon("reload.ico");
            icon.Visible = true;
            icon.Text = "Выдача нарядов";
            Application.Run();
        }

        private void CloseProgram(object sender, EventArgs e)
        {
            Environment.Exit(0);
        }
    }
}
