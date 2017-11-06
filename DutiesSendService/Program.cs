using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DutiesSendService
{
    class Program
    {
        private static Mutex mutex = new Mutex(true, Process.GetCurrentProcess().ProcessName);
        static void Main(string[] args)
        {
            if (!mutex.WaitOne(TimeSpan.Zero, true))
            {
                MessageBox.Show("Выдача нарядов уже запущена");
                return;
            }

            var notifyThread = new Thread(() =>
            {
                var trayIcon = new TrayIcon();

                trayIcon.AddIcon();
            });
            notifyThread.Start();

            var worker = new Worker();
            worker.ReadSettings();
            worker.Start();
        }
    }
}
