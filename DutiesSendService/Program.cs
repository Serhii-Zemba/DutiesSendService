using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace DutiesSendService
{
    class Program
    {
        static void Main(string[] args)
        {

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
