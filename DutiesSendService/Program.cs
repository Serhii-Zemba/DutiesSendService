using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace DutiesSendService
{
    class Program
    {
        static void Main(string[] args)
        {
            var worker = new Worker();
            worker.ReadSettings();
            worker.Start();
        }
    }
}
