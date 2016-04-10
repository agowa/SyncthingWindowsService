using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace $safeprojectname$
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.CancelKeyPress += delegate
            {
                System.Diagnostics.Debug.WriteLine("Nicely exit this App");
            };
            int i = 0;
            while (true)
            {
                System.Console.Out.WriteLine("Line Nr. " + i++);
                Thread.Sleep(200);
            }
        }
    }
}
