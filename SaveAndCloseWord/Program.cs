using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Win32_API;
using System.Globalization;

namespace SaveAndCloseWord
{
    class Program
    {
        [DllImport("User32.dll")]
        private static extern bool
         GetLastInputInfo(ref LASTINPUTINFO plii);

        static void Main(string[] args)
        {
            Run();
            
        }

        static void Run()
        {
            while (true)
            {
                System.Threading.Thread.Sleep(600000);
                if (Win32.GetIdleTime() > 600000 && (System.DateTime.Now.Hour > 6 && System.DateTime.Now.Minute > 10 && System.DateTime.Now.ToString("tt", CultureInfo.InvariantCulture) == "PM"))
                {
                    foreach (var process in System.Diagnostics.Process.GetProcessesByName("winword"))
                    {
                        try
                        {
                            Console.Write("The Time is " + System.DateTime.Now.Hour.ToString() + ":" + System.DateTime.Now.Minute +" " + System.DateTime.Now.ToString("tt", CultureInfo.InvariantCulture));
                            Console.WriteLine(" Time to close word");
                            process.Kill();
                        }
                        catch { }
                    }
                }
            }
        }

    }
}
