using CDEAutomation.classes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;


namespace CDEAutomation
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            


            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i] == "-c")
                    {
                        xmlHelper.location = args[i + 1];
                    }
                    if (args[i] == "-l")
                    {
                        log.location = args[i + 1];
                    }
                    if (args[i] == "-a")
                    {
                        appMethods.xlsLoc = args[i + 1];
                    }
                }
                appMethods.runCDEAuto();
            }
            else
            {
                Console.WriteLine("Missing arugments");
                Console.WriteLine(@"Usage: ""CDEAutomation.exe -c config.xml -l log.txt -a attributes.xlsx"" ");

                //log.write("Failed", "Missing appication arugments");

                //Console.ReadLine();
                //return;
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainUI());
            }
         
        }
}
}
