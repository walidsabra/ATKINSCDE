using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CDEAutomation
{
    class log
    {
        public static string location;

        public static void write(string status, string message)
        {
            StreamWriter log;
            try
            {
                if (!File.Exists(location))
                {
                    log = new StreamWriter(location);
                }
                else
                {
                    log = File.AppendText(location);
                }
                string line = status + " - " + message;
                log.WriteLine(DateTime.Now);
                log.WriteLine(line);
                log.WriteLine();

                log.Close();
            }
            catch (Exception)
            {
                

            }


        }
    }
}
