using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace Testing_App
{
    class Program
    {
        [DllImport("PWSDKMethods.dll")]
        public static extern bool CheckIn(int ProjectID, int DocumentID);

        static void Main(string[] args)
        {
            Console.WriteLine("Good Morning ... Let's Test the App");
            Console.WriteLine("Enter your ProjectWise User Name");
            var UserName = Console.ReadLine();

            Console.WriteLine("Enter your Password");
            var Password = Console.ReadLine();

            Console.WriteLine("User Name Is: " + UserName  +  " And password is: " +  Password);
            bool Test  = CheckIn(50, 1);
            Console.ReadLine();
        }
    }
}
