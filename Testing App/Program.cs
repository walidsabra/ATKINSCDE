using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;

namespace Testing_App
{
    class Program
    {
        //[DllImport("PWSDKMethods.dll")]
        //public static extern bool CreatePWFile();
        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll")]
        public static extern bool aaApi_Initialize(uint ulModule);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll")]
        public static extern int aaApi_GetLastErrorId(IntPtr A);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern bool aaApi_Login(int IDsType, String DSName, String UserName, String password, String schema);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern bool aaApi_CreateDocument(
            int plDocumentId, 
            int lProjectId, 
            int lStorageId, 
            int lFileType, 
            int lItemType, 
            int lApplicationId, 
            int lDepartmentId, 
            int lWorkspaceProfileId, 
            String lpctstrSourceFile, 
            String lpctstrFileName, 
            String lpctstrName, 
            String lpctstrDesc, 
            String lpctstrVersion, 
            bool bLeaveOut, 
            uint ulFlags, 
            StringBuilder lptstrWorkingFile, 
            int lBufferSize, 
            int plAttributeId);


        [StructLayout(LayoutKind.Sequential)]
        public struct AADOC_ITEM
        {
            public int lProjectId;
            public int lDocumentId;
        }


        static void Main(string[] args)
        {
            Console.WriteLine("... Let's Test the App ...");
            Console.WriteLine("Enter your ProjectWise User Name");
            var UserName = Console.ReadLine();

            Console.WriteLine("Enter your Password");
            var Password = Console.ReadLine();

            Console.WriteLine("User Name Is: " + UserName  +  " And password is: " +  Password);
            //bool Test = CreatePWFile();

            //Login to ProjectWise
            bool test2 = aaApi_Initialize(0);
            bool test3 = aaApi_Login(0, "DCH CDE", "admin", "admin", null);

            Console.WriteLine(test2);
            Console.WriteLine(test3);

            string filepath = "D:\\PWTEST";
            DirectoryInfo d = new DirectoryInfo(filepath);

            //Loop on file in Folder 
            foreach (var file in d.GetFiles("*.docx"))
            {
                String ofilePath = file.FullName;
                String ofileName = file.Name;
                String oName = Path.GetFileNameWithoutExtension(file.FullName);

                //if File already Exist on ProjectWise Folder - Create Version


                //Else - Create New 
                bool test4 = aaApi_CreateDocument(0, 282, 0, 0, 0, 0, 0, 0, ofilePath, ofileName, oName, null, null, false, 0, null, 1000, 0);
                //bool test4 = aaApi_CreateDocument(0, 282, 0, 0, 0, 0, 0, 0, "D:\\PWTEST\\PWSDK.docx", "PWSDK.docx", "Test 4 from App", null, null, false, 0, null, 1000, 0);
                Console.WriteLine(test4);

                //Delete File/s from Folder in here or outside the loop
            }

            
            
            //int ErrorId = aaApi_GetLastErrorId(0);

            
    

            //aaApi_ExecuteDocumentCommand



            Console.ReadLine();



        }
    }
}
