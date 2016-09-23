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




        
        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectDocumentsByProjectId(int lProjectId);

                
        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr aaApi_GetDocumentStringProperty(int lPropertyId, int index);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        private static extern int aaApi_GetDocumentNumericProperty(int lPropertyId, int index);




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
            bool Initialized = aaApi_Initialize(0);
            bool validLogin = aaApi_Login(0, "DCH CDE", "admin", "admin", null);

            Console.WriteLine("Initialized:" + Initialized);
            Console.WriteLine("validLogin" + validLogin);

            //Select Folder By Name

            //Get Folder Id
            //use her the API to get the ID
            int PWFolderId = 282;

            // Msppings

            string filepath = "D:\\PWTEST";
            DirectoryInfo d = new DirectoryInfo(filepath);

            //Select files (all general attributes) within PW Folder
            int selected = aaApi_SelectDocumentsByProjectId(PWFolderId);
            Console.WriteLine("Selected?: " + selected);

            //Loop on file in Staging Folder 
            foreach (var file in d.GetFiles("*.docx"))
            {
                String ofilePath = file.FullName;
                String ofileName = file.Name;
                String oName = Path.GetFileNameWithoutExtension(file.FullName);


                //if File already Exist on ProjectWise Folder - Create Version

                for (int i = 0; i < selected; i++)
                {
                    string documentFileName = Marshal.PtrToStringUni(aaApi_GetDocumentStringProperty(21, i));
                    int docSequence = aaApi_GetDocumentNumericProperty(2, i);
                    string docVersion = Marshal.PtrToStringUni(aaApi_GetDocumentStringProperty(23, i));

                    if (documentFileName == ofileName)
                    {
                        Console.WriteLine("File: " + documentFileName +" Sequence: " + docSequence + " Version: " + docVersion+ " Already Exists in ProjectWise");
                        //Change PW file State to Archived

                        //creat new Revision here
                    }
                    else
                    {
                        //Else - Create New 
                        bool test4 = aaApi_CreateDocument(0, PWFolderId, 0, 0, 0, 0, 0, 0, ofilePath, ofileName, oName, null, null, false, 0, null, 1000, 0);
                        //bool test4 = aaApi_CreateDocument(0, 282, 0, 0, 0, 0, 0, 0, "D:\\PWTEST\\PWSDK.docx", "PWSDK.docx", "Test 4 from App", null, null, false, 0, null, 1000, 0);
                        if (test4)
                        {
                            Console.WriteLine(ofileName + " Created Sucessfully on ProjectWise");

                        }
    
                    }
                    //change WF state to Shared

                    // copy to P drive: - Overright

                    // Notify Users 

                    //Delete/Move File/s from Folder in here or outside the loop
                }


                

               
            }

            
            
            //int ErrorId = aaApi_GetLastErrorId(0);

            
    

            //aaApi_ExecuteDocumentCommand



            Console.ReadLine();



        }
    }
}
