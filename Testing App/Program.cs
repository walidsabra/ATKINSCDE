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
        static void Main(string[] args)
        {
            xmlHelper.readxml();
            Console.WriteLine("... Let's Test the App ...");
            Console.WriteLine("Enter your ProjectWise User Name");
            var UserName = Console.ReadLine();

            Console.WriteLine("Enter your Password");
            var Password = Console.ReadLine();

            Console.WriteLine("User Name Is: " + UserName  +  " And password is: " +  Password);
            //bool Test = CreatePWFile();

            //Login to ProjectWise
            bool Initialized = PWMethods.aaApi_Initialize(0);
            bool validLogin = PWMethods.aaApi_Login(0, "DCH CDE", "admin", "admin", null);

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
            int selected = PWMethods.aaApi_SelectDocumentsByProjectId(PWFolderId);
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
                    int docSequence = PWMethods.aaApi_GetDocumentNumericProperty(2, i);
                    int docId = PWMethods.aaApi_GetDocumentNumericProperty(1, i);

                    string documentFileName = Marshal.PtrToStringUni(PWMethods.aaApi_GetDocumentStringProperty(21, i));
                    string docVersion = Marshal.PtrToStringUni(PWMethods.aaApi_GetDocumentStringProperty(23, i));

                    if (documentFileName == ofileName)
                    {
                        Console.WriteLine("File: " + documentFileName +" Sequence: " + docSequence + " Version: " + docVersion+ " Already Exists in ProjectWise");
                        //Change PW file State to Archived

                        //creat new Revision here
                        PWMethods.aaApi_NewDocumentVersion(0, PWFolderId, docId, 0,"Version created by automation tool",0);
                        
                        //Change Document File
                        PWMethods.aaApi_ChangeDocumentFile(PWFolderId, docId, ofilePath);
                    }
                    else
                    {
                        //Else - Create New 
                        bool test4 = PWMethods.aaApi_CreateDocument(0, PWFolderId, 0, 0, 0, 0, 0, 0, ofilePath, ofileName, oName, null, null, false, 0, null, 1000, 0);
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
