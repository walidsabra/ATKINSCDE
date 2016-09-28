using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using CDEAutomation.classes;

namespace CDEAutomation
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("... CDE Automation Tool Started ...");


            dsInfo dsInfo = new dsInfo();
            dsInfo = xmlHelper.getDatasourceInfo();

            Console.WriteLine( "Connecting to datasource: " + dsInfo.dsName + " with user: " + dsInfo.dsUser);


            //Login to ProjectWise
            bool Initialized = PWMethods.aaApi_Initialize(0);
            bool validLogin = PWMethods.aaApi_Login(0, dsInfo.dsName, dsInfo.dsUser, dsInfo.dsPass, null);

            Console.WriteLine("ProjectWise Initialized?: " + Initialized);
            Console.WriteLine("Valid ProjectWise Login?: " + validLogin);

            // Mappings
            List<FolderMapping> folderList = new List<FolderMapping>();
            folderList = xmlHelper.getMappings();

            foreach (FolderMapping item in folderList)
            {
                string stagingPath = item.stagingfolder;
                string pwPath = item.ProjectWiseFolder;
                string pPath = item.PDriveFolder;
                

                Console.WriteLine("Processing Folder: " + stagingPath + " to ProjectWise Folder: " + pwPath);


                DirectoryInfo d = new DirectoryInfo(item.stagingfolder);

                //Select Folder By Name
                int countFolders = PWMethods.aaApi_SelectProjectsByProp(null, item.ProjectWiseFolder, null, null);

                //if folder doesn't exist on ProjectWise
                if (countFolders < 1)
                {
                    Console.WriteLine(item.ProjectWiseFolder + " - Mapped folder doesn't Exists on ProjectWise, Please check you mappings");
                    Console.ReadLine();

                    //check if list of records stored in memory to send noti is sent here 
                    return;
                }


                //Get Folder Id
                int PWFolderId = PWMethods.aaApi_GetProjectNumericProperty(1, 0);
                //Console.WriteLine(PWFolderId);

                //Select files (all general attributes) within PW Folder
                int selected = PWMethods.aaApi_SelectDocumentsByProjectId(PWFolderId);
                Console.WriteLine("Document in Folder?: " + selected);
            
                //Loop on file in Staging Folder  sorted by Size Ascending 
                foreach (var file in d.GetFiles("*.*").OrderBy(f => f.Length))
                {
                    String ofilePath = file.FullName;
                    String ofileName = file.Name;
                    String oName = Path.GetFileNameWithoutExtension(file.FullName);
                    

                    //if File already Exist on ProjectWise Folder - Create Version
                    if (selected > 0)
                    {
                        for (int i = 0; i < selected; i++)
                        {
                            int docSequence = PWMethods.aaApi_GetDocumentNumericProperty(2, i);
                             int docId = PWMethods.aaApi_GetDocumentNumericProperty(1, i);
                             bool isActive = true; //check if Active Version

                            string documentFileName = Marshal.PtrToStringUni(PWMethods.aaApi_GetDocumentStringProperty(21, i));
                            string docVersion = Marshal.PtrToStringUni(PWMethods.aaApi_GetDocumentStringProperty(23, i));

                            if (documentFileName == ofileName)  //if exists in PW
                            {
                                if (isActive) 
                                {
                                    Console.WriteLine("Processing File: " + documentFileName + " Sequence: " + docSequence + " Version: " + docVersion );
                                    Console.WriteLine("File already exists on ProjectWise");

                                    //Change PW file State to Archived
                                    appMethods.ChangeWFState(docId, PWFolderId, 2, "Changed to Archived by CDE Automation");

                                    //Create Revision
                                    appMethods.CreateDocRevision(PWFolderId, ofilePath, docId);

                                    //Change PW file State to Shared
                                    appMethods.ChangeWFState(docId, PWFolderId, 1, "Changed to Shared by CDE Automation");
                                }
                            }
                            else
                            {
                                appMethods.CreatPWDoc(PWFolderId, ofilePath, ofileName, oName);
                               
                            }
                        }
                    }
                    else
                    {
                       appMethods.CreatPWDoc(PWFolderId, ofilePath, ofileName, oName);
                    }

                    //update attributes 

                    otherConfig myconfig =  xmlHelper.getAppConfigs();

                    // copy to P drive: - Overright
                    if (myconfig.copytoP == "true")
                    {
                        bool exists = System.IO.Directory.Exists(item.PDriveFolder);
                        if (!exists)
                        {
                            //create folder
                            System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(item.PDriveFolder);
                            dir.Create();
                        }
                        Console.WriteLine("Copying " + ofileName + " to folder: " + item.PDriveFolder);
                        System.IO.File.Copy(ofilePath, Path.Combine( item.PDriveFolder, ofileName), true);
                    }

                    // Notify Users 
                    if (myconfig.SendNotification == "true")
                    {
                        Console.WriteLine("Ready to Send ....");
                    }

                    //Delete/Move File/s from Folder in here or outside the loop
                    //?
                }
            }


            Console.ReadLine();
        }

    }
}



//int ErrorId = aaApi_GetLastErrorId(0);
//aaApi_ExecuteDocumentCommand