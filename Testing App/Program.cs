using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.IO;
using Testing_App.classes;

namespace Testing_App
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
                //Console.WriteLine(countFolders);

                //Get Folder Id
                int PWFolderId = PWMethods.aaApi_GetProjectNumericProperty(1, 0);
                //Console.WriteLine(PWFolderId);

                //Select files (all general attributes) within PW Folder
                int selected = PWMethods.aaApi_SelectDocumentsByProjectId(PWFolderId);
                Console.WriteLine("Document in Folder?: " + selected);
            
                //Loop on file in Staging Folder 
                foreach (var file in d.GetFiles("*.*"))
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
                             bool isActive = true;

                            string documentFileName = Marshal.PtrToStringUni(PWMethods.aaApi_GetDocumentStringProperty(21, i));
                            string docVersion = Marshal.PtrToStringUni(PWMethods.aaApi_GetDocumentStringProperty(23, i));

                            if (documentFileName == ofileName)  //if exists in PW
                            {
                                if (isActive) //if Active Version
                                {
                                    Console.WriteLine("Processing File: " + documentFileName + " Sequence: " + docSequence + " Version: " + docVersion + " Already Exists in ProjectWise");

                                    //Change PW file State to Archived
                                    ChangeWFState(docId, PWFolderId, 2, "Changed to Archived by CDE Automation");

                                    //Create Revision
                                    CreateDocRevision(PWFolderId, ofilePath, docId);

                                    //Change PW file State to Shared
                                    ChangeWFState(docId, PWFolderId, 1, "Changed to Shared by CDE Automation");
                                }
                            }
                            else
                            {
                                CreatPWDoc(PWFolderId, ofilePath, ofileName, oName);
                               
                            }
                        }
                    }
                    else
                    {
                        CreatPWDoc(PWFolderId, ofilePath, ofileName, oName);
                    }

                    otherConfig myconfig =  xmlHelper.getAppConfigs();


                   
                    
                    // copy to P drive: - Overright
                    if (myconfig.copytoP == "True")
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
                    if (myconfig.SendNotification == "True")
                    {
                        Console.WriteLine("Ready to Send ....");
                    }

                    //Delete/Move File/s from Folder in here or outside the loop
                    //?
                }
            }


            Console.ReadLine();
        }

        /// <summary>
        /// Change Document WF State
        /// </summary>
        /// <param name="docId"></param>
        /// <param name="PWFolderId"></param>
        /// <param name="myWF"></param>
        private static void ChangeWFState(int docId, int PWFolderId,  int state, string comment)
        {
            wfObject myWF = new wfObject(); //set an object to be used in WF state change
            wfObject WFObj = xmlHelper.getWFConfigs(); //get what stored in xml

            int wfCount = PWMethods.aaApi_SelectWorkflowsForProject(PWFolderId);
            //Console.WriteLine("Number of WF: " + wfCount);

            for (int i = 0; i < wfCount; i++)
            {
                myWF.WorkflowName = Marshal.PtrToStringUni(PWMethods.aaApi_GetWorkflowStringProperty(3, i));
                if (myWF.WorkflowName == WFObj.WorkflowName)
                {
                    myWF.WorkflowId = PWMethods.aaApi_GetWorkflowNumericProperty(1, 0);
                    //Console.WriteLine("WF Id is: " + myWF.WorkflowId);
                }

            }

            int stCount = PWMethods.aaApi_SelectStatesByWorkflow(myWF.WorkflowId);
            //Console.WriteLine("Number of States is: " + stCount);
            
            
            for (int i = 0; i < stCount; i++)
            {
                myWF.StateName = Marshal.PtrToStringUni(PWMethods.aaApi_GetStateStringProperty(2, i));
                if (myWF.StateName == WFObj.SharedStateName)
                {
                    myWF.StateOneId = PWMethods.aaApi_GetStateNumericProperty(1, i);
                }
                if (myWF.StateName == WFObj.ArchivedStateName)
                {
                    myWF.StateTwoId = PWMethods.aaApi_GetStateNumericProperty(1, i);
                }
            }

            //Console.WriteLine("Pending Approval states id is: " + myWF.StateOneId);
            //Console.WriteLine("Approved  states id is: " + myWF.StateTwoId);

            if (state == 1)
            {
                bool stateChanged = PWMethods.aaApi_SetDocumentState(0, PWFolderId, docId, myWF.WorkflowId, myWF.StateOneId, comment);
            }
            if (state == 2)
            {
                bool stateChanged = PWMethods.aaApi_SetDocumentState(0, PWFolderId, docId, myWF.WorkflowId, myWF.StateTwoId, comment);
            }
            
        }

        /// <summary>
        /// Create New Version then Change the FIle
        /// </summary>
        /// <param name="folderId"></param>
        /// <param name="filePath"></param>
        /// <param name="docId"></param>
        private static void CreateDocRevision(int folderId, String filePath, int docId)
        {
            //creat new Version here
            PWMethods.aaApi_NewDocumentVersion(0, folderId, docId, 0, "Version created by automation tool", 0);

            //Change Document File
            PWMethods.aaApi_ChangeDocumentFile(folderId, docId, filePath);
        }


        /// <summary>
        /// Create ProjectWise Document Using Folder ID and basic Details
        /// </summary>
        /// <param name="folderId"></param>
        /// <param name="filePath"></param>
        /// <param name="fileName"></param>
        /// <param name="docName"></param>
        private static void CreatPWDoc(int folderId, String filePath, String fileName, String docName)
        {
            //Creat in PW
            bool docCreated = PWMethods.aaApi_CreateDocument(0, folderId, 0, 0, 0, 0, 0, 0, filePath, fileName, docName, null, null, false, 0, null, 1000, 0);
            if (docCreated)
            {
                Console.WriteLine(fileName + " Created Sucessfully on ProjectWise");

                //select created document to update WF state
                int selected = PWMethods.aaApi_SelectDocumentsByProjectId(folderId);
                for (int i = 0; i < selected; i++)
                {                    
                    string cName = Marshal.PtrToStringUni(PWMethods.aaApi_GetDocumentStringProperty(20, i));

                    if (cName == docName)
                    {
                        //get docId 
                        int docId = PWMethods.aaApi_GetDocumentNumericProperty(1, i);

                        //change WF state to Shared
                        ChangeWFState(docId, folderId, 1, "Changed to Shared by CDE Automation");
                    }
                }
                

            }
            else
            {
                //Document was not created - error?
            }
        }


    }
}



//int ErrorId = aaApi_GetLastErrorId(0);
//aaApi_ExecuteDocumentCommand