using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using CDEAutomation.classes;

namespace CDEAutomation
{
    class appMethods
    {

        /// <summary>
        /// Change Document WF State
        /// </summary>
        /// <param name="docId"></param>
        /// <param name="PWFolderId"></param>
        /// <param name="myWF"></param>
        public static void ChangeWFState(int docId, int PWFolderId, int state, string comment)
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
        public static void CreateDocRevision(int folderId, String filePath, int docId)
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
        public static void CreatPWDoc(int folderId, String filePath, String fileName, String docName)
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
