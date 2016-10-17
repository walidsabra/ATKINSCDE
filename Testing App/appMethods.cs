using CDEAutomation.classes;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;

namespace CDEAutomation
{
    class appMethods
    {

        public static int GetDocIdByFolderId(int PWFolderId, String oName)
        {
            int docId = 0;
            int select = PWMethods.aaApi_SelectDocumentsByProjectId(PWFolderId);
            for (int v = 0; v < select; v++)
            {
                string cName = Marshal.PtrToStringUni(PWMethods.aaApi_GetDocumentStringProperty(20, v));

                if (cName == oName)
                {
                    //get docId 
                    docId = PWMethods.aaApi_GetDocumentNumericProperty(1, v);
                }
            }
            return docId;
        }

        /// <summary>
        /// Get Columns ID, Labels By Environment Id 
        /// </summary>
        /// <param name="oEnvId"></param>
        /// <param name="ColumnsList"></param>
        public static void GetColumnsByEnvID(int oEnvId, List<PWAttr> ColumnsList)
        {
            int intrAttrCount = PWMethods.aaApi_SelectEnvAttrGuiDefs(oEnvId, -1, -1, -1, -1);
            if (intrAttrCount > 0)
            {
                for (int m = 0; m < intrAttrCount; m++)
                {
                    //AADMSBUFFER_EATTRGUI - 17 = Label Text ~ 3 = ColId ~ 2 = TableID
                    PWAttr Column = new PWAttr();
                    Column.ColumnId = PWMethods.aaApi_GetEnvAttrGuiDefNumericProperty(3, m);
                    Column.TableId = PWMethods.aaApi_GetEnvAttrGuiDefNumericProperty(2, m);
                    Column.LabelText = Marshal.PtrToStringUni(PWMethods.aaApi_GetEnvAttrGuiDefStringProperty(17, m));
                    ColumnsList.Add(Column);
                }
            }
        }

        /// <summary>
        /// Send Mail with processed documents 
        /// </summary>
        /// <param name="body"></param>
        public static void SendMail2(List<emailColumn> docList, int NuCols)
        {
            string body;
            body = "The following files located in ProjectWise are issued for your action/information. Click on File Number to access files.";
            body = body + "<br/><br/>";
            body = body + @"<table border=""1"">" +
                            @"<tr style=""background-color:blue; color:white"">";
            int g = 0;
                            foreach (emailColumn col in docList)
	                        {
                                g++;
                                body = body + "<th>" + col.ColumnName + "</th>";
                                if (g == NuCols)
                                {
                                    break;
                                }
	                        }
            body = body + "</tr>";

            int h = 0;
            foreach (emailColumn item in docList)
            {
                if (h==0)
                {
                    body = body + "<tr>";
                }
                if (item.ColumnName == "Name")
                {
                    body = body + "<td>" + @"<a href=""" + item.URL + @""">" + item.ColumnVaule + "</a> </td>";
                }
                else
                {
                    body = body + "<td>" + item.ColumnVaule + "</td>";
                }
                
                if (h==NuCols -1)
                {
                    body = body + "</tr>";
                }
                h++;
                if (h == NuCols)
                {
                    h = 0;
                }
            }

            body = body + "</table>";
            body = body + "<br/><br/>";
            body = body + "Note: These files are also replicated to Z: drive.";

            Console.WriteLine(body);

            Console.WriteLine("Sending email...");

            oNotification msg = xmlHelper.getMsgConfigs();
            msg.Body = body;

            //MailAddress to = new MailAddress(msg.To);
            MailAddress from = new MailAddress(msg.From);
            MailMessage mail = new MailMessage();
            mail.From = from;
            mail.To.Add(msg.To);  //add array of addresses and loop to add
            mail.Subject = msg.Subject;
            mail.Body = msg.Body;
            mail.IsBodyHtml = true;



            SmtpClient smtp = new SmtpClient();
            smtp.Host = msg.host;
            smtp.Port = msg.port;

            smtp.Credentials = new NetworkCredential(msg.user, msg.password);
            smtp.EnableSsl = msg.ssl;

            try
            {
                smtp.Send(mail);
                log.write("Success", "Email Sent");
            }
            catch (Exception)
            {

                log.write("Failed", "Couldn't send email");
            }

        }

        /// <summary>
        /// Send Mail with processed documents 
        /// </summary>
        /// <param name="body"></param>
        public static void SendMail(List<emailColumns> docList)
        {
            string body;
            body = "The following files located in ProjectWise are issued for your action/information. Click on File Number to access files.";
            body = body + "<br/><br/>";
            body = body +  @"<table border=""1"">" +
                            @"<tr style=""background-color:blue; color:white"">" +
                                "<th>Package</th>" +
                                "<th>File Number</th>" +
                                "<th>Revision</th>" +
                                "<th>Title</th>" +
                                "<th>Revision Note</th>" +
                                "<th>Suitability</th>" +
                                "<th>Design Stage</th>" +
                                "<th>File Updated</th>" +
                                "<th>File Updated By</th>" +
                            "</tr>";

            foreach (emailColumns item in docList)
            {
                string row = 
                            "<tr>" +
                                "<td>" + item.Package + "</td>" +
                                "<td>" + @"<a href=""" + item.URL + @""">" + item.FileNumber + "</a> </td>" + 
                                "<td>" + item.Revision + "</td>" +
                                "<td>" + item.Title + "</td>" +
                                "<td>" + item.RevisionNote + "</td>" +
                                "<td>" + item.Suitability + "</td>" +
                                "<td>" + item.DesignStage + "</td>" + 
                                "<td>" + item.FileUpdated + "</td>" +
                                "<td>" + item.FileUpdatedBy + "</td>" +
                            "</tr>";
                body = body + row;
            }

             body = body + "</table>";
             body = body + "<br/><br/>";
             body = body + "Note: These files are also replicated to Z: drive.";

            //Console.WriteLine(body);

            Console.WriteLine("Sending email...");

            oNotification msg = xmlHelper.getMsgConfigs();
            msg.Body = body;

            //MailAddress to = new MailAddress(msg.To);
            MailAddress from = new MailAddress(msg.From);
            MailMessage mail = new MailMessage();
            mail.From = from;
            mail.To.Add(msg.To);  //add array of addresses and loop to add
            mail.Subject = msg.Subject;
            mail.Body = msg.Body;
            mail.IsBodyHtml = true;
            


            SmtpClient smtp = new SmtpClient();
            smtp.Host = msg.host;
            smtp.Port = msg.port;

            smtp.Credentials = new NetworkCredential(msg.user, msg.password);
            smtp.EnableSsl = msg.ssl;

            try
            {
                smtp.Send(mail);
                log.write("Success", "Email Sent");
            }
            catch (Exception)
            {

                log.write("Failed", "Couldn't send email");
            }
          
        }

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
                log.write("Success", fileName + " Created Sucessfully on ProjectWise");

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

                        //Update Attributes

                        //Add Row
                    }
                }


            }
            else
            {
                //Document was not created - error?
                log.write("Error", "Error Creating document");
            }
        }


    }
}
