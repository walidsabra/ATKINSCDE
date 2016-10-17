using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using CDEAutomation.classes;
using System.Net.Mail;
using System.Net;

namespace CDEAutomation
{

    class backupCode
    {
        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode, CallingConvention = CallingConvention.StdCall)]
        public static extern int aaApi_SelectDocuments(IntPtr lpSelectInfo, IntPtr fpCallback, int lUserData);

        [StructLayout(LayoutKind.Sequential)]
        public struct AADOC_ITEM
        {
            public int lProjectId;
            public int lDocumentId;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct AADOCSELECT_ITEM
        {
            public uint ulFlags;
            //Specifies flags for the document select operation.
            public int lEnvironmentId;
            //Specifies the environment identifier of the document item.
            public int lProjectId;
            //Specifies the project identifier of the document item.
            public int lDocumentId;
            //Specifies the unique document identifier.
            public int lSetId;
            //Specifies the set identifier of the document item.
            public int lSetType;
            //Specifies the set type of the document item.
            public string lpctstrFileName;
            //Pointer to a null-terminated string specifying the document file name.
            public string lpctstrName;
            //Pointer to a null-terminated string specifying the document name.
            public string lpctstrDescription;
            //Pointer to a null-terminated string specifying the document description.
            public string lpctstrVersion;
            //Pointer to a null-terminated string specifying the document version string.
            public int lVersionSeq;
            //Specifies the version sequence of the document item.
            public int lOriginal;
            //Specifies the original number of the document item.
            public int lCreatorId;
            //Specifies the creator's user identifier.
            public int lUpdaterId;
            //Specifies the last updater user identifier.
            public int lLastUserId;
            //Specifies the identifier of the user who has currently checked out the document or if the document is not checked out this member specifies the identifier of the user who was the last one that had the document checked out.
            public string lpctstrStatus;
            //Pointer to a null-terminated string specifying the document status.
            public int lFileType;
            //Specifies the file type of the document.
            public int lItemType;
            //Specifies the item type of the document.
            public int lStorageId;
            //Specifies the document file's storage identifier.
            public int lWorkflowId;
            ////Specifies the identifier of a workflow associated with the document.
            public int lStateId;
            ////Specifies the document's state identifier.
            public int lApplicationId;
            ////Specifies the identifier of an application assigned to the document.
            public int lDepartmentId;
            ////Specifies the identifier of the department assigned to the document.
            public int lManagerId;
            public int lItemFlags;
            public int lManagerType;
            public IntPtr lpProject;
            public int lFileUpdator;

        }

        IntPtr AADMSBUFFER_DOCUMENT = IntPtr.Zero;

        public static int SelectDocuments(AADOCSELECT_ITEM lpSelectInfo)
        {
            try
            {
                IntPtr ptr = Marshal.AllocHGlobal(Marshal.SizeOf(lpSelectInfo));
                Marshal.StructureToPtr(lpSelectInfo, ptr, true);
                return aaApi_SelectDocuments(ptr, IntPtr.Zero, 0);
            }
            catch (Exception)
            {

                throw;
            }
        }

  //        <!--Environment-->
  //<Environment>
  //  <EnvName>Env1</EnvName>
  //  <TitleCol>xTitle</TitleCol>
  //  <PackageCol>yPackage</PackageCol>
  //  <FileNoCol>zFileNumber</FileNoCol>
  //  <RevisionCol>xRevision</RevisionCol>
  //  <RevNotesCol>yRevNote</RevNotesCol>
  //  <SuitabilityCol>zSus</SuitabilityCol>
  //  <DesignStageCol>xDesSta</DesignStageCol>
  //</Environment>

        /// <summary>
        /// Send Mail with processed documents 
        /// </summary>
        /// <param name="body"></param>
        public static void SendMail(List<emailColumns> docList)
        {
            string body;
            body = "The following files located in ProjectWise are issued for your action/information. Click on File Number to access files.";
            body = body + "<br/><br/>";
            body = body + @"<table border=""1"">" +
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


                            //int envSelected = PWMethods.aaApi_SelectEnvByTableId(9);
                    //int coId = PWMethods.aaApi_GetEnvNumericProperty(4, 0);
                    //Console.WriteLine("Col Id: " + coId);

         

                    //bool valid1 = PWMethods.aaApi_SetLinkDataColumnValue(10, 1, "Ttestttttttt");
                    //bool valid2 = PWMethods.aaApi_CreateLinkData(10, 9, 100000);

                    //int err = PWMethods.aaApi_GetLastErrorId();
                    //Console.WriteLine("Error Id: " +err);

                    //string colValue = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, 5));
                    //for (int i = 0; i < 9; i++)
                    //{
                    //    string colValue = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, i));
                    //    Console.WriteLine("Id is: {0} - Vaule is: {1}", i, colValue);
                    //}


                    //Console.WriteLine(colValue);
                    //bool valid4 = PWMethods.aaApi_UpdateLinkDataColumnValue(10, 5, "Test From Code");
                    //bool valid3 = PWMethods.aaApi_UpdateEnvAttr(10, 14);
                    //Console.WriteLine(colValue);

                    //bool valid2 = PWMethods.aaApi_CreateLinkData(10, 1, 1000000);

                    //Console.WriteLine("Set: " + valid1);
                    //Console.WriteLine("create: " + valid2);
                    //Console.WriteLine("Update: " + valid3);
                    //Console.WriteLine("Updated: " + valid4);


        //int lTableCount = PWMethods.aaApi_SelectAllTables();
        //for (int x = 0; x < lTableCount; x++)
        //{
        //    string tblName = Marshal.PtrToStringUni(PWMethods.aaApi_GetTableStringProperty(3, x));
        //    //aaApi_SelectEnv 
        //    //aaApi_GetEnvStringProperty
        //    //AADMSBUFFER_ENVIRONMENT
        //    if (tblName == EnvInfo.EnvironmentName)
        //    {
        //        tblId = PWMethods.aaApi_GetTableNumericProperty(1, x);
        //        int EnvCount = PWMethods.aaApi_SelectEnvByTableId(tblId);
        //        colCount = PWMethods.aaApi_SelectColumnsByTable(tblId);
        //        Console.WriteLine("TableName: {0} - TableId: {1} - Columns:{2}", tblName, tblId, colCount);
        //    }

        //}



        //if (!exists)
        //{
        //    //create folder
        //    System.IO.DirectoryInfo dir = new System.IO.DirectoryInfo(item.PDriveFolder);
        //    dir.Create();
        //}



        ////Get ColId 
        ////1 - SelectColByTabId
        ////2 - Get COl Numeric 
        //for (int z = 0; z < colCount; z++)
        //{
        //    string colName = Marshal.PtrToStringUni(PWMethods.aaApi_GetColumnStringProperty(9, z));
        //    int colId = PWMethods.aaApi_GetColumnNumericProperty(1, z);
        //    //Console.WriteLine("Name: {0} - Id: {1}", colName, colId);
        //    if (colName == EnvInfo.Title)
        //    {
        //        bool valid1 = PWMethods.aaApi_SetLinkDataColumnValue(tblId, colId, "Title testttttttt");
        //        if (valid1)
        //        {
        //            bool valid2 = PWMethods.aaApi_CreateLinkData(tblId, colId, 10000);
        //        }
        //    }
        //}

        ////aaApi_SetLinkDataColumnValue
        ////aaApi_CreateLinkData() 





        //Get Env Attributes 
        //GetFileAttributesforEmail(EnvInfo, PWFolderId, tblId, colCount, row, docId);
        //docList.Add(row);

        private static void GetFileAttributesforEmail(EnvObject EnvInfo, int PWFolderId, int tblId, int colCount, emailColumns row, int docId)
        {
            //get Env attributes
            int selectedRows = PWMethods.aaApi_SelectLinkDataByObject(tblId, 2, PWFolderId, docId, "", 0, 0, 0);

            int colId;

            if (selectedRows > 0)
            {
                //AADMSBUFFER_TABLE
                //AADMSBUFFER_ENVIRONMENT
                //aaApi_SelectEnvByProjectId
                //aaApi_SelectEnvByTableId
                //aaApi_GetEnvStringProperty
                for (int y = 0; y < colCount; y++)
                {
                    string colName = Marshal.PtrToStringUni(PWMethods.aaApi_GetColumnStringProperty(9, y));
                    colId = PWMethods.aaApi_GetColumnNumericProperty(1, y);

                    if (colName == EnvInfo.Title)
                    {
                        row.Title = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, y));
                    }
                    if (colName == EnvInfo.Package)
                    {
                        row.Package = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, y));
                    }
                    if (colName == EnvInfo.FileNumber)
                    {
                        row.FileNumber = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, y));
                    }
                    if (colName == EnvInfo.Revision)
                    {
                        row.Revision = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, y));
                    }
                    if (colName == EnvInfo.RevisionNote)
                    {
                        row.RevisionNote = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, y));
                    }
                    if (colName == EnvInfo.Suitability)
                    {
                        row.Suitability = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, y));
                    }
                    if (colName == EnvInfo.DesignStage)
                    {
                        row.DesignStage = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, y));
                    }
                    bool valid1 = PWMethods.aaApi_UpdateLinkData(tblId, colId, "Hi");
                    //Console.WriteLine("Update Col" + valid1);
                    if (valid1)
                    {
                        bool valid2 = PWMethods.aaApi_UpdateEnvAttr(tblId, colId);
                        //Console.WriteLine("Update Env" + valid2);
                    }
                }
            }
        }
    }
}
