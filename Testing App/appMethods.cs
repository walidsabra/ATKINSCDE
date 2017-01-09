using CDEAutomation.classes;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Linq;


namespace CDEAutomation
{
    class appMethods
    {
        public static string xlsLoc = string.Empty;
        public static void runCDEAuto()
        {
            Console.WriteLine("... CDE Automation Tool Started ...");
            log.write("", "CDE Automation Tool Started ");


            List<emailColumns> docList = new List<emailColumns>();
            List<emailColumn> emlColumns = new List<emailColumn>();

            List<string> columns = new List<string>();
            columns = xmlHelper.getEmailColumns();

            dsInfo dsInfo = new dsInfo();
            dsInfo = xmlHelper.getDatasourceInfo();

            //EnvObject EnvInfo = new EnvObject();
            //EnvInfo = xmlHelper.getEnvConfigs();

            otherConfig myconfig = xmlHelper.getAppConfigs();

            List<FolderMapping> folderList = new List<FolderMapping>();
            folderList = xmlHelper.getMappings();

            Console.WriteLine("Connecting to datasource: " + dsInfo.dsName + " with user: " + dsInfo.dsUser);
            log.write("", "Connecting to datasource: " + dsInfo.dsName + " with user: " + dsInfo.dsUser);


            //Login to ProjectWise
            bool Initialized = PWMethods.aaApi_Initialize(0);
            bool validLogin = PWMethods.aaApi_Login(0, dsInfo.dsName, dsInfo.dsUser, dsInfo.dsPass, null);

            Console.WriteLine("ProjectWise Initialized?: " + Initialized);
            Console.WriteLine("Valid ProjectWise Login?: " + validLogin);
            if (validLogin)
            {
                log.write("Success", "Logged in sucessfully to ProjectWise datasource");
            }
            else
            {
                log.write("Failed", "Couldn't login to ProjectWise");
                return;
            }


            // Mappings
            foreach (FolderMapping item in folderList)
            {
                string stagingPath = item.stagingfolder;
                string pwPath = item.ProjectWiseFolder;
                string pPath = item.PDriveFolder;

                Console.WriteLine("Processing Folder: " + stagingPath + " to ProjectWise Folder: " + pwPath);
                log.write("", "Processing Folder: " + stagingPath + " to ProjectWise Folder: " + pwPath);

                DirectoryInfo d = new DirectoryInfo(item.stagingfolder);

                //Get Folder Id
                int PWFolderId = 0;
                int oEnvId = 0;
                int TableId = 0;
                //int colCount = 0;

                string pathStart = "Documents";
                item.ProjectWiseFolder = item.ProjectWiseFolder.Substring(item.ProjectWiseFolder.IndexOf(pathStart) + pathStart.Length);
                string[] pathFolders;
                string[] separators = { @"\" };
                pathFolders = item.ProjectWiseFolder.Split(separators, StringSplitOptions.RemoveEmptyEntries);

                int folderId = PWMethods.aaApi_GetProjectIdByNamePath(item.ProjectWiseFolder);
                //Select Folder By Name
                int xFolders = PWMethods.aaApi_SelectProjectsByProp(null, pathFolders.Last(), null, null);

                //is there is PW folder with this name?
                if (xFolders == 0)
                {
                    Console.WriteLine("Error - Folder doesn't exist in ProjectWise");
                    log.write("Error", "Folder doesn't exist in ProjectWise");
                    continue;
                }

               
                //are PW States Valide for this Folder?
                bool statesValidated = appMethods.ValidateStates(folderId);
                if (!statesValidated)
                {
                    Console.WriteLine("Error - Couldn't Select State; check your configuration File ");
                    log.write("Error", "Couldn't Select State; check your configuration File");
                    continue;
                }

                //if staging folder doesn't exists 
                bool StagingExists = System.IO.Directory.Exists(item.stagingfolder);
                if (!StagingExists)
                {
                    Console.WriteLine("Error - Staging Folder doesn't exist");
                    log.write("Error", "Staging Folder doesn't exist");
                    continue;
                }

                //if folder doesn't exists on Z Drive
                bool exists = System.IO.Directory.Exists(item.PDriveFolder);
                if (myconfig.CopytoFileSystem && !exists)
                {
                    Console.WriteLine("Error - Folder doesn't exist in FileSystem Location");
                    log.write("Error", "Folder doesn't exist in FileSystem Location");
                    continue;
                }

                //check PW & Z drive permissions 
                string myguid = Guid.NewGuid().ToString();
                string textFile = Path.Combine(item.PDriveFolder, myguid + ".tmp");
                // perhaps check File.Exists(file), but it would be a long-shot...
                bool canCreate;
                //check can create doc in PW?
                bool canCreatePW = PWMethods.aaApi_CreateDocument(0, folderId, 0, 0, 0, 0, 0, 0, textFile, myguid + ".tmp", myguid, null, null, false, 0, null, 1000, 0);
                if (!canCreatePW)
                {
                    Console.WriteLine("Error - User doesn't have the right permissions to this ProjectWise folder");
                    log.write("Error", "User doesn't have the right permissions to this ProjectWise folder");
                    continue;
                }
                else
                {
                    //delete file here from PW
                }

                try
                {
                    using (File.Create(textFile)) { }
                    File.Delete(textFile);
                    canCreate = true;
                }
                catch
                {
                    canCreate = false;
                }
                if (!canCreate)
                {
                    Console.WriteLine("Error - User doesn't have the right permissions to this folder");
                    log.write("Error", "User doesn't have the right permissions to this folder");
                    continue;
                }


                List<PWAttr> ColumnsList = new List<PWAttr>();
                for (int f = 0; f < xFolders; f++)
                {
                    int oLastFolderId = PWMethods.aaApi_GetProjectNumericProperty(1, f);
                    if (folderId == oLastFolderId)
                    {
                        PWFolderId = PWMethods.aaApi_GetProjectNumericProperty(1, f);
                        oEnvId = PWMethods.aaApi_GetProjectNumericProperty(21, f);
                        int envCount = PWMethods.aaApi_SelectEnv(oEnvId);

                        TableId = PWMethods.aaApi_GetEnvNumericProperty(2, 0);

                        Console.WriteLine("Folder Id: {0}, Env Id: {1}, TableId: {2}", PWFolderId, oEnvId, TableId);
                        log.write("Info", "Folder Id: " + PWFolderId + " Env Id: " + oEnvId + " TableId: " + TableId);

                        appMethods.GetColumnsByEnvID(oEnvId, ColumnsList);
                    }
                }




                //Loop on file in Staging Folder  sorted by Size Ascending 
                foreach (var file in d.GetFiles("*.*").OrderBy(f => f.Length))
                {
                    log.write("File Name1", file.Name);
                    //Select files (all general attributes) within PW Folder
                    int selected = PWMethods.aaApi_SelectDocumentsByProjectId(PWFolderId);
                    Console.WriteLine("Document in Folder?: " + selected);
                    log.write("", "Document in Folder?: " + selected);

                    try
                    {
                        string ofilePath = file.FullName;
                        string ofileName = file.Name;
                        string oName = Path.GetFileNameWithoutExtension(file.FullName);
                        int docId = 0;

                        log.write("File Name2", file.Name);

                        //loop on PW Select Files within folder to check if File already Exist - Create Version
                        if (selected > 0)
                        {
                            bool created = false;

                            for (int i = 0; i < selected; i++)
                            {
                                string documentFileName = string.Empty;
                                try
                                {
                                    documentFileName = Marshal.PtrToStringUni(PWMethods.aaApi_GetDocumentStringProperty(21, i));
                                }
                                catch (Exception)
                                {
                                    documentFileName = string.Empty;


                                }
                                if (!string.IsNullOrEmpty(documentFileName) && documentFileName.ToLower() == ofileName.ToLower())
                                {
                                    docId = PWMethods.aaApi_GetDocumentNumericProperty(1, i);

                                    Console.WriteLine("Processing Existing File: " + documentFileName);
                                    log.write("", "Processing Existing File: " + documentFileName);

                                    //Change PW file State to Archived
                                    appMethods.ChangeWFState(docId, PWFolderId, 2, "Changed to Archived by CDE Automation");

                                    //Create Revision
                                    appMethods.CreateDocRevision(PWFolderId, ofilePath, docId);

                                    //Change PW file State to Shared
                                    appMethods.ChangeWFState(docId, PWFolderId, 1, "Changed to Shared by CDE Automation");
                                    created = true;
                                    break;
                                }

                            }
                            if (!created)
                            {
                                Console.WriteLine("Processing New File: " + ofileName);
                                log.write("Info", "Processing New File: " + ofileName);

                                appMethods.CreatPWDoc(PWFolderId, ofilePath, ofileName, oName);
                                docId = appMethods.GetDocIdByFolderId(PWFolderId, oName);

                            }
                        }
                        else
                        {
                            Console.WriteLine("Processing New File: " + ofileName);
                            log.write("Info", "Processing New File: " + ofileName);

                            appMethods.CreatPWDoc(PWFolderId, ofilePath, ofileName, oName);
                            docId = appMethods.GetDocIdByFolderId(PWFolderId, oName);
                        }



                        //Update Attribute from xls
                        int colCount = PWMethods.aaApi_SelectColumnsByTable(TableId);
                        int selectedRows = PWMethods.aaApi_SelectLinkDataByObject(TableId, 2, PWFolderId, docId, "", 0, 0, 0);
                        Console.WriteLine("Link rows selected: " + selectedRows);
                    rowExists:
                        //Case document already has attributes in Env
                        if (selectedRows > 0)
                        {
                            string attrnoValue = string.Empty;
                            for (int i = 0; i < colCount; i++)
                            {
                                string colName = Marshal.PtrToStringUni(PWMethods.aaApi_GetColumnStringProperty(9, i));
                                //Console.WriteLine(colName);
                                if (colName == "a_attrno")
                                {
                                    attrnoValue = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, i));
                                    Console.WriteLine("a_attrno:" + attrnoValue);
                                }
                            }

                            if (xlsLoc != string.Empty)
                            {
                                try
                                {
                                    //read xls file
                                    var excel = new LinqToExcel.ExcelQueryFactory();
                                    excel.FileName = xlsLoc;
                                    var docAttr = from x in excel.Worksheet()
                                                  select x;
                                    foreach (var xlrow in docAttr)
                                    {
                                        string NameInXls = xlrow["Name"].Value.ToString().ToUpper();
                                        string NameInPW = oName.ToUpper();
                                        //Console.WriteLine("PW: {0} - XLS: {1}",NameInPW,NameInXls);
                                        int n = 0;
                                        for (int i = 0; i < ColumnsList.Count; i++)
                                        {
                                            string colLabel = ColumnsList[i].LabelText;

                                            string colValue = string.Empty;

                                            try
                                            {
                                                colValue = xlrow[colLabel];
                                            }
                                            catch (Exception)
                                            {

                                                colValue = string.Empty;
                                            }
                                            if (colValue != string.Empty && NameInPW == NameInXls)//xlrow[colName]!="")
                                            {
                                                //log.write("Info", "Id: " + i + " - Column: " + colLabel + " - Vaule: " + colValue);
                                                bool UpdateLinkData = PWMethods.aaApi_UpdateLinkDataColumnValue(TableId, ColumnsList[i].ColumnId, colValue);
                                                Console.WriteLine("Update Link Data: {0} - Column: {1} - Id: {2} - Vaule: {3}", UpdateLinkData, colLabel, ColumnsList[i].ColumnId, colValue);
                                                n++;
                                            }
                                        }

                                        if (n != 0)
                                        {
                                            bool UpdateEnv = false;
                                            try
                                            {
                                                UpdateEnv = PWMethods.aaApi_UpdateEnvAttr(TableId, int.Parse(attrnoValue));

                                            }
                                            catch (Exception)
                                            {

                                                continue;
                                            }

                                            if (UpdateEnv)
                                            {
                                                Console.WriteLine("{0} Attributes Updated", n);
                                                log.write("Info", n + " Attributes Updated");
                                                break;
                                            }
                                        }
                                    }
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("Couldn't Read Excel or Update attributes, Error is:");
                                    Console.WriteLine(e.Message);

                                    log.write("Error", "Couldn't Read Excel or Update attributes, Error is:");
                                    log.write("Error", e.Message);
                                    continue;
                                }
                            }
                        }
                        //Case no record in Env
                        if (selectedRows == 0)
                        {
                            try
                            {
                                bool blankCreated = PWMethods.aaApi_CreateLinkDataAndLink(TableId, 1, PWFolderId, docId, 0, "", 1);
                                Console.WriteLine("Blank Created?:" + blankCreated);
                                selectedRows = PWMethods.aaApi_SelectLinkDataByObject(TableId, 2, PWFolderId, docId, "", 0, 0, 0);
                                goto rowExists;
                            }
                            catch (Exception e)
                            {
                                log.write("Error", "Case no record in Env");
                                log.write("Error", e.Message);
                                continue;
                            }
                        }
                        //Reselect the Link data
                        try
                        {
                            selectedRows = PWMethods.aaApi_SelectLinkDataByObject(TableId, 2, PWFolderId, docId, "", 0, 0, 0);
                        }
                        catch (Exception e)
                        {
                            log.write("Error", "Reselect the Link data");
                            log.write("Error", e.Message);
                            continue;
                        }

                        //Prepare the email row
                        if (docId != 0)
                        {
                            try
                            {
                                for (int x = 0; x < columns.Count; x++)
                                {
                                    emailColumn oneCol = new emailColumn();
                                    oneCol.ColumnName = columns[x];
                                    foreach (var m in ColumnsList)
                                    {
                                        if (m.LabelText == columns[x])
                                        {
                                            oneCol.ColumnId = m.ColumnId;
                                            break;
                                        }
                                        else
                                        {
                                            oneCol.ColumnId = 0;
                                        }
                                    }
                                    //Case env Attribute
                                    if (oneCol.ColumnId > 0)
                                    {
                                        oneCol.ColumnVaule = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, oneCol.ColumnId - 1));
                                    }
                                    //Case General Attribute or missing column
                                    if (oneCol.ColumnId == 0)
                                    {
                                        //int selectDocument = PWMethods.aaApi_SelectDocument(PWFolderId, docId);
                                        //get value by column name
                                        if (oneCol.ColumnName == "Name")
                                        {
                                            oneCol.ColumnVaule = oName;
                                            oneCol.URL = pwPath + oName.Replace(" ", "&space;");
                                        }
                                        if (oneCol.ColumnName == "Sequence")
                                        {
                                            oneCol.ColumnVaule = PWMethods.aaApi_GetDocumentNumericProperty(2, 0).ToString();
                                        }
                                        if (oneCol.ColumnName == "File Updated")
                                        {
                                            oneCol.ColumnVaule = Marshal.PtrToStringUni(PWMethods.aaApi_GetDocumentStringProperty(25, 0));
                                        }
                                        if (oneCol.ColumnName == "File Updated By")
                                        {
                                            int userId = PWMethods.aaApi_GetDocumentNumericProperty(37, 0);
                                            int userSelectd = PWMethods.aaApi_SelectUser(userId);
                                            if (userSelectd == 1)
                                            {
                                                oneCol.ColumnVaule = Marshal.PtrToStringUni(PWMethods.aaApi_GetUserStringProperty(2, 0));
                                            }
                                        }
                                    }
                                    //Console.WriteLine("ColId: {0} - Name: {1} - Vaule: {2}", oneCol.ColumnId, oneCol.ColumnName, oneCol.ColumnVaule);
                                    emlColumns.Add(oneCol);
                                }
                            }
                            catch (Exception e)
                            {
                                log.write("Error", "Prepare the email row");
                                log.write("Exception", e.Message);
                                continue;
                            }
                        }

                        // copy to Z drive: - Overwrite
                        if (myconfig.CopytoFileSystem)
                        {
                            Console.WriteLine("Copying " + ofileName + " to folder: " + item.PDriveFolder);
                            log.write("Info", "Copying " + ofileName + " to folder: " + item.PDriveFolder);

                            try
                            {
                                System.IO.File.Copy(ofilePath, Path.Combine(item.PDriveFolder, ofileName), true);

                            }
                            catch (Exception ex)
                            {

                                log.write("Copy Exception", ex.Message);
                                continue;
                            }
                        }
                        //Delete/Move File/s from Folder in here or outside the loop
                        try
                        {
                            File.Delete(Path.Combine(item.stagingfolder, ofileName));
                        }
                        catch (Exception ex)
                        {
                            log.write("Delete Exception", ex.Message);
                            continue;
                        }
                    }
                    catch (Exception e)
                    {
                        log.write("Error", e.Message);
                        continue;
                    }
                }
            }

            // Notify Users 
            if (myconfig.SendNotification)
            {
                if (emlColumns.Count > 0)
                {
                    appMethods.SendMail2(emlColumns, columns.Count);
                }
                
            }

        }


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
            body = "The following files located in ProjectWise are issued for your action/information. Click on Name to access files.";
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
                if (h == 0)
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

                if (h == NuCols - 1)
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

        public static bool ValidateStates(int PWFolderId)
        {
            bool Validated = false;
            wfObject myWF = new wfObject(); //set an object to be used in WF state change
            wfObject WFObj = xmlHelper.getWFConfigs(); //get what stored in xml

            int wfCount = PWMethods.aaApi_SelectWorkflowsForProject(PWFolderId);
            //Console.WriteLine("Number of WF: " + wfCount);

            for (int i = 0; i < wfCount; i++)
            {
                myWF.WorkflowName = Marshal.PtrToStringUni(PWMethods.aaApi_GetWorkflowStringProperty(3, i));
                if (myWF.WorkflowName == WFObj.WorkflowName)
                {
                    myWF.WorkflowId = PWMethods.aaApi_GetWorkflowNumericProperty(1, i);
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

            //Console.WriteLine("Shared state id is: " + myWF.StateOneId);
            //Console.WriteLine("Archived states id is: " + myWF.StateTwoId);

            if (myWF.StateOneId >0 && myWF.StateTwoId >0 )
            {
                Validated = true;
            }
            
            return Validated;
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
                    myWF.WorkflowId = PWMethods.aaApi_GetWorkflowNumericProperty(1, i);
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

            //Console.WriteLine("Shared state id is: " + myWF.StateOneId);
            //Console.WriteLine("Archived states id is: " + myWF.StateTwoId);

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
            //get Version from xls
            string ver = string.Empty;


            //creat new Version here
            PWMethods.aaApi_NewDocumentVersion(1, folderId, docId, ver, "Version created by automation tool", 0);

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
