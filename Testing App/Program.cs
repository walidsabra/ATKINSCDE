using CDEAutomation.classes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;


namespace CDEAutomation
{
    class Program
    {

        static void Main(string[] args)
        {
            string xlsLoc = string.Empty;


            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    if (args[i] == "-c")
                    {
                        xmlHelper.location = args[i + 1];
                    }
                    if (args[i] == "-l")
                    {
                        log.location = args[i + 1];
                    }
                    if (args[i] == "-a")
                    {
                        xlsLoc = args[i + 1];
                    }
                }
            }
            else
            {
                Console.WriteLine("Missing arugments");
                Console.WriteLine(@"Usage: ""CDEAutomation.exe -c config.xml -l log.txt -a attributes.xlsx"" ");

                log.write("Failed", "Missing appication arugments");

                Console.ReadLine();
                return;
            }



            Console.WriteLine("... CDE Automation Tool Started ...");
            log.write("", "CDE Automation Tool Started ");


            List<emailColumns> docList = new List<emailColumns>();
            List<emailColumn> emlColumns = new List<emailColumn>();

            List<string> columns = new List<string>();
            columns = xmlHelper.getEmailColumns();

            dsInfo dsInfo = new dsInfo();
            dsInfo = xmlHelper.getDatasourceInfo();

            EnvObject EnvInfo = new EnvObject();
            EnvInfo = xmlHelper.getEnvConfigs();

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
            }

            //read xls file
            var excel = new LinqToExcel.ExcelQueryFactory();
            excel.FileName = xlsLoc;
            var docAttr = from x in excel.Worksheet()
                          select x;


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
                int TableId= 0;
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
                    ////check if list of records stored in memory to send noti is sent here 
                    //if (myconfig.SendNotification)
                    //{
                    //    if (docList.Count() > 0)
                    //    {
                    //        appMethods.SendMail(docList);
                    //    }

                    //}
                    Console.WriteLine("Error - Folder doesn't exist in ProjectWise");
                    log.write("Error", "Folder doesn't exist in ProjectWise");
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

                List<PWAttr> ColumnsList = new List<PWAttr>();
                for (int f = 0; f < xFolders; f++)
                {
                    int oLastFolderId = PWMethods.aaApi_GetProjectNumericProperty(1, f);
                    if (folderId == oLastFolderId)
                    {
                        PWFolderId = PWMethods.aaApi_GetProjectNumericProperty(1, f);
                        oEnvId = PWMethods.aaApi_GetProjectNumericProperty(21, f);
                        int envCount = PWMethods.aaApi_SelectEnv(oEnvId);

                        TableId = PWMethods.aaApi_GetEnvNumericProperty(2,0);

                        Console.WriteLine("Folder Id: {0}, Env Id: {1}, TableId: {2}", PWFolderId, oEnvId, TableId);
                        log.write("Info", "Folder Id: " + PWFolderId + " Env Id: " + oEnvId + " TableId: " + TableId);

                        appMethods.GetColumnsByEnvID(oEnvId, ColumnsList);
                    }
                }

                //Select files (all general attributes) within PW Folder
                int selected = PWMethods.aaApi_SelectDocumentsByProjectId(PWFolderId);
                Console.WriteLine("Document in Folder?: " + selected);
                log.write("", "Document in Folder?: " + selected);



                //Loop on file in Staging Folder  sorted by Size Ascending 
                foreach (var file in d.GetFiles("*.*").OrderBy(f => f.Length))
                {
                    String ofilePath = file.FullName;
                    String ofileName = file.Name;
                    String oName = Path.GetFileNameWithoutExtension(file.FullName);
                    int docId = 0;


                    //loop on PW Select Files within folder to check if File already Exist - Create Version
                    if (selected > 0)
                    {
                        bool created = false;

                        for (int i = 0; i < selected; i++)
                        {
                            string documentFileName = Marshal.PtrToStringUni(PWMethods.aaApi_GetDocumentStringProperty(21, i));
                            if (documentFileName == ofileName)
                            {
                                //Active Version?
                                //Get General Attributes 

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
                            //Fill docId
                            docId = appMethods.GetDocIdByFolderId(PWFolderId, oName);
                        }
                    }
                    else
                    {
                        Console.WriteLine("Processing New File: " + ofileName);
                        log.write("Info", "Processing New File: " + ofileName);

                        appMethods.CreatPWDoc(PWFolderId, ofilePath, ofileName, oName);
                        //Fill docId
                        docId = appMethods.GetDocIdByFolderId(PWFolderId, oName);
                    }


                    //Update Attribute from xls
                    int colCount = PWMethods.aaApi_SelectColumnsByTable(TableId);
                    int selectedRows = PWMethods.aaApi_SelectLinkDataByObject(TableId, 2, PWFolderId, docId, "", 0, 0, 0);
                    Console.WriteLine("Link rows selected: " + selectedRows);
                rowExists:                
                        //Case document already has attributes in Env
                        if (selectedRows>0)
                        {
                            string attrnoValue = string.Empty;
                            for (int i = 0; i < colCount; i++)
                            {
                                string colName = Marshal.PtrToStringUni(PWMethods.aaApi_GetColumnStringProperty(colCount, i));
                                //Console.WriteLine(colName);
                                if (colName == "a_attrno")
                                {
                                    attrnoValue = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, i));
                                    //Console.WriteLine(attrnoValue);
                                }
                            }

                            foreach (var xlrow in docAttr)
                            {
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
                                    if (colValue !=string.Empty && oName==xlrow["Name"])//xlrow[colName]!="")
                                    {
                                        Console.WriteLine("id: {0} - column: {1} - Value: {2}", i, colLabel, colValue);
                                        log.write("Info", "Id: " + i + " - Column: " + colLabel + " - Vaule: " + colValue);
                                        bool UpdateLinkData = PWMethods.aaApi_UpdateLinkDataColumnValue(TableId, i+1, colValue);
                                        n++;
                                    }
                                }
                                if (n!=0)
                                {               
                                    bool UpdateEnv = PWMethods.aaApi_UpdateEnvAttr(TableId, int.Parse(attrnoValue));
                                    if (UpdateEnv)
                                    {
                                        Console.WriteLine("{0} Attributes Updated", n);
                                        log.write("Info", n + " Attributes Updated");
                                        break;
                                    }
                                }
                            }
                        }
                        //Case no record in Env
                        if (selectedRows==0)
                        {
                            bool blankCreated = PWMethods.aaApi_CreateLinkDataAndLink(TableId, 1, PWFolderId, docId,0,"",1);
                            //Console.WriteLine(blankCreated);
                            selectedRows = PWMethods.aaApi_SelectLinkDataByObject(TableId, 2, PWFolderId, docId, "", 0, 0, 0);
                            goto rowExists;
                        }
                        //Reselect the Link data
                        selectedRows = PWMethods.aaApi_SelectLinkDataByObject(TableId, 2, PWFolderId, docId, "", 0, 0, 0);     

                    //Prepare the email row
                    if (docId != 0)
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
                            if (oneCol.ColumnId >0)
                            {
                                oneCol.ColumnVaule = Marshal.PtrToStringUni(PWMethods.aaApi_GetLinkDataColumnValue(0, oneCol.ColumnId - 1));
                            }
                            //Case General Attribute or missing column
                            if (oneCol.ColumnId==0)
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
                            Console.WriteLine("ColId: {0} - Name: {1} - Vaule: {2}", oneCol.ColumnId , oneCol.ColumnName, oneCol.ColumnVaule);
                            emlColumns.Add(oneCol);
                        }
                    }

                    // copy to Z drive: - Overwrite
                    if (myconfig.CopytoFileSystem)
                    {
                        Console.WriteLine("Copying " + ofileName + " to folder: " + item.PDriveFolder);
                        log.write("", "Copying " + ofileName + " to folder: " + item.PDriveFolder);

                        System.IO.File.Copy(ofilePath, Path.Combine(item.PDriveFolder, ofileName), true);
                    }
                    //Delete/Move File/s from Folder in here or outside the loop
                    File.Delete(Path.Combine(item.stagingfolder, ofileName));
                }
            }

            // Notify Users 
            if (myconfig.SendNotification)
            {
                appMethods.SendMail2(emlColumns,columns.Count);
            }
            Console.ReadLine();
        }
    }
}
