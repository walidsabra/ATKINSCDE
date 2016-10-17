using CDEAutomation.classes;
using System.Collections.Generic;
using System.Xml;


namespace CDEAutomation
{
    class xmlHelper
    {
        public static string location;

        /// <summary>
        /// Get email columns
        /// </summary>
        /// <returns></returns>
        public static List<string> getEmailColumns()
        {
            List<string> columns = new List<string>();
            XmlDocument xmlDoc = new XmlDocument();
            XmlNodeList xnList;

            try
            {

                xmlDoc.Load(location);
                xnList = xmlDoc.SelectNodes("/Config/emailColumns/column");
            }
            catch (System.Exception)
            {

                throw;
            }

            foreach (XmlNode item in xnList)
            {
                columns.Add(item.InnerText);
            }
            return columns;
        }

        /// <summary>
        /// Get Notification Settings from xml
        /// </summary>
        /// <returns></returns>
        public static EnvObject getEnvConfigs()
        {
            EnvObject Env = new EnvObject();
            XmlDocument xmlDoc = new XmlDocument();
            XmlNodeList xnList;

            try
            {

                xmlDoc.Load(location);
                xnList = xmlDoc.SelectNodes("/Config/Environment");
            }
            catch (System.Exception)
            {

                throw;
            }

            foreach (XmlNode item in xnList)
            {
                Env.EnvironmentName = item["EnvName"].InnerText;
                Env.Title = item["TitleCol"].InnerText;
                Env.FileNumber = item["FileNoCol"].InnerText;
                Env.Revision = item["RevisionCol"].InnerText;
                Env.RevisionNote = item["RevNotesCol"].InnerText;
                Env.Suitability = item["SuitabilityCol"].InnerText;
                Env.Package = item["PackageCol"].InnerText;
                Env.DesignStage = item["DesignStageCol"].InnerText;
            }
            return Env;
        }

        /// <summary>
        /// Get Notification Settings from xml
        /// </summary>
        /// <returns></returns>
        public static oNotification getMsgConfigs()
        {
            oNotification msg = new oNotification();
            XmlDocument xmlDoc = new XmlDocument();
            XmlNodeList xnList;

            try
            {

                xmlDoc.Load(location);
                xnList = xmlDoc.SelectNodes("/Config/Notification");
            }
            catch (System.Exception)
            {

                throw;
            }

            foreach (XmlNode item in xnList)
            {
                msg.host = item["Host"].InnerText;
                msg.port = int.Parse(item["Port"].InnerText);
                msg.ssl = bool.Parse(item["SSL"].InnerText);
                msg.To = item["To"].InnerText;
                msg.From = item["From"].InnerText;
                msg.Subject = item["Subject"].InnerText;
                msg.user = item["user"].InnerText;
                msg.password = item["password"].InnerText;
            }
            return msg;
        }

        /// <summary>
        /// Get WF Section from xml
        /// </summary>
        /// <returns></returns>
        public static wfObject getWFConfigs()
        {
            wfObject WFObj = new wfObject();
            
            XmlDocument xmlDoc = new XmlDocument();
            XmlNodeList xnList;

            try
            {
                xmlDoc.Load(location);
                xnList = xmlDoc.SelectNodes("/Config/workflow");
            }
            catch (System.Exception)
            {
                
                throw;
            }
            

       
            foreach (XmlNode item in xnList)
            {
                WFObj.WorkflowName = item["WorkflowName"].InnerText;
                WFObj.SharedStateName = item["ShardedStateName"].InnerText;
                WFObj.ArchivedStateName = item["ArchivedStateName"].InnerText;
            }
            return WFObj;
        }

        /// <summary>
        /// Get True/False for P drive and emails.
        /// </summary>
        /// <returns></returns>
        public static otherConfig getAppConfigs()
        {
            otherConfig myappconfig = new otherConfig();
            //myappconfig = xmlHelper.getAppConfigs();
            
            XmlDocument xmlDoc = new XmlDocument();
            XmlNodeList xnList;
            try
            {
                xmlDoc.Load(location);
                xnList = xmlDoc.SelectNodes("/Config/appConfig");
            }
            catch (System.Exception)
            {
                
                throw;
            }
            

           
            foreach (XmlNode item in xnList)
            {
                myappconfig.CopytoFileSystem = bool.Parse(item["copytofileShare"].InnerText);
                myappconfig.SendNotification = bool.Parse(item["sendNotification"].InnerText);
            }
            return myappconfig;
        }

        /// <summary>
        /// Get list of Mappings in XML
        /// </summary>
        /// <returns></returns>
        public static List<FolderMapping> getMappings()
        {
            List<FolderMapping> myList = new List<FolderMapping>();

            XmlDocument xmlDoc = new XmlDocument();
            XmlNodeList xnList;
            try
            {
                xmlDoc.Load(location);
                xnList = xmlDoc.SelectNodes("/Config/FoldersMapping/FolderMapping");
            }
            catch (System.Exception)
            {
                
                throw;
            }
          
            foreach (XmlNode item in xnList)
            {
                FolderMapping oneMapping = new FolderMapping();

                oneMapping.stagingfolder = item["StagingFolder"].InnerText;
                oneMapping.ProjectWiseFolder = item["ProjectWiseFolder"].InnerText;
                oneMapping.PDriveFolder = item["FileshareMapping"].InnerText;

                myList.Add(oneMapping);
            }
           

            return myList;
        }

        /// <summary>
        /// Provide ProjectWise Datasource login info
        /// </summary>
        /// <returns></returns>
        public static dsInfo getDatasourceInfo()
        {
            dsInfo dsInf = new dsInfo();

            XmlDocument xmlDoc = new XmlDocument();
            XmlNodeList xnList;

            try
            {
                xmlDoc.Load(location);
                xnList = xmlDoc.SelectNodes("/Config/ProjectWiseLogin");
            }
            catch (System.Exception)
            {
                
                throw;
            }
             
            foreach (XmlNode item in xnList)
            {
                dsInf.dsName = item["Datasource"].InnerText;
                dsInf.dsUser = item["UserName"].InnerText;
                dsInf.dsPass = item["password"].InnerText;
            }
            return dsInf;
        }
    }


}
