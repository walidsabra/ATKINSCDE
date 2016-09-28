using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using CDEAutomation.classes;

namespace CDEAutomation
{
    class xmlHelper
    {
        /// <summary>
        /// Get WF Section from xml
        /// </summary>
        /// <returns></returns>
        public static wfObject getWFConfigs()
        {
            wfObject WFObj = new wfObject();
            
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("configurations\\config.xml");

            XmlNodeList xnList = xmlDoc.SelectNodes("/Config/workflow");
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
            xmlDoc.Load("configurations\\config.xml");

            XmlNodeList xnList = xmlDoc.SelectNodes("/Config/appConfig");
            foreach (XmlNode item in xnList)
            {
                myappconfig.copytoP = item["copytofileShare"].InnerText;
                myappconfig.SendNotification = item["sendNotification"].InnerText;
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
            xmlDoc.Load("configurations\\config.xml");

            XmlNodeList xnList = xmlDoc.SelectNodes("/Config/FoldersMapping/FolderMapping");
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
            xmlDoc.Load("configurations\\config.xml");

            XmlNodeList xnList = xmlDoc.SelectNodes("/Config/ProjectWiseLogin");
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
