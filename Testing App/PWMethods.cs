using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CDEAutomation
{
    class PWMethods
    {
        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll")]
        public static extern bool aaApi_Initialize(uint ulModule);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll")]
        public static extern int aaApi_GetLastErrorId(IntPtr A);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern bool aaApi_Login(int IDsType, String DSName, String UserName, String password, String schema);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern bool aaApi_CreateDocument(
            int plDocumentId,
            int lProjectId,
            int lStorageId,
            int lFileType,
            int lItemType,
            int lApplicationId,
            int lDepartmentId,
            int lWorkspaceProfileId,
            String lpctstrSourceFile,
            String lpctstrFileName,
            String lpctstrName,
            String lpctstrDesc,
            String lpctstrVersion,
            bool bLeaveOut,
            uint ulFlags,
            StringBuilder lptstrWorkingFile,
            int lBufferSize,
            int plAttributeId);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectDocumentsByProjectId(int lProjectId);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr aaApi_GetDocumentStringProperty(int lPropertyId, int index);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_GetDocumentNumericProperty(int lPropertyId, int index);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern bool aaApi_ChangeDocumentFile(int lProjectId, int lDocumentId, String lpctstrFileName);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern bool aaApi_NewDocumentVersion(
            int ulFlags, 
            int lProjectId, 
            int lDocumentId, 
            int lpctstrVersion, 
            String lpctstrComment, 
            int lplVersionDocId
            );

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectProjectsByProp(String lpctstrCode, String lpctstrName, String lpctstrDesc, String lpctstrVersion);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_GetProjectNumericProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectWorkflowsForProject(int projectId);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr aaApi_GetWorkflowStringProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_GetWorkflowNumericProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectStatesByWorkflow(int lWorkflowId);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr aaApi_GetStateStringProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_GetStateNumericProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern bool aaApi_SetDocumentState (
            int ulFlags,
            int lProjectId,
            int lDocumentId,
            int lWorkflowId,
            int lStateId,
            String lpctstrComment
            );
    }
}
