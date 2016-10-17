using System;
using System.Runtime.InteropServices;
using System.Text;

namespace CDEAutomation
{
    class PWMethods
    {

        //c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll
        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll")]
        public static extern bool aaApi_Initialize(uint ulModule);

        //[DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll")]
        //public static extern int aaApi_GetLastErrorId();

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

        //[DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        //public static extern int aaApi_SelectLinks(int lProjectId, int lDocumentId);

        //[DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        //public static extern int aaApi_GetLinkNumericProperty(int lPropertyId, int lIndex);

        //[DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        //public static extern IntPtr aaApi_GetLinkStringProperty(int lPropertyId, int lIndex);

        //[DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        //public static extern int aaApi_SelectChildProjects(int lProjectId);

        //[DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        //public static extern int aaApi_SelectTopLevelProjects();

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_GetProjectIdByNamePath(String lpctstrPath);

        //[DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        //public static extern bool aaApi_GetDocumentNamePath2(
        //    int lProjectId,
        //    int lDocumentId,
        //    bool bUseDesc,
        //    String tchSeparator,
        //    String lptstrBuffer,
        //    int lBufferLen
        //    );


        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectUser(int lUserId);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr aaApi_GetUserStringProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectAllTables();

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr aaApi_GetTableStringProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_GetTableNumericProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectLinkDataByObject(
            int lTableId,
            int lItemType,
            int lItemId1, 
            int lItemId2,
            String lpctstrWhere,
            int lplColumnCount,
            int lplColumnIds,
            int ulFlags
            );

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr aaApi_GetLinkDataColumnValue(int lRowIndex, int lColumnIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr aaApi_GetLinkDataColumnStringProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_GetColumnCount(int lTableId);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_GetColumnNumericProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectColumnsByTable(int lTableId);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectEnvByTableId(int lTableId);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectEnvByProjectId(int lProjectId);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr aaApi_GetColumnStringProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr aaApi_GetEnvStringProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern bool aaApi_UpdateLinkData(int lTableId, int lColumnId, String lpctstrValue);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern bool aaApi_UpdateLinkDataColumnValue(int lTableId, int lColumnId, String lpctstrValue);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern bool aaApi_UpdateEnvAttr(int lTableId, int lAttrRecId);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern bool aaApi_SetLinkDataColumnValue(int lTableId, int lColumnId, String lpctstrValue);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool aaApi_CreateLinkData(int lTableId, int lColumnId, int llengthBuffer);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectEnvAttrGuiDefs(int lEnvironmentId, int lTableId, int lColumnId, int lGuiId, int lPageNo);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr aaApi_GetEnvAttrGuiDefStringProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_GetEnvAttrGuiDefNumericProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_GetEnvNumericProperty(int lPropertyId, int lIndex);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern bool aaApi_CreateLinkDataAndLink(
            int lTableId, 
            int lLinkType, 
            int lObjectId1, 
            int lObjectId2, 
            int lplColumnId,
            String lptstrValueBuffer,
            int llengthBuffer);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectEnv(int lEnvironmentId);

        [DllImport("c:\\Program Files\\Bentley\\ProjectWise\\bin\\dmscli.dll", CharSet = CharSet.Unicode)]
        public static extern int aaApi_SelectDocument(int lProjectId, int lDocumentId);

    }
}
