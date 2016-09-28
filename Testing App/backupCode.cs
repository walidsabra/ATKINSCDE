using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

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
    }
}
