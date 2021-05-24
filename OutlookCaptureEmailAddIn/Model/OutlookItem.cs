using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCaptureEmailAddIn.Model
{
    class OutlookItem
    {
        private object m_item;  // the wrapped Outlook item
        private Type m_type;  // type for the Outlook item 
        private object[] m_args;  // dummy argument array
        private System.Type m_typeOlObjectClass;

        #region OutlookItem Constants

        private const string OlActions = "Actions";
        private const string OlApplication = "Application";
        private const string OlAttachments = "Attachments";
        private const string OlBillingInformation = "BillingInformation";
        private const string OlBody = "Body";
        private const string OlCategories = "Categories";
        private const string OlClass = "Class";
        private const string OlClose = "Close";
        private const string OlCompanies = "Companies";
        private const string OlConversationIndex = "ConversationIndex";
        private const string OlConversationTopic = "ConversationTopic";
        private const string OlCopy = "Copy";
        private const string OlCreationTime = "CreationTime";
        private const string OlDisplay = "Display";
        private const string OlDownloadState = "DownloadState";
        private const string OlEntryID = "EntryID";
        private const string OlFormDescription = "FormDescription";
        private const string OlGetInspector = "GetInspector";
        private const string OlImportance = "Importance";
        private const string OlIsConflict = "IsConflict";
        private const string OlItemProperties = "ItemProperties";
        private const string OlLastModificationTime = "LastModificationTime";
        private const string OlLinks = "Links";
        private const string OlMarkForDownload = "MarkForDownload";
        private const string OlMessageClass = "MessageClass";
        private const string OlMileage = "Mileage";
        private const string OlMove = "Move";
        private const string OlNoAging = "NoAging";
        private const string OlOutlookInternalVersion = "OutlookInternalVersion";
        private const string OlOutlookVersion = "OutlookVersion";
        private const string OlParent = "Parent";
        private const string OlPrintOut = "PrintOut";
        private const string OlPropertyAccessor = "PropertyAccessor";
        private const string OlSave = "Save";
        private const string OlSaveAs = "SaveAs";
        private const string OlSaved = "Saved";
        private const string OlSensitivity = "Sensitivity";
        private const string OlSession = "Session";
        private const string OlShowCategoriesDialog = "ShowCategoriesDialog";
        private const string OlSize = "Size";
        private const string OlSubject = "Subject";
        private const string OlUnRead = "UnRead";
        private const string OlUserProperties = "UserProperties";
        private const string pSenderEmailAddress = "SenderEmailAddress";
        private const string pEntryId = "EntryId";
        private const string pFolderPath = "FolderPath";
        private const string pStoreID = "StoreID";
        #endregion

        #region Constructor
        public OutlookItem(object item)
        {
            m_item = item;
            m_type = m_item.GetType();
            m_args = new Object[] { };
        }
        #endregion

        #region Public Methods and Properties
        public Microsoft.Office.Interop.Outlook.Actions Actions
        {
            get
            {
                return this.GetPropertyValue(OlActions) as Microsoft.Office.Interop.Outlook.Actions;
            }
        }

        public Microsoft.Office.Interop.Outlook.Application Application
        {
            get
            {
                return this.GetPropertyValue(OlApplication) as Microsoft.Office.Interop.Outlook.Application;
            }
        }

        public Microsoft.Office.Interop.Outlook.Attachments Attachments
        {
            get
            {
                return this.GetPropertyValue(OlAttachments) as Microsoft.Office.Interop.Outlook.Attachments;
            }
        }

        public string BillingInformation
        {
            get
            {
                return this.GetPropertyValue(OlBillingInformation).ToString();
            }
            set
            {
                SetPropertyValue(OlBillingInformation, value);
            }
        }

        public string Body
        {
            get
            {
                return this.GetPropertyValue(OlBody).ToString();
            }
            set
            {
                SetPropertyValue(OlBody, value);
            }
        }

        public string SenderEmailAddress
        {
            get
            {
                return this.GetPropertyValue(pSenderEmailAddress).ToString();
            }
            set
            {
                SetPropertyValue(pSenderEmailAddress, value);
            }
        }

        public string EntryId
        {
            get
            {
                return this.GetPropertyValue(pEntryId).ToString();
            }
            set
            {
                SetPropertyValue(pEntryId, value);
            }
        }

        public string FolderPath
        {
            get
            {
                return this.GetPropertyValue(pFolderPath).ToString();
            }
            set
            {
                SetPropertyValue(pFolderPath, value);
            }
        }

        public string StoreID
        {
            get
            {
                return this.GetPropertyValue(pStoreID).ToString();
            }
            set
            {
                SetPropertyValue(pStoreID, value);
            }
        }

        public string Categories
        {
            get
            {
                return this.GetPropertyValue(OlCategories).ToString();
            }
            set
            {
                SetPropertyValue(OlCategories, value);
            }
        }


        public void Close(Microsoft.Office.Interop.Outlook.OlInspectorClose SaveMode)
        {
            object[] MyArgs = { SaveMode };
            this.CallMethod(OlClose);
        }

        public string Companies
        {
            get
            {
                return this.GetPropertyValue(OlCompanies).ToString();
            }
            set
            {
                SetPropertyValue(OlCompanies, value);
            }
        }

        public Microsoft.Office.Interop.Outlook.OlObjectClass Class
        {
            get
            {
                if (m_typeOlObjectClass == null)
                {
                    // Note: instantiate dummy ObjectClass enumeration to get type.
                    //       type = System.Type.GetType("Outlook.OlObjectClass") doesn't seem to work
                    Microsoft.Office.Interop.Outlook.OlObjectClass objClass = Microsoft.Office.Interop.Outlook.OlObjectClass.olAction;
                    m_typeOlObjectClass = objClass.GetType();
                }
                return (Microsoft.Office.Interop.Outlook.OlObjectClass)System.Enum.ToObject(m_typeOlObjectClass, this.GetPropertyValue(OlClass));
            }
        }

        public string ConversationIndex
        {
            get
            {
                return this.GetPropertyValue(OlConversationIndex).ToString();
            }
        }

        public string ConversationTopic
        {
            get
            {
                return this.GetPropertyValue(OlConversationTopic).ToString();
            }
        }

        public object Copy()
        {
            return (this.CallMethod(OlCopy));
        }

        public System.DateTime CreationTime
        {
            get
            {
                return (System.DateTime)this.GetPropertyValue(OlCreationTime);
            }
        }

        public void Display()
        {
            this.CallMethod(OlDisplay);
        }

        public Microsoft.Office.Interop.Outlook.OlDownloadState DownloadState
        {
            get
            {
                return (Microsoft.Office.Interop.Outlook.OlDownloadState)this.GetPropertyValue(OlDownloadState);
            }
        }

        public string EntryID
        {
            get
            {
                return this.GetPropertyValue(OlEntryID).ToString();
            }
        }

        public Microsoft.Office.Interop.Outlook.FormDescription FormDescription
        {
            get
            {
                return (Microsoft.Office.Interop.Outlook.FormDescription)this.GetPropertyValue(OlFormDescription);
            }
        }


        public Object InnerObject
        {
            get
            {
                return this.m_item;
            }
        }

        public Microsoft.Office.Interop.Outlook.Inspector GetInspector
        {
            get
            {
                return this.GetPropertyValue(OlGetInspector) as Microsoft.Office.Interop.Outlook.Inspector;
            }
        }

        public Microsoft.Office.Interop.Outlook.OlImportance Importance
        {
            get
            {
                return (Microsoft.Office.Interop.Outlook.OlImportance)this.GetPropertyValue(OlImportance);
            }
            set
            {
                SetPropertyValue(OlImportance, value);
            }
        }

        public bool IsConflict
        {
            get
            {
                return (bool)this.GetPropertyValue(OlIsConflict);
            }
        }

        public Microsoft.Office.Interop.Outlook.ItemProperties ItemProperties
        {
            get
            {
                return (Microsoft.Office.Interop.Outlook.ItemProperties)this.GetPropertyValue(OlItemProperties);
            }
        }

        public System.DateTime LastModificationTime
        {
            get
            {
                return (System.DateTime)this.GetPropertyValue(OlLastModificationTime);
            }
        }

        public Microsoft.Office.Interop.Outlook.Links Links
        {
            get
            {
                return this.GetPropertyValue(OlLinks) as Microsoft.Office.Interop.Outlook.Links;
            }
        }

        public Microsoft.Office.Interop.Outlook.OlRemoteStatus MarkForDownload
        {
            get
            {
                return (Microsoft.Office.Interop.Outlook.OlRemoteStatus)this.GetPropertyValue(OlMarkForDownload);
            }
            set
            {
                SetPropertyValue(OlMarkForDownload, value);
            }
        }

        public string MessageClass
        {
            get
            {
                return this.GetPropertyValue(OlMessageClass).ToString();
            }
            set
            {
                SetPropertyValue(OlMessageClass, value);
            }
        }

        public string Mileage
        {
            get
            {
                return this.GetPropertyValue(OlMileage).ToString();
            }
            set
            {
                SetPropertyValue(OlMileage, value);
            }
        }

        public object Move(Microsoft.Office.Interop.Outlook.Folder DestinationFolder)
        {
            object[] myArgs = { DestinationFolder };
            return this.CallMethod(OlMove, myArgs);
        }

        public bool NoAging
        {
            get
            {
                return (bool)this.GetPropertyValue(OlNoAging);
            }
            set
            {
                SetPropertyValue(OlNoAging, value);
            }
        }

        public long OutlookInternalVersion
        {
            get
            {
                return (long)this.GetPropertyValue(OlOutlookInternalVersion);
            }
        }

        public string OutlookVersion
        {
            get
            {
                return this.GetPropertyValue(OlOutlookVersion).ToString();
            }
        }

        public Microsoft.Office.Interop.Outlook.Folder Parent
        {
            get
            {
                return this.GetPropertyValue(OlParent) as Microsoft.Office.Interop.Outlook.Folder;
            }
        }

        public Microsoft.Office.Interop.Outlook.PropertyAccessor PropertyAccessor
        {
            get
            {
                return this.GetPropertyValue(OlPropertyAccessor) as Microsoft.Office.Interop.Outlook.PropertyAccessor;
            }
        }

        public void PrintOut()
        {
            this.CallMethod(OlPrintOut);
        }

        public void Save()
        {
            this.CallMethod(OlSave);
        }

        public void SaveAs(string path, Microsoft.Office.Interop.Outlook.OlSaveAsType type)
        {
            object[] myArgs = { path, type };
            this.CallMethod(OlSaveAs, myArgs);
        }

        public bool Saved
        {
            get
            {
                return (bool)this.GetPropertyValue(OlSaved);
            }
        }

        public Microsoft.Office.Interop.Outlook.OlSensitivity Sensitivity
        {
            get
            {
                return (Microsoft.Office.Interop.Outlook.OlSensitivity)this.GetPropertyValue(OlSensitivity);
            }
            set
            {
                SetPropertyValue(OlSensitivity, value);
            }
        }

        public Microsoft.Office.Interop.Outlook.NameSpace Session
        {
            get
            {
                return this.GetPropertyValue(OlSession) as Microsoft.Office.Interop.Outlook.NameSpace;
            }
        }

        public void ShowCategoriesDialog()
        {
            this.CallMethod(OlShowCategoriesDialog);
        }

        public long Size
        {
            get
            {
                return (long)this.GetPropertyValue(OlSize);
            }
        }

        public string Subject
        {
            get
            {
                return this.GetPropertyValue(OlSubject).ToString();
            }
            set
            {
                SetPropertyValue(OlSubject, value);
            }
        }

        public bool UnRead
        {
            get
            {
                return (bool)this.GetPropertyValue(OlUnRead);
            }
            set
            {
                SetPropertyValue(OlUnRead, value);
            }
        }

        public Microsoft.Office.Interop.Outlook.UserProperties UserProperties
        {
            get
            {
                return this.GetPropertyValue(OlUserProperties) as Microsoft.Office.Interop.Outlook.UserProperties;
            }
        }

        #endregion

        #region Private Helper Functions
        private object GetPropertyValue(string propertyName)
        {
            try
            {
                // An invalid property name exception is propagated to client
                return m_type.InvokeMember(
                    propertyName,
                    BindingFlags.Public | BindingFlags.GetField | BindingFlags.GetProperty,
                    null,
                    m_item,
                    m_args);
            }
            catch (SystemException ex)
            {
                Debug.WriteLine(
                    string.Format(
                    "OutlookItem: GetPropertyValue for {0} Exception: {1} ",
                    propertyName, ex.Message));
                throw;
            }
        }

        private void SetPropertyValue(string propertyName, object propertyValue)
        {
            try
            {
                m_type.InvokeMember(
                    propertyName,
                    BindingFlags.Public | BindingFlags.SetField | BindingFlags.SetProperty,
                    null,
                    m_item,
                    new object[] { propertyValue });
            }
            catch (SystemException ex)
            {
                Debug.WriteLine(
                   string.Format(
                   "OutlookItem: SetPropertyValue for {0} Exception: {1} ",
                   propertyName, ex.Message));
                throw;
            }
        }

        private object CallMethod(string methodName)
        {
            try
            {
                // An invalid property name exception is propagated to client
                return m_type.InvokeMember(
                    methodName,
                    BindingFlags.Public | BindingFlags.InvokeMethod,
                    null,
                    m_item,
                    m_args);
            }
            catch (SystemException ex)
            {
                Debug.WriteLine(
                    string.Format(
                    "OutlookItem: CallMethod for {0} Exception: {1} ",
                    methodName, ex.Message));
                throw;
            }
        }

        private object CallMethod(string methodName, object[] args)
        {
            try
            {
                // An invalid property name exception is propagated to client
                return m_type.InvokeMember(
                    methodName,
                    BindingFlags.Public | BindingFlags.InvokeMethod,
                    null,
                    m_item,
                    args);
            }
            catch (SystemException ex)
            {
                Debug.WriteLine(
                    string.Format(
                    "OutlookItem: CallMethod for {0} Exception: {1} ",
                    methodName, ex.Message));
                throw;
            }
        }
        #endregion

    }
}
