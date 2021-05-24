using MongoDB.Bson;
using OutlookCaptureEmailAddIn.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCaptureEmailAddIn.Controller
{
    public static class Read
    {
        public static Microsoft.Office.Interop.Outlook.Application oApp = new Microsoft.Office.Interop.Outlook.Application();

        public static Model.POCO.MailInfo Selection()
        {
            Model.POCO.MailInfo ret = new Model.POCO.MailInfo();

            Microsoft.Office.Interop.Outlook.Selection selection = oApp.ActiveExplorer().Selection;

            for (int i = 1; i <= selection.Count; i++)
            {
                OutlookItem myItem = new OutlookItem(selection[i]);

                ret.EntryID = myItem.EntryID;
                ret.SenderEmailAddress = myItem.SenderEmailAddress;
                ret.StoreID = myItem.Parent.StoreID;
                ret.Folder = myItem.Parent;
            }

            return ret;
        }

        public static List<Model.POCO.FolderInfo> FolderCollection()
        {
            Model.POCO.MailInfo selected = Selection();
            Microsoft.Office.Interop.Outlook.Folder root = selected.Folder.Parent;

            List<Model.POCO.FolderInfo> ret = new List<Model.POCO.FolderInfo>();

            Model.POCO.FolderInfo row = new Model.POCO.FolderInfo();
            row.EntryID = root.EntryID;
            row.FolderPath = root.FolderPath;
            row.StoreID = root.StoreID;
            ret.Add(row);

            foreach(Microsoft.Office.Interop.Outlook.Folder folder in root.Folders)
            {
                row = new Model.POCO.FolderInfo();
                row.EntryID = folder.EntryID;
                row.FolderPath = folder.FolderPath;
                row.StoreID = folder.StoreID;
                ret.Add(row);

                foreach (Microsoft.Office.Interop.Outlook.Folder childFolder in folder.Folders)
                {
                    row = new Model.POCO.FolderInfo();
                    row.EntryID = childFolder.EntryID;
                    row.FolderPath = childFolder.FolderPath;
                    row.StoreID = childFolder.StoreID;
                    ret.Add(row);
                }
            }

            return ret.OrderBy(r => r.FolderPath).ToList();
        }

        public static Model.POCO.Delete FindDeleteRule(Model.POCO.MailInfo _info)
        {
            var rowColl = Model.DB.ReadAllDelete();

            Model.POCO.Delete ret = new Model.POCO.Delete();

            ret = rowColl.Where(r => r.SenderEmailAddress == _info.SenderEmailAddress).FirstOrDefault();

            return ret;
        }

        public static Model.POCO.Move FindMoveRule(Model.POCO.MailInfo _info)
        {
            var rowColl = Model.DB.ReadAllMove();

            Model.POCO.Move ret = new Model.POCO.Move();

            ret = rowColl.Where(r => r.SenderEmailAddress == _info.SenderEmailAddress).FirstOrDefault();

            return ret;
        }
    }
}
