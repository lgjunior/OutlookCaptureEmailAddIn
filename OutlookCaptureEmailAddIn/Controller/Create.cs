using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCaptureEmailAddIn.Controller
{
    public static class Create
    {
        public static void Delete(Model.POCO.MailInfo _info)
        {
            Model.POCO.Delete record = new Model.POCO.Delete();
            record.SenderEmailAddress = _info.SenderEmailAddress;
            record.StoreID = _info.StoreID;

            Model.DB.CreateDelete(record);
        }

        public static void Move(Model.POCO.MailInfo _mail, Model.POCO.FolderInfo _folder)
        {
            Model.POCO.Move record = new Model.POCO.Move();
            record.SenderEmailAddress = _mail.SenderEmailAddress;
            record.EntryID = _folder.EntryID;
            record.FolderPath = _folder.FolderPath;
            record.StoreID = _folder.StoreID;

            Model.DB.CreateMove(record);
        }
    }
}
