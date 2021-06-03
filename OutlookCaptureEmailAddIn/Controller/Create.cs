using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
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
            record.Condition = "SenderEmailAddress";

            Model.DB.CreateDelete(record);
        }

        public static void DeleteDomain(Model.POCO.MailInfo _info)
        {
            Model.POCO.Delete record = new Model.POCO.Delete();

            MailAddress addr = new MailAddress(_info.SenderEmailAddress);

            record.SenderEmailAddress = "@" + addr.Host;
            record.StoreID = _info.StoreID;
            record.Condition = "Domain";

            Model.DB.CreateDelete(record);
        }

        public static void Move(Model.POCO.MailInfo _mail, Model.POCO.FolderInfo _folder, string _condition, string _subject)
        {
            Model.POCO.Move record = new Model.POCO.Move();
            record.SenderEmailAddress = _mail.SenderEmailAddress;
            record.EntryID = _folder.EntryID;
            record.FolderPath = _folder.FolderPath;
            record.StoreID = _folder.StoreID;
            record.Condition = _condition;
            record.Subject = _subject;

            Model.DB.CreateMove(record);
        }
    }
}
