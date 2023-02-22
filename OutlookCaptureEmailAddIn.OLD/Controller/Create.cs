using OutlookCaptureEmailAddIn.Model.POCO;
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
        public static void MoveRule(MailInfo _mailInfo, String _selected, String _value, String _FolderPath)
        {
            Model.POCO.Move record = new Model.POCO.Move();
            record.StoreID = _mailInfo.StoreID;
            record.Value = _value;
            record.Field = _selected;
            record.FolderPath = _FolderPath;

            Model.DB.CreateMove(record);
        }

        public static void DeleteRule(MailInfo _mailInfo, String _selected, String _value)
        {
            Model.POCO.Delete record = new Model.POCO.Delete();
            record.StoreID = _mailInfo.StoreID;
            record.Value = _value;
            record.Field = _selected;

            Model.DB.CreateDelete(record);
        }

    }
}
