using Microsoft.Office.Interop.Outlook;
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

                MailItem mailItem = (MailItem)selection.Application.GetNamespace("MAPI").GetItemFromID(ret.EntryID);

                ret.Folder = myItem.Parent;
                ret.StoreID = ret.Folder.StoreID;
                ret.Header = mailItem.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E");
                ret.Actions = Convert.ToString(mailItem.Actions);
                ret.AlternateRecipientAllowed = Convert.ToString(mailItem.AlternateRecipientAllowed);
                ret.Application = Convert.ToString(mailItem.Application);
                ret.Attachments = Convert.ToString(mailItem.Attachments);
                ret.AutoForwarded = Convert.ToString(mailItem.AutoForwarded);
                ret.AutoResolvedWinner = Convert.ToString(mailItem.AutoResolvedWinner);
                ret.BCC = Convert.ToString(mailItem.BCC);
                ret.BillingInformation = Convert.ToString(mailItem.BillingInformation);
                ret.Body = Convert.ToString(mailItem.Body);
                ret.BodyFormat = Convert.ToString(mailItem.BodyFormat);
                ret.Categories = Convert.ToString(mailItem.Categories);
                ret.CC = Convert.ToString(mailItem.CC);
                ret.Class = Convert.ToString(mailItem.Class);
                ret.Companies = Convert.ToString(mailItem.Companies);
                ret.Conflicts = Convert.ToString(mailItem.Conflicts);
                ret.ConversationID = Convert.ToString(mailItem.ConversationID);
                ret.ConversationIndex = Convert.ToString(mailItem.ConversationIndex);
                ret.ConversationTopic = Convert.ToString(mailItem.ConversationTopic);
                ret.CreationTime = Convert.ToString(mailItem.CreationTime);
                ret.DeferredDeliveryTime = Convert.ToString(mailItem.DeferredDeliveryTime);
                ret.DeleteAfterSubmit = Convert.ToString(mailItem.DeleteAfterSubmit);
                ret.DownloadState = Convert.ToString(mailItem.DownloadState);
                ret.EnableSharedAttachments = Convert.ToString(mailItem.EnableSharedAttachments);
                ret.EntryID = Convert.ToString(mailItem.EntryID);
                ret.ExpiryTime = Convert.ToString(mailItem.ExpiryTime);
                ret.FlagDueBy = Convert.ToString(mailItem.FlagDueBy);
                ret.FlagIcon = Convert.ToString(mailItem.FlagIcon);
                ret.FlagRequest = Convert.ToString(mailItem.FlagRequest);
                ret.FlagStatus = Convert.ToString(mailItem.FlagStatus);
                ret.FormDescription = Convert.ToString(mailItem.FormDescription);
                ret.GetInspector = Convert.ToString(mailItem.GetInspector);
                ret.HasCoverSheet = Convert.ToString(mailItem.HasCoverSheet);
                ret.HTMLBody = Convert.ToString(mailItem.HTMLBody);
                ret.Importance = Convert.ToString(mailItem.Importance);
                ret.InternetCodepage = Convert.ToString(mailItem.InternetCodepage);
                ret.IsConflict = Convert.ToString(mailItem.IsConflict);
                ret.IsIPFax = Convert.ToString(mailItem.IsIPFax);
                ret.IsMarkedAsTask = Convert.ToString(mailItem.IsMarkedAsTask);
                ret.ItemProperties = Convert.ToString(mailItem.ItemProperties);
                ret.LastModificationTime = Convert.ToString(mailItem.LastModificationTime);
                ret.Links = Convert.ToString(mailItem.Links);
                ret.MAPIOBJECT = Convert.ToString(mailItem.MAPIOBJECT);
                ret.MarkForDownload = Convert.ToString(mailItem.MarkForDownload);
                ret.MessageClass = Convert.ToString(mailItem.MessageClass);
                ret.Mileage = Convert.ToString(mailItem.Mileage);
                ret.NoAging = Convert.ToString(mailItem.NoAging);
                ret.OriginatorDeliveryReportRequested = Convert.ToString(mailItem.OriginatorDeliveryReportRequested);
                ret.OutlookInternalVersion = Convert.ToString(mailItem.OutlookInternalVersion);
                ret.OutlookVersion = Convert.ToString(mailItem.OutlookVersion);
                ret.Parent = Convert.ToString(mailItem.Parent);
                ret.Permission = Convert.ToString(mailItem.Permission);
                ret.PermissionService = Convert.ToString(mailItem.PermissionService);
                ret.PermissionTemplateGuid = Convert.ToString(mailItem.PermissionTemplateGuid);
                ret.PropertyAccessor = Convert.ToString(mailItem.PropertyAccessor);
                ret.ReadReceiptRequested = Convert.ToString(mailItem.ReadReceiptRequested);
                ret.ReceivedByEntryID = Convert.ToString(mailItem.ReceivedByEntryID);
                ret.ReceivedByName = Convert.ToString(mailItem.ReceivedByName);
                ret.ReceivedOnBehalfOfEntryID = Convert.ToString(mailItem.ReceivedOnBehalfOfEntryID);
                ret.ReceivedOnBehalfOfName = Convert.ToString(mailItem.ReceivedOnBehalfOfName);
                ret.ReceivedTime = Convert.ToString(mailItem.ReceivedTime);
                ret.RecipientReassignmentProhibited = Convert.ToString(mailItem.RecipientReassignmentProhibited);
                ret.Recipients = Convert.ToString(mailItem.Recipients);
                ret.ReminderOverrideDefault = Convert.ToString(mailItem.ReminderOverrideDefault);
                ret.ReminderPlaySound = Convert.ToString(mailItem.ReminderPlaySound);
                ret.ReminderSet = Convert.ToString(mailItem.ReminderSet);
                ret.ReminderSoundFile = Convert.ToString(mailItem.ReminderSoundFile);
                ret.ReminderTime = Convert.ToString(mailItem.ReminderTime);
                ret.RemoteStatus = Convert.ToString(mailItem.RemoteStatus);
                ret.ReplyRecipientNames = Convert.ToString(mailItem.ReplyRecipientNames);
                ret.ReplyRecipients = Convert.ToString(mailItem.ReplyRecipients);
                //ret.RetentionExpirationDate = Convert.ToString(mailItem.RetentionExpirationDate);
                //ret.RetentionPolicyName = Convert.ToString(mailItem.RetentionPolicyName);
                ret.RTFBody = Convert.ToString(mailItem.RTFBody);
                ret.Saved = Convert.ToString(mailItem.Saved);
                ret.SaveSentMessageFolder = Convert.ToString(mailItem.SaveSentMessageFolder);
                ret.Sender = Convert.ToString(mailItem.Sender);
                ret.SenderEmailAddress = Convert.ToString(mailItem.SenderEmailAddress);
                ret.SenderEmailType = Convert.ToString(mailItem.SenderEmailType);
                ret.SenderName = Convert.ToString(mailItem.SenderName);
                ret.SendUsingAccount = Convert.ToString(mailItem.SendUsingAccount);
                ret.Sensitivity = Convert.ToString(mailItem.Sensitivity);
                ret.Sent = Convert.ToString(mailItem.Sent);
                ret.SentOn = Convert.ToString(mailItem.SentOn);
                ret.SentOnBehalfOfName = Convert.ToString(mailItem.SentOnBehalfOfName);
                ret.Session = Convert.ToString(mailItem.Session);
                ret.Size = Convert.ToString(mailItem.Size);
                ret.Subject = Convert.ToString(mailItem.Subject);
                ret.Submitted = Convert.ToString(mailItem.Submitted);
                ret.TaskCompletedDate = Convert.ToString(mailItem.TaskCompletedDate);
                ret.TaskDueDate = Convert.ToString(mailItem.TaskDueDate);
                ret.TaskStartDate = Convert.ToString(mailItem.TaskStartDate);
                ret.TaskSubject = Convert.ToString(mailItem.TaskSubject);
                ret.To = Convert.ToString(mailItem.To);
                ret.ToDoTaskOrdinal = Convert.ToString(mailItem.ToDoTaskOrdinal);
                ret.UnRead = Convert.ToString(mailItem.UnRead);
                ret.UserProperties = Convert.ToString(mailItem.UserProperties);
                ret.VotingOptions = Convert.ToString(mailItem.VotingOptions);
                ret.VotingResponse = Convert.ToString(mailItem.VotingResponse);
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
            //var rowColl = Model.DB.ReadAllDelete();

            //Model.POCO.Delete ret = new Model.POCO.Delete();

            //ret = rowColl.Where(r => r.SenderEmailAddress == _info.SenderEmailAddress && r.StoreID == _info.StoreID).FirstOrDefault();

            //return ret;
            return null;
        }

        public static Model.POCO.Move FindMoveRule(Model.POCO.MailInfo _info)
        {
            //var rowColl = Model.DB.ReadAllMove();

            //Model.POCO.Move ret = new Model.POCO.Move();

            //ret = rowColl.Where(r => r.SenderEmailAddress == _info.SenderEmailAddress).FirstOrDefault();

            //return ret;
            return null;
        }

        public static List<Model.POCO.Rule> GetAllRules()
        {
            List<Model.POCO.Rule> ret = new List<Model.POCO.Rule>();

            foreach(var record in Model.DB.ReadAllDelete())
            {
                Model.POCO.Rule rule = new Model.POCO.Rule();
                rule.Field = record.Field;
                rule.Value = record.Value;
                rule.Action = "DELETE";
                rule.Target = "";
                rule.StoreID = record.StoreID;

                ret.Add(rule);
            }

            foreach (var record in Model.DB.ReadAllMove())
            {
                Model.POCO.Rule rule = new Model.POCO.Rule();
                rule.Field = record.Field;
                rule.Value = record.Value;
                rule.Action = "MOVE";
                rule.Target = record.FolderPath;
                rule.StoreID = record.StoreID;

                ret.Add(rule);
            }

            var selected = Selection();

            return ret.Where(r => r.StoreID == selected.StoreID).OrderBy(r => r.Value).ToList();
        }
    }
}
