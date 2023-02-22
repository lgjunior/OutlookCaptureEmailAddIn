using MongoDB.Bson;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookCaptureEmailAddIn.Model.POCO
{
    public class MailInfo
    {
        public ObjectId _id { get; set; }
        public Microsoft.Office.Interop.Outlook.Folder Folder { get; set; }
        public string Field { get; set; }
        public string Header { get; set; }
        public string StoreID { get; set; }
        public string Actions { get; set; }
        public string AlternateRecipientAllowed { get; set; }
        public string Application { get; set; }
        public string Attachments { get; set; }
        public string AutoForwarded { get; set; }
        public string AutoResolvedWinner { get; set; }
        public string BCC { get; set; }
        public string BillingInformation { get; set; }
        public string Body { get; set; }
        public string BodyFormat { get; set; }
        public string Categories { get; set; }
        public string CC { get; set; }
        public string Class { get; set; }
        public string Companies { get; set; }
        public string Conflicts { get; set; }
        public string ConversationID { get; set; }
        public string ConversationIndex { get; set; }
        public string ConversationTopic { get; set; }
        public string CreationTime { get; set; }
        public string DeferredDeliveryTime { get; set; }
        public string DeleteAfterSubmit { get; set; }
        public string DownloadState { get; set; }
        public string EnableSharedAttachments { get; set; }
        public string EntryID { get; set; }
        public string ExpiryTime { get; set; }
        public string FlagDueBy { get; set; }
        public string FlagIcon { get; set; }
        public string FlagRequest { get; set; }
        public string FlagStatus { get; set; }
        public string FormDescription { get; set; }
        public string GetInspector { get; set; }
        public string HasCoverSheet { get; set; }
        public string HTMLBody { get; set; }
        public string Importance { get; set; }
        public string InternetCodepage { get; set; }
        public string IsConflict { get; set; }
        public string IsIPFax { get; set; }
        public string IsMarkedAsTask { get; set; }
        public string ItemProperties { get; set; }
        public string LastModificationTime { get; set; }
        public string Links { get; set; }
        public string MAPIOBJECT { get; set; }
        public string MarkForDownload { get; set; }
        public string MessageClass { get; set; }
        public string Mileage { get; set; }
        public string NoAging { get; set; }
        public string OriginatorDeliveryReportRequested { get; set; }
        public string OutlookInternalVersion { get; set; }
        public string OutlookVersion { get; set; }
        public string Parent { get; set; }
        public string Permission { get; set; }
        public string PermissionService { get; set; }
        public string PermissionTemplateGuid { get; set; }
        public string PropertyAccessor { get; set; }
        public string ReadReceiptRequested { get; set; }
        public string ReceivedByEntryID { get; set; }
        public string ReceivedByName { get; set; }
        public string ReceivedOnBehalfOfEntryID { get; set; }
        public string ReceivedOnBehalfOfName { get; set; }
        public string ReceivedTime { get; set; }
        public string RecipientReassignmentProhibited { get; set; }
        public string Recipients { get; set; }
        public string ReminderOverrideDefault { get; set; }
        public string ReminderPlaySound { get; set; }
        public string ReminderSet { get; set; }
        public string ReminderSoundFile { get; set; }
        public string ReminderTime { get; set; }
        public string RemoteStatus { get; set; }
        public string ReplyRecipientNames { get; set; }
        public string ReplyRecipients { get; set; }
        public string RetentionExpirationDate { get; set; }
        public string RetentionPolicyName { get; set; }
        public string RTFBody { get; set; }
        public string Saved { get; set; }
        public string SaveSentMessageFolder { get; set; }
        public string Sender { get; set; }
        public string SenderEmailAddress { get; set; }
        public string SenderEmailType { get; set; }
        public string SenderName { get; set; }
        public string SendUsingAccount { get; set; }
        public string Sensitivity { get; set; }
        public string Sent { get; set; }
        public string SentOn { get; set; }
        public string SentOnBehalfOfName { get; set; }
        public string Session { get; set; }
        public string Size { get; set; }
        public string Subject { get; set; }
        public string Submitted { get; set; }
        public string TaskCompletedDate { get; set; }
        public string TaskDueDate { get; set; }
        public string TaskStartDate { get; set; }
        public string TaskSubject { get; set; }
        public string To { get; set; }
        public string ToDoTaskOrdinal { get; set; }
        public string UnRead { get; set; }
        public string UserProperties { get; set; }
        public string VotingOptions { get; set; }
        public string VotingResponse { get; set; }

    }
}
