using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookCaptureEmailAddIn
{
    public partial class Form1 : Form
    {
        Model.POCO.MailInfo mail;
        string selected = "";
        string value = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var folderColl = Controller.Read.FolderCollection();

            mail = Controller.Read.Selection();

            listBox1.DataSource = folderColl;
            listBox1.DisplayMember = "FolderPath";
            listBox1.ValueMember = "EntryID";

            txReplyRecipientName.Text = mail.ReplyRecipientNames;
            txSenderEmailAddress.Text = mail.SenderEmailAddress;
            txSenderName.Text = mail.SenderName;
            txSentOnBehalfOfName.Text = mail.SentOnBehalfOfName;
            txSubject.Text = mail.Subject;
            txHeader.Text = mail.Header;

            rbSenderEmailAddress.Checked = true;
            selected = "SenderEmailAddress";
            value = txSenderEmailAddress.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Model.POCO.FolderInfo folder = (Model.POCO.FolderInfo)listBox1.SelectedItem;

            Controller.Create.MoveRule(mail, selected, value, folder.FolderPath);
            this.Close();

            /*
            var ruleColl = Model.DB.ReadAllMove();
            bool hasRule = ruleColl.Where(r => r.SenderEmailAddress == mail.SenderEmailAddress).Count() > 0;
            if (hasRule)
            {
                if(condition == "Subject")
                {
                    hasRule = ruleColl.Where(r => r.Subject == subject).Count() > 0;
                }
            }

            if (!hasRule)
            {
                Controller.Create.Move(mail, folder, condition, subject);
            }
            else
            {
                List<string> message = new List<string>();
                message.Add("The following rule exist:");
                message.Add("SenderEmailAddress: " + mail.SenderEmailAddress);
                message.Add("Condition: " + condition);
                message.Add("Subject: " + subject);
                MessageBox.Show(String.Join("\n", message), "Information", MessageBoxButtons.OK);
            }

            this.Close();
            */
        }

        private void txSenderName_TextChanged(object sender, EventArgs e)
        {
            rbSenderName.Checked = true;
            selected = "SenderName";
            value = txSenderName.Text;
        }

        private void txSenderEmailAddress_TextChanged(object sender, EventArgs e)
        {
            rbSenderEmailAddress.Checked = true;
            selected = "SenderEmailAddress";
            value = txSenderEmailAddress.Text;
        }

        private void txSentOnBehalfOfName_TextChanged(object sender, EventArgs e)
        {
            rbSentOnBehalfOfName.Checked = true;
            selected = "SentOnBehalfOfName";
            value = txSentOnBehalfOfName.Text;
        }

        private void txReplyRecipientName_TextChanged(object sender, EventArgs e)
        {
            rbReplyRecipientName.Checked = true;
            selected = "ReplyRecipientName";
            value = txReplyRecipientName.Text;
        }

        private void txSubject_TextChanged(object sender, EventArgs e)
        {
            rbSubject.Checked = true;
            selected = "Subject";
            value = txSubject.Text;
        }

        private void txHeader_TextChanged(object sender, EventArgs e)
        {
            rbHeader.Checked = true;
            selected = "Header";
            value = txHeader.Text;
        }

        private void rbSenderName_CheckedChanged(object sender, EventArgs e)
        {
            selected = "SenderName";
            value = txSenderName.Text;
        }

        private void rbSenderEmailAddress_CheckedChanged(object sender, EventArgs e)
        {
            selected = "SenderEmailAddress";
            value = txSenderEmailAddress.Text;
        }

        private void rbSentOnBehalfOfName_CheckedChanged(object sender, EventArgs e)
        {
            selected = "ReplyRecipientName";
            value = txReplyRecipientName.Text;
        }

        private void rbReplyRecipientName_CheckedChanged(object sender, EventArgs e)
        {
            selected = "ReplyRecipientName";
            value = txReplyRecipientName.Text;
        }

        private void rbSubject_CheckedChanged(object sender, EventArgs e)
        {
            selected = "Subject";
            value = txSubject.Text;
        }

        private void rbHeader_CheckedChanged(object sender, EventArgs e)
        {
            selected = "Header";
            value = txHeader.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
