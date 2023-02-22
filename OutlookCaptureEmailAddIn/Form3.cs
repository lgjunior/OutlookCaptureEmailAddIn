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
    public partial class Form3 : Form
    {
        Model.POCO.MailInfo mail;
        string selected = "";
        string value = "";

        public Form3()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            Controller.Create.DeleteRule(mail, selected, value);
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
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

        private void Form3_Load(object sender, EventArgs e)
        {
            mail = Controller.Read.Selection();

            txReplyRecipientName.Text = mail.ReplyRecipientNames;
            txSenderEmailAddress.Text = mail.SenderEmailAddress;
            txSenderName.Text = mail.SenderName;
            txSentOnBehalfOfName.Text = mail.SentOnBehalfOfName;
            txSubject.Text = mail.Subject;
            txHeader.Text = mail.Header;
            txSenderEmailAddressDomain.Text = "@" + txSenderEmailAddress.Text.Split('@')[1];

            rbSenderEmailAddress.Checked = true;
            selected = "SenderEmailAddress";
            value = txSenderEmailAddress.Text;
        }

        private void rbSenderEmailAddressDomain_CheckedChanged(object sender, EventArgs e)
        {
            selected = "SenderEmailAddress";
            value = "@" + txSenderEmailAddress.Text.Split('@')[1];
        }

        private void txSenderEmailAddressDomain_TextChanged(object sender, EventArgs e)
        {
            rbSenderEmailAddressDomain.Checked = true;
            selected = "SenderEmailAddress";
            value = "@" + txSenderEmailAddress.Text.Split('@')[1];
        }
    }
}
