using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using OutlookCaptureEmailAddIn.Model;

namespace OutlookCaptureEmailAddIn
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Model.POCO.MailInfo mail = Controller.Read.Selection();

            var rule = Controller.Read.FindDeleteRule(mail);

            //if(mail.SenderEmailAddress.IndexOf("@") == -1 && !mail.SenderEmailAddress.StartsWith(@"/"))
            //{
            //    string message = mail.SenderEmailAddress + " is an invalid email";
            //    MessageBox.Show(message, "Information", MessageBoxButtons.OK);
            //}
            //else if (rule != null)
            //{
            //    string message = mail.SenderEmailAddress + " rule already exists";
            //    MessageBox.Show(message, "Information", MessageBoxButtons.OK);
            //}
            //else
            //{
            //    Controller.Create.Delete(mail);
            //}

            Form3 form = new Form3();
            form.ShowDialog();
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Model.POCO.MailInfo mail = Controller.Read.Selection();

            var rule = Controller.Read.FindMoveRule(mail);

            if (mail.SenderEmailAddress.IndexOf("@") == -1 && !mail.SenderEmailAddress.StartsWith(@"/"))
            {
                string message = mail.SenderEmailAddress + " is an invalid email";
                MessageBox.Show(message, "Information", MessageBoxButtons.OK);
            }
            else
            {
                Form1 form = new Form1();
                form.ShowDialog();
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Form2 form = new Form2();
            form.ShowDialog();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            Model.POCO.MailInfo mail = Controller.Read.Selection();

            var rule = Controller.Read.FindDeleteRule(mail);

            if (mail.SenderEmailAddress.IndexOf("@") == -1 && !mail.SenderEmailAddress.StartsWith(@"/"))
            {
                string message = mail.SenderEmailAddress + " is an invalid email";
                MessageBox.Show(message, "Information", MessageBoxButtons.OK);
            }
            else if (rule != null)
            {
                string message = mail.SenderEmailAddress + " rule already exists";
                MessageBox.Show(message, "Information", MessageBoxButtons.OK);
            }
            else
            {
                //Controller.Create.DeleteDomain(mail);
            }
        }
    }
}
