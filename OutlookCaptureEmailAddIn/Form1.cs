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

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var folderColl = Controller.Read.FolderCollection();

            mail = Controller.Read.Selection();

            textBox1.Text = mail.Subject;
            comboBox1.SelectedIndex = 0;

            listBox1.DataSource = folderColl;
            listBox1.DisplayMember = "FolderPath";
            listBox1.ValueMember = "EntryID";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Model.POCO.FolderInfo folder = (Model.POCO.FolderInfo)listBox1.SelectedItem;

            String condition = comboBox1.SelectedItem.ToString();
            String subject = condition == "SenderEmailAddress" ? "" : textBox1.Text;

            var ruleColl = Model.DB.ReadAllMove();
            bool hasRule = ruleColl.Where(r => r.SenderEmailAddress == mail.SenderEmailAddress).Count() > 0;
            if (hasRule)
            {
                if(condition == "SUBJECT")
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
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 1;
        }
    }
}
