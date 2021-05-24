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
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var folderColl = Controller.Read.FolderCollection();

            listBox1.DataSource = folderColl;
            listBox1.DisplayMember = "FolderPath";
            listBox1.ValueMember = "EntryID";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Model.POCO.FolderInfo folder = (Model.POCO.FolderInfo)listBox1.SelectedItem;
            Model.POCO.MailInfo mail = Controller.Read.Selection();

            Controller.Create.Move(mail, folder);

            this.Close();
        }
    }
}
