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
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns.Add("Email", "Email");
            dataGridView1.Columns.Add("Action", "Action");
            dataGridView1.Columns.Add("Target", "Target");

            for(int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            foreach (var doc in Controller.Read.GetAllRules())
            {
                dataGridView1.Rows.Add(doc.Value, doc.Action, doc.Target);
            }
        }
    }
}
