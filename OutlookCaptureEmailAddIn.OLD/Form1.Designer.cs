namespace OutlookCaptureEmailAddIn
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.txHeader = new System.Windows.Forms.TextBox();
            this.rbHeader = new System.Windows.Forms.RadioButton();
            this.rbSubject = new System.Windows.Forms.RadioButton();
            this.rbReplyRecipientName = new System.Windows.Forms.RadioButton();
            this.rbSentOnBehalfOfName = new System.Windows.Forms.RadioButton();
            this.txSubject = new System.Windows.Forms.TextBox();
            this.rbSenderEmailAddress = new System.Windows.Forms.RadioButton();
            this.rbSenderName = new System.Windows.Forms.RadioButton();
            this.txSenderName = new System.Windows.Forms.TextBox();
            this.txReplyRecipientName = new System.Windows.Forms.TextBox();
            this.txSenderEmailAddress = new System.Windows.Forms.TextBox();
            this.txSentOnBehalfOfName = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(511, 13);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(277, 420);
            this.listBox1.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 24);
            this.button1.TabIndex = 1;
            this.button1.Text = "Save";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(93, 13);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 2;
            this.button2.Text = "Cancel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txHeader);
            this.groupBox1.Controls.Add(this.rbHeader);
            this.groupBox1.Controls.Add(this.rbSubject);
            this.groupBox1.Controls.Add(this.rbReplyRecipientName);
            this.groupBox1.Controls.Add(this.rbSentOnBehalfOfName);
            this.groupBox1.Controls.Add(this.txSubject);
            this.groupBox1.Controls.Add(this.rbSenderEmailAddress);
            this.groupBox1.Controls.Add(this.rbSenderName);
            this.groupBox1.Controls.Add(this.txSenderName);
            this.groupBox1.Controls.Add(this.txReplyRecipientName);
            this.groupBox1.Controls.Add(this.txSenderEmailAddress);
            this.groupBox1.Controls.Add(this.txSentOnBehalfOfName);
            this.groupBox1.Location = new System.Drawing.Point(12, 47);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(493, 391);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "groupBox1";
            // 
            // txHeader
            // 
            this.txHeader.Location = new System.Drawing.Point(7, 294);
            this.txHeader.Name = "txHeader";
            this.txHeader.Size = new System.Drawing.Size(486, 20);
            this.txHeader.TabIndex = 10;
            this.txHeader.TextChanged += new System.EventHandler(this.txHeader_TextChanged);
            // 
            // rbHeader
            // 
            this.rbHeader.AutoSize = true;
            this.rbHeader.Location = new System.Drawing.Point(7, 270);
            this.rbHeader.Name = "rbHeader";
            this.rbHeader.Size = new System.Drawing.Size(60, 17);
            this.rbHeader.TabIndex = 9;
            this.rbHeader.TabStop = true;
            this.rbHeader.Text = "Header";
            this.rbHeader.UseVisualStyleBackColor = true;
            this.rbHeader.CheckedChanged += new System.EventHandler(this.rbHeader_CheckedChanged);
            // 
            // rbSubject
            // 
            this.rbSubject.AutoSize = true;
            this.rbSubject.Location = new System.Drawing.Point(7, 220);
            this.rbSubject.Name = "rbSubject";
            this.rbSubject.Size = new System.Drawing.Size(61, 17);
            this.rbSubject.TabIndex = 8;
            this.rbSubject.TabStop = true;
            this.rbSubject.Text = "Subject";
            this.rbSubject.UseVisualStyleBackColor = true;
            this.rbSubject.CheckedChanged += new System.EventHandler(this.rbSubject_CheckedChanged);
            // 
            // rbReplyRecipientName
            // 
            this.rbReplyRecipientName.AutoSize = true;
            this.rbReplyRecipientName.Location = new System.Drawing.Point(7, 170);
            this.rbReplyRecipientName.Name = "rbReplyRecipientName";
            this.rbReplyRecipientName.Size = new System.Drawing.Size(125, 17);
            this.rbReplyRecipientName.TabIndex = 6;
            this.rbReplyRecipientName.TabStop = true;
            this.rbReplyRecipientName.Text = "ReplyRecipientName";
            this.rbReplyRecipientName.UseVisualStyleBackColor = true;
            this.rbReplyRecipientName.CheckedChanged += new System.EventHandler(this.rbReplyRecipientName_CheckedChanged);
            // 
            // rbSentOnBehalfOfName
            // 
            this.rbSentOnBehalfOfName.AutoSize = true;
            this.rbSentOnBehalfOfName.Location = new System.Drawing.Point(7, 120);
            this.rbSentOnBehalfOfName.Name = "rbSentOnBehalfOfName";
            this.rbSentOnBehalfOfName.Size = new System.Drawing.Size(130, 17);
            this.rbSentOnBehalfOfName.TabIndex = 4;
            this.rbSentOnBehalfOfName.TabStop = true;
            this.rbSentOnBehalfOfName.Text = "SentOnBehalfOfName";
            this.rbSentOnBehalfOfName.UseVisualStyleBackColor = true;
            this.rbSentOnBehalfOfName.CheckedChanged += new System.EventHandler(this.rbSentOnBehalfOfName_CheckedChanged);
            // 
            // txSubject
            // 
            this.txSubject.Location = new System.Drawing.Point(7, 243);
            this.txSubject.Name = "txSubject";
            this.txSubject.Size = new System.Drawing.Size(486, 20);
            this.txSubject.TabIndex = 8;
            this.txSubject.TextChanged += new System.EventHandler(this.txSubject_TextChanged);
            // 
            // rbSenderEmailAddress
            // 
            this.rbSenderEmailAddress.AutoSize = true;
            this.rbSenderEmailAddress.Location = new System.Drawing.Point(7, 70);
            this.rbSenderEmailAddress.Name = "rbSenderEmailAddress";
            this.rbSenderEmailAddress.Size = new System.Drawing.Size(122, 17);
            this.rbSenderEmailAddress.TabIndex = 1;
            this.rbSenderEmailAddress.TabStop = true;
            this.rbSenderEmailAddress.Text = "SenderEmailAddress";
            this.rbSenderEmailAddress.UseVisualStyleBackColor = true;
            this.rbSenderEmailAddress.CheckedChanged += new System.EventHandler(this.rbSenderEmailAddress_CheckedChanged);
            // 
            // rbSenderName
            // 
            this.rbSenderName.AutoSize = true;
            this.rbSenderName.Location = new System.Drawing.Point(7, 20);
            this.rbSenderName.Name = "rbSenderName";
            this.rbSenderName.Size = new System.Drawing.Size(87, 17);
            this.rbSenderName.TabIndex = 0;
            this.rbSenderName.TabStop = true;
            this.rbSenderName.Text = "SenderName";
            this.rbSenderName.UseVisualStyleBackColor = true;
            this.rbSenderName.CheckedChanged += new System.EventHandler(this.rbSenderName_CheckedChanged);
            // 
            // txSenderName
            // 
            this.txSenderName.Location = new System.Drawing.Point(7, 43);
            this.txSenderName.Name = "txSenderName";
            this.txSenderName.Size = new System.Drawing.Size(486, 20);
            this.txSenderName.TabIndex = 0;
            this.txSenderName.TextChanged += new System.EventHandler(this.txSenderName_TextChanged);
            // 
            // txReplyRecipientName
            // 
            this.txReplyRecipientName.Location = new System.Drawing.Point(7, 193);
            this.txReplyRecipientName.Name = "txReplyRecipientName";
            this.txReplyRecipientName.Size = new System.Drawing.Size(486, 20);
            this.txReplyRecipientName.TabIndex = 7;
            this.txReplyRecipientName.TextChanged += new System.EventHandler(this.txReplyRecipientName_TextChanged);
            // 
            // txSenderEmailAddress
            // 
            this.txSenderEmailAddress.Location = new System.Drawing.Point(7, 93);
            this.txSenderEmailAddress.Name = "txSenderEmailAddress";
            this.txSenderEmailAddress.Size = new System.Drawing.Size(486, 20);
            this.txSenderEmailAddress.TabIndex = 3;
            this.txSenderEmailAddress.TextChanged += new System.EventHandler(this.txSenderEmailAddress_TextChanged);
            // 
            // txSentOnBehalfOfName
            // 
            this.txSentOnBehalfOfName.Location = new System.Drawing.Point(7, 143);
            this.txSentOnBehalfOfName.Name = "txSentOnBehalfOfName";
            this.txSentOnBehalfOfName.Size = new System.Drawing.Size(486, 20);
            this.txSentOnBehalfOfName.TabIndex = 5;
            this.txSentOnBehalfOfName.TextChanged += new System.EventHandler(this.txSentOnBehalfOfName_TextChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.listBox1);
            this.Name = "Form1";
            this.Text = "Move Rule";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txHeader;
        private System.Windows.Forms.RadioButton rbHeader;
        private System.Windows.Forms.RadioButton rbSubject;
        private System.Windows.Forms.RadioButton rbReplyRecipientName;
        private System.Windows.Forms.RadioButton rbSentOnBehalfOfName;
        private System.Windows.Forms.TextBox txSubject;
        private System.Windows.Forms.RadioButton rbSenderEmailAddress;
        private System.Windows.Forms.RadioButton rbSenderName;
        private System.Windows.Forms.TextBox txSenderName;
        private System.Windows.Forms.TextBox txReplyRecipientName;
        private System.Windows.Forms.TextBox txSenderEmailAddress;
        private System.Windows.Forms.TextBox txSentOnBehalfOfName;
    }
}