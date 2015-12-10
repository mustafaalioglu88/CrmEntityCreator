namespace EntityCreator
{
    partial class MainForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.crmServerText = new System.Windows.Forms.TextBox();
            this.usernameText = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.passwordText = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.excelFileText = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.openFileButton = new System.Windows.Forms.Button();
            this.createEntitiesButton = new System.Windows.Forms.Button();
            this.domainText = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.exportSampleButton = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.statusLabel = new System.Windows.Forms.Label();
            this.openFolderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Crm Server";
            // 
            // crmServerText
            // 
            this.crmServerText.Location = new System.Drawing.Point(89, 10);
            this.crmServerText.Name = "crmServerText";
            this.crmServerText.Size = new System.Drawing.Size(226, 20);
            this.crmServerText.TabIndex = 1;
            this.crmServerText.Text = "http://crm.domain.com/OrganizationName";
            // 
            // usernameText
            // 
            this.usernameText.Location = new System.Drawing.Point(89, 36);
            this.usernameText.Name = "usernameText";
            this.usernameText.Size = new System.Drawing.Size(138, 20);
            this.usernameText.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 39);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(55, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Username";
            // 
            // passwordText
            // 
            this.passwordText.Location = new System.Drawing.Point(89, 62);
            this.passwordText.Name = "passwordText";
            this.passwordText.Size = new System.Drawing.Size(138, 20);
            this.passwordText.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 65);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "Password";
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls";
            // 
            // excelFileText
            // 
            this.excelFileText.Location = new System.Drawing.Point(89, 114);
            this.excelFileText.Name = "excelFileText";
            this.excelFileText.Size = new System.Drawing.Size(104, 20);
            this.excelFileText.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 117);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(52, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Excel File";
            // 
            // openFileButton
            // 
            this.openFileButton.Location = new System.Drawing.Point(199, 113);
            this.openFileButton.Name = "openFileButton";
            this.openFileButton.Size = new System.Drawing.Size(28, 20);
            this.openFileButton.TabIndex = 8;
            this.openFileButton.Text = "...";
            this.openFileButton.UseVisualStyleBackColor = true;
            this.openFileButton.Click += new System.EventHandler(this.OpenFileButton_Click);
            // 
            // createEntitiesButton
            // 
            this.createEntitiesButton.Location = new System.Drawing.Point(233, 36);
            this.createEntitiesButton.Name = "createEntitiesButton";
            this.createEntitiesButton.Size = new System.Drawing.Size(82, 72);
            this.createEntitiesButton.TabIndex = 9;
            this.createEntitiesButton.Text = "Create";
            this.createEntitiesButton.UseVisualStyleBackColor = true;
            this.createEntitiesButton.Click += new System.EventHandler(this.createEntitiesButton_Click);
            // 
            // domainText
            // 
            this.domainText.Location = new System.Drawing.Point(89, 88);
            this.domainText.Name = "domainText";
            this.domainText.Size = new System.Drawing.Size(138, 20);
            this.domainText.TabIndex = 11;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(13, 91);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(43, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Domain";
            // 
            // exportSampleButton
            // 
            this.exportSampleButton.Location = new System.Drawing.Point(233, 117);
            this.exportSampleButton.Name = "exportSampleButton";
            this.exportSampleButton.Size = new System.Drawing.Size(82, 43);
            this.exportSampleButton.TabIndex = 12;
            this.exportSampleButton.Text = "Export Sample";
            this.exportSampleButton.UseVisualStyleBackColor = true;
            this.exportSampleButton.Click += new System.EventHandler(this.exportSampleButton_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(13, 143);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(37, 13);
            this.label6.TabIndex = 13;
            this.label6.Text = "Status";
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(86, 143);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(13, 13);
            this.statusLabel.TabIndex = 14;
            this.statusLabel.Text = ":(";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(12, 166);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(303, 23);
            this.progressBar.TabIndex = 15;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(327, 197);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.exportSampleButton);
            this.Controls.Add(this.domainText);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.createEntitiesButton);
            this.Controls.Add(this.openFileButton);
            this.Controls.Add(this.excelFileText);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.passwordText);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.usernameText);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.crmServerText);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "MainForm";
            this.Text = "Entity Creator";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox crmServerText;
        private System.Windows.Forms.TextBox usernameText;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox passwordText;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.TextBox excelFileText;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button openFileButton;
        private System.Windows.Forms.Button createEntitiesButton;
        private System.Windows.Forms.TextBox domainText;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button exportSampleButton;
        private System.Windows.Forms.Label label6;
        internal System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.FolderBrowserDialog openFolderBrowserDialog;
        private System.Windows.Forms.ProgressBar progressBar;
    }
}

