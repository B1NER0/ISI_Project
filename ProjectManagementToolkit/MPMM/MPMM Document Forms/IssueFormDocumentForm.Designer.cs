﻿namespace ProjectManagementToolkit.MPMM.MPMM_Document_Forms
{
    partial class IssueFormDocumentForm
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
            this.tbpIssueDescription = new System.Windows.Forms.TabPage();
            this.txtIssueDescription = new System.Windows.Forms.TextBox();
            this.tbcQualityReviewForm = new System.Windows.Forms.TabControl();
            this.tbpInformation = new System.Windows.Forms.TabPage();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.txtRaisedBy = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtIssueID = new System.Windows.Forms.TextBox();
            this.txtProjectManagerName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbpIssueImpact = new System.Windows.Forms.TabPage();
            this.txtIssueImpact = new System.Windows.Forms.TextBox();
            this.tbpIssueResolution = new System.Windows.Forms.TabPage();
            this.txtIssueResolution = new System.Windows.Forms.TextBox();
            this.tbpApprovalDetails = new System.Windows.Forms.TabPage();
            this.label8 = new System.Windows.Forms.Label();
            this.pnlSupportingDocumentation = new System.Windows.Forms.Panel();
            this.txtSupportingDocumentation = new System.Windows.Forms.TextBox();
            this.btnSendEmail = new System.Windows.Forms.Button();
            this.btnIssueSign = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtSignature = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtDate = new System.Windows.Forms.TextBox();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnNewForm = new System.Windows.Forms.Button();
            this.cmbIssueForms = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.tbpIssueDescription.SuspendLayout();
            this.tbcQualityReviewForm.SuspendLayout();
            this.tbpInformation.SuspendLayout();
            this.tbpIssueImpact.SuspendLayout();
            this.tbpIssueResolution.SuspendLayout();
            this.tbpApprovalDetails.SuspendLayout();
            this.pnlSupportingDocumentation.SuspendLayout();
            this.SuspendLayout();
            // 
            // tbpIssueDescription
            // 
            this.tbpIssueDescription.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.tbpIssueDescription.Controls.Add(this.txtIssueDescription);
            this.tbpIssueDescription.Location = new System.Drawing.Point(4, 25);
            this.tbpIssueDescription.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.tbpIssueDescription.Name = "tbpIssueDescription";
            this.tbpIssueDescription.Size = new System.Drawing.Size(1293, 575);
            this.tbpIssueDescription.TabIndex = 3;
            this.tbpIssueDescription.Text = "Issue Description";
            // 
            // txtIssueDescription
            // 
            this.txtIssueDescription.BackColor = System.Drawing.SystemColors.Control;
            this.txtIssueDescription.Location = new System.Drawing.Point(14, 13);
            this.txtIssueDescription.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.txtIssueDescription.Multiline = true;
            this.txtIssueDescription.Name = "txtIssueDescription";
            this.txtIssueDescription.Size = new System.Drawing.Size(1265, 555);
            this.txtIssueDescription.TabIndex = 6;
            // 
            // tbcQualityReviewForm
            // 
            this.tbcQualityReviewForm.AllowDrop = true;
            this.tbcQualityReviewForm.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbcQualityReviewForm.Controls.Add(this.tbpInformation);
            this.tbcQualityReviewForm.Controls.Add(this.tbpIssueDescription);
            this.tbcQualityReviewForm.Controls.Add(this.tbpIssueImpact);
            this.tbcQualityReviewForm.Controls.Add(this.tbpIssueResolution);
            this.tbcQualityReviewForm.Controls.Add(this.tbpApprovalDetails);
            this.tbcQualityReviewForm.Location = new System.Drawing.Point(16, 63);
            this.tbcQualityReviewForm.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.tbcQualityReviewForm.Name = "tbcQualityReviewForm";
            this.tbcQualityReviewForm.SelectedIndex = 0;
            this.tbcQualityReviewForm.Size = new System.Drawing.Size(1301, 604);
            this.tbcQualityReviewForm.TabIndex = 18;
            // 
            // tbpInformation
            // 
            this.tbpInformation.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.tbpInformation.Controls.Add(this.dateTimePicker1);
            this.tbpInformation.Controls.Add(this.txtRaisedBy);
            this.tbpInformation.Controls.Add(this.label4);
            this.tbpInformation.Controls.Add(this.label3);
            this.tbpInformation.Controls.Add(this.label2);
            this.tbpInformation.Controls.Add(this.txtIssueID);
            this.tbpInformation.Controls.Add(this.txtProjectManagerName);
            this.tbpInformation.Controls.Add(this.label1);
            this.tbpInformation.Location = new System.Drawing.Point(4, 25);
            this.tbpInformation.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.tbpInformation.Name = "tbpInformation";
            this.tbpInformation.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.tbpInformation.Size = new System.Drawing.Size(1293, 575);
            this.tbpInformation.TabIndex = 7;
            this.tbpInformation.Text = "Information";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(17, 170);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(396, 23);
            this.dateTimePicker1.TabIndex = 28;
            // 
            // txtRaisedBy
            // 
            this.txtRaisedBy.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.txtRaisedBy.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRaisedBy.ForeColor = System.Drawing.Color.Black;
            this.txtRaisedBy.Location = new System.Drawing.Point(17, 124);
            this.txtRaisedBy.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.txtRaisedBy.Name = "txtRaisedBy";
            this.txtRaisedBy.Size = new System.Drawing.Size(396, 23);
            this.txtRaisedBy.TabIndex = 23;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.ForeColor = System.Drawing.Color.Black;
            this.label4.Location = new System.Drawing.Point(14, 153);
            this.label4.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 16);
            this.label4.TabIndex = 26;
            this.label4.Text = "Date Raised:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.ForeColor = System.Drawing.Color.Black;
            this.label3.Location = new System.Drawing.Point(14, 59);
            this.label3.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(68, 16);
            this.label3.TabIndex = 22;
            this.label3.Text = "Issue ID:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(14, 106);
            this.label2.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(81, 16);
            this.label2.TabIndex = 24;
            this.label2.Text = "Raised By:";
            // 
            // txtIssueID
            // 
            this.txtIssueID.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.txtIssueID.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtIssueID.ForeColor = System.Drawing.Color.Black;
            this.txtIssueID.Location = new System.Drawing.Point(17, 77);
            this.txtIssueID.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.txtIssueID.Name = "txtIssueID";
            this.txtIssueID.Size = new System.Drawing.Size(396, 23);
            this.txtIssueID.TabIndex = 21;
            // 
            // txtProjectManagerName
            // 
            this.txtProjectManagerName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.txtProjectManagerName.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtProjectManagerName.ForeColor = System.Drawing.Color.Black;
            this.txtProjectManagerName.Location = new System.Drawing.Point(17, 30);
            this.txtProjectManagerName.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.txtProjectManagerName.Name = "txtProjectManagerName";
            this.txtProjectManagerName.Size = new System.Drawing.Size(396, 23);
            this.txtProjectManagerName.TabIndex = 27;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(14, 12);
            this.label1.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(264, 16);
            this.label1.TabIndex = 20;
            this.label1.Text = "Please Enter Project Manager Name:";
            // 
            // tbpIssueImpact
            // 
            this.tbpIssueImpact.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.tbpIssueImpact.Controls.Add(this.txtIssueImpact);
            this.tbpIssueImpact.Location = new System.Drawing.Point(4, 25);
            this.tbpIssueImpact.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.tbpIssueImpact.Name = "tbpIssueImpact";
            this.tbpIssueImpact.Size = new System.Drawing.Size(1293, 575);
            this.tbpIssueImpact.TabIndex = 4;
            this.tbpIssueImpact.Text = "Issue Impact";
            // 
            // txtIssueImpact
            // 
            this.txtIssueImpact.BackColor = System.Drawing.SystemColors.Control;
            this.txtIssueImpact.Location = new System.Drawing.Point(14, 14);
            this.txtIssueImpact.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.txtIssueImpact.Multiline = true;
            this.txtIssueImpact.Name = "txtIssueImpact";
            this.txtIssueImpact.Size = new System.Drawing.Size(1262, 549);
            this.txtIssueImpact.TabIndex = 6;
            // 
            // tbpIssueResolution
            // 
            this.tbpIssueResolution.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.tbpIssueResolution.Controls.Add(this.txtIssueResolution);
            this.tbpIssueResolution.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.tbpIssueResolution.Location = new System.Drawing.Point(4, 25);
            this.tbpIssueResolution.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.tbpIssueResolution.Name = "tbpIssueResolution";
            this.tbpIssueResolution.Size = new System.Drawing.Size(1293, 575);
            this.tbpIssueResolution.TabIndex = 5;
            this.tbpIssueResolution.Text = "Issue Resolution";
            // 
            // txtIssueResolution
            // 
            this.txtIssueResolution.BackColor = System.Drawing.SystemColors.Control;
            this.txtIssueResolution.Location = new System.Drawing.Point(18, 17);
            this.txtIssueResolution.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.txtIssueResolution.Multiline = true;
            this.txtIssueResolution.Name = "txtIssueResolution";
            this.txtIssueResolution.Size = new System.Drawing.Size(1256, 545);
            this.txtIssueResolution.TabIndex = 6;
            // 
            // tbpApprovalDetails
            // 
            this.tbpApprovalDetails.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.tbpApprovalDetails.Controls.Add(this.label8);
            this.tbpApprovalDetails.Controls.Add(this.pnlSupportingDocumentation);
            this.tbpApprovalDetails.Controls.Add(this.btnSendEmail);
            this.tbpApprovalDetails.Controls.Add(this.btnIssueSign);
            this.tbpApprovalDetails.Controls.Add(this.label7);
            this.tbpApprovalDetails.Controls.Add(this.label6);
            this.tbpApprovalDetails.Controls.Add(this.txtSignature);
            this.tbpApprovalDetails.Controls.Add(this.label5);
            this.tbpApprovalDetails.Controls.Add(this.txtDate);
            this.tbpApprovalDetails.Location = new System.Drawing.Point(4, 25);
            this.tbpApprovalDetails.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.tbpApprovalDetails.Name = "tbpApprovalDetails";
            this.tbpApprovalDetails.Size = new System.Drawing.Size(1293, 575);
            this.tbpApprovalDetails.TabIndex = 6;
            this.tbpApprovalDetails.Text = "Approval Details";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label8.Location = new System.Drawing.Point(16, 151);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(200, 16);
            this.label8.TabIndex = 31;
            this.label8.Text = "Supporting Documentation:";
            // 
            // pnlSupportingDocumentation
            // 
            this.pnlSupportingDocumentation.BackColor = System.Drawing.Color.Silver;
            this.pnlSupportingDocumentation.Controls.Add(this.txtSupportingDocumentation);
            this.pnlSupportingDocumentation.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.pnlSupportingDocumentation.Location = new System.Drawing.Point(19, 169);
            this.pnlSupportingDocumentation.Name = "pnlSupportingDocumentation";
            this.pnlSupportingDocumentation.Size = new System.Drawing.Size(1252, 387);
            this.pnlSupportingDocumentation.TabIndex = 30;
            // 
            // txtSupportingDocumentation
            // 
            this.txtSupportingDocumentation.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtSupportingDocumentation.Location = new System.Drawing.Point(0, 0);
            this.txtSupportingDocumentation.Multiline = true;
            this.txtSupportingDocumentation.Name = "txtSupportingDocumentation";
            this.txtSupportingDocumentation.Size = new System.Drawing.Size(1252, 387);
            this.txtSupportingDocumentation.TabIndex = 0;
            // 
            // btnSendEmail
            // 
            this.btnSendEmail.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnSendEmail.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSendEmail.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnSendEmail.Location = new System.Drawing.Point(285, 71);
            this.btnSendEmail.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnSendEmail.Name = "btnSendEmail";
            this.btnSendEmail.Size = new System.Drawing.Size(228, 29);
            this.btnSendEmail.TabIndex = 29;
            this.btnSendEmail.Text = "Send Email";
            this.btnSendEmail.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnSendEmail.UseVisualStyleBackColor = false;
            // 
            // btnIssueSign
            // 
            this.btnIssueSign.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnIssueSign.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnIssueSign.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnIssueSign.Location = new System.Drawing.Point(19, 71);
            this.btnIssueSign.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnIssueSign.Name = "btnIssueSign";
            this.btnIssueSign.Size = new System.Drawing.Size(223, 29);
            this.btnIssueSign.TabIndex = 28;
            this.btnIssueSign.Text = "Sign";
            this.btnIssueSign.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnIssueSign.UseVisualStyleBackColor = false;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Cambria", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.ForeColor = System.Drawing.Color.Red;
            this.label7.Location = new System.Drawing.Point(17, 118);
            this.label7.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(452, 16);
            this.label7.TabIndex = 27;
            this.label7.Text = "PLEASE FORWARD THIS FORM TO THE PROJECT MANAGER FOR ACTION\r\n";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label6.Location = new System.Drawing.Point(16, 16);
            this.label6.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 16);
            this.label6.TabIndex = 26;
            this.label6.Text = "Signature:";
            // 
            // txtSignature
            // 
            this.txtSignature.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.txtSignature.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSignature.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.txtSignature.Location = new System.Drawing.Point(19, 34);
            this.txtSignature.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.txtSignature.Name = "txtSignature";
            this.txtSignature.Size = new System.Drawing.Size(227, 23);
            this.txtSignature.TabIndex = 25;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label5.Location = new System.Drawing.Point(282, 16);
            this.label5.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(44, 16);
            this.label5.TabIndex = 24;
            this.label5.Text = "Date:";
            // 
            // txtDate
            // 
            this.txtDate.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.txtDate.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtDate.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.txtDate.Location = new System.Drawing.Point(285, 34);
            this.txtDate.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(227, 23);
            this.txtDate.TabIndex = 23;
            // 
            // btnExport
            // 
            this.btnExport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnExport.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExport.ForeColor = System.Drawing.Color.Black;
            this.btnExport.Location = new System.Drawing.Point(956, 12);
            this.btnExport.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(130, 44);
            this.btnExport.TabIndex = 29;
            this.btnExport.Text = "Export Current Form to Word";
            this.btnExport.UseVisualStyleBackColor = false;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSave.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSave.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnSave.Location = new System.Drawing.Point(198, 12);
            this.btnSave.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(130, 44);
            this.btnSave.TabIndex = 28;
            this.btnSave.Text = "Complete Current Form";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnDelete.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnDelete.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnDelete.Location = new System.Drawing.Point(351, 12);
            this.btnDelete.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(130, 44);
            this.btnDelete.TabIndex = 30;
            this.btnDelete.Text = "Delete Current Form";
            this.btnDelete.UseVisualStyleBackColor = false;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnNewForm
            // 
            this.btnNewForm.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnNewForm.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnNewForm.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNewForm.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.btnNewForm.Location = new System.Drawing.Point(502, 12);
            this.btnNewForm.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnNewForm.Name = "btnNewForm";
            this.btnNewForm.Size = new System.Drawing.Size(130, 44);
            this.btnNewForm.TabIndex = 31;
            this.btnNewForm.Text = "Add New Form";
            this.btnNewForm.UseVisualStyleBackColor = false;
            this.btnNewForm.Click += new System.EventHandler(this.btnNewForm_Click);
            // 
            // cmbIssueForms
            // 
            this.cmbIssueForms.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbIssueForms.FormattingEnabled = true;
            this.cmbIssueForms.Location = new System.Drawing.Point(639, 12);
            this.cmbIssueForms.Name = "cmbIssueForms";
            this.cmbIssueForms.Size = new System.Drawing.Size(310, 24);
            this.cmbIssueForms.TabIndex = 32;
            this.cmbIssueForms.SelectedIndexChanged += new System.EventHandler(this.cmbIssueForms_SelectedIndexChanged);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.button1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.button1.Location = new System.Drawing.Point(51, 12);
            this.button1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(130, 44);
            this.button1.TabIndex = 28;
            this.button1.Text = "Save Current Form Progress";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // IssueFormDocumentForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.ClientSize = new System.Drawing.Size(1332, 687);
            this.Controls.Add(this.cmbIssueForms);
            this.Controls.Add(this.btnNewForm);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.tbcQualityReviewForm);
            this.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.Name = "IssueFormDocumentForm";
            this.Text = "Add New Form";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.IssueFormDocumentForm_Load);
            this.tbpIssueDescription.ResumeLayout(false);
            this.tbpIssueDescription.PerformLayout();
            this.tbcQualityReviewForm.ResumeLayout(false);
            this.tbpInformation.ResumeLayout(false);
            this.tbpInformation.PerformLayout();
            this.tbpIssueImpact.ResumeLayout(false);
            this.tbpIssueImpact.PerformLayout();
            this.tbpIssueResolution.ResumeLayout(false);
            this.tbpIssueResolution.PerformLayout();
            this.tbpApprovalDetails.ResumeLayout(false);
            this.tbpApprovalDetails.PerformLayout();
            this.pnlSupportingDocumentation.ResumeLayout(false);
            this.pnlSupportingDocumentation.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl tbcQualityReviewForm;
        private System.Windows.Forms.TabPage tbpIssueImpact;
        private System.Windows.Forms.TabPage tbpIssueResolution;
        private System.Windows.Forms.TabPage tbpApprovalDetails;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtRaisedBy;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtIssueID;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtProjectManagerName;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtSignature;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtDate;
        private System.Windows.Forms.TabPage tbpIssueDescription;
        private System.Windows.Forms.Button btnSendEmail;
        private System.Windows.Forms.Button btnIssueSign;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TabPage tbpInformation;
        private System.Windows.Forms.TextBox txtIssueDescription;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.TextBox txtIssueImpact;
        private System.Windows.Forms.TextBox txtIssueResolution;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Panel pnlSupportingDocumentation;
        private System.Windows.Forms.TextBox txtSupportingDocumentation;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnNewForm;
        private System.Windows.Forms.ComboBox cmbIssueForms;
        private System.Windows.Forms.Button button1;
    }
}