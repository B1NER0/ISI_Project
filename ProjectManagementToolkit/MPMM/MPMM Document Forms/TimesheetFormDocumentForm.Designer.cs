﻿namespace ProjectManagementToolkit.MPMM.MPMM_Document_Forms
{
    partial class TimesheetFormDocumentForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle6 = new System.Windows.Forms.DataGridViewCellStyle();
            this.lblProjectName = new System.Windows.Forms.Label();
            this.txtProjectName = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.dateTimePickerApprovedBy = new System.Windows.Forms.DateTimePicker();
            this.txtApprovedBySignature = new System.Windows.Forms.TextBox();
            this.txtApprovedByProjectRole = new System.Windows.Forms.TextBox();
            this.txtApprovedByName = new System.Windows.Forms.TextBox();
            this.lblApprovedBy = new System.Windows.Forms.Label();
            this.dateTimePickerSubmittedBy = new System.Windows.Forms.DateTimePicker();
            this.txtSignature = new System.Windows.Forms.TextBox();
            this.txtProjectRole = new System.Windows.Forms.TextBox();
            this.txtName = new System.Windows.Forms.TextBox();
            this.lblSubmittedBy = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.lblTimesheetProjectName = new System.Windows.Forms.Label();
            this.dataGridViewTimesheetForm = new System.Windows.Forms.DataGridView();
            this.Date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.StartTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.EndTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Duration = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Activity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Task = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.StartComplete = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.EndComplete = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Result = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtTimesheetFormTeamMember = new System.Windows.Forms.TextBox();
            this.txtTimesheetFormProjectManager = new System.Windows.Forms.TextBox();
            this.txtTimesheetFormProjectName = new System.Windows.Forms.TextBox();
            this.btnSaveProjectName = new System.Windows.Forms.Button();
            this.btnExportToWord = new System.Windows.Forms.Button();
            this.btnSaveProgress = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewTimesheetForm)).BeginInit();
            this.SuspendLayout();
            // 
            // lblProjectName
            // 
            this.lblProjectName.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblProjectName.AutoSize = true;
            this.lblProjectName.Location = new System.Drawing.Point(15, 15);
            this.lblProjectName.Name = "lblProjectName";
            this.lblProjectName.Size = new System.Drawing.Size(239, 16);
            this.lblProjectName.TabIndex = 0;
            this.lblProjectName.Text = "Please Enter Your Project Name: ";
            // 
            // txtProjectName
            // 
            this.txtProjectName.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtProjectName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.txtProjectName.ForeColor = System.Drawing.Color.Black;
            this.txtProjectName.Location = new System.Drawing.Point(209, 13);
            this.txtProjectName.Name = "txtProjectName";
            this.txtProjectName.Size = new System.Drawing.Size(116, 23);
            this.txtProjectName.TabIndex = 1;
            this.txtProjectName.Text = "Project Name";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.dateTimePickerApprovedBy);
            this.groupBox1.Controls.Add(this.txtApprovedBySignature);
            this.groupBox1.Controls.Add(this.txtApprovedByProjectRole);
            this.groupBox1.Controls.Add(this.txtApprovedByName);
            this.groupBox1.Controls.Add(this.lblApprovedBy);
            this.groupBox1.Controls.Add(this.dateTimePickerSubmittedBy);
            this.groupBox1.Controls.Add(this.txtSignature);
            this.groupBox1.Controls.Add(this.txtProjectRole);
            this.groupBox1.Controls.Add(this.txtName);
            this.groupBox1.Controls.Add(this.lblSubmittedBy);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.lblTimesheetProjectName);
            this.groupBox1.Controls.Add(this.dataGridViewTimesheetForm);
            this.groupBox1.Controls.Add(this.txtTimesheetFormTeamMember);
            this.groupBox1.Controls.Add(this.txtTimesheetFormProjectManager);
            this.groupBox1.Controls.Add(this.txtTimesheetFormProjectName);
            this.groupBox1.ForeColor = System.Drawing.Color.Black;
            this.groupBox1.Location = new System.Drawing.Point(17, 42);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(902, 428);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "TIMESHEET FORM";
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 7F);
            this.label3.Location = new System.Drawing.Point(350, 17);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(564, 15);
            this.label3.TabIndex = 20;
            this.label3.Text = "ONCE COMPLETED, PLEASE FORWARD THIS FORM TO THE PROJECT MANAGER FOR APPROVAL";
            // 
            // dateTimePickerApprovedBy
            // 
            this.dateTimePickerApprovedBy.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dateTimePickerApprovedBy.Location = new System.Drawing.Point(664, 374);
            this.dateTimePickerApprovedBy.Name = "dateTimePickerApprovedBy";
            this.dateTimePickerApprovedBy.Size = new System.Drawing.Size(233, 23);
            this.dateTimePickerApprovedBy.TabIndex = 19;
            // 
            // txtApprovedBySignature
            // 
            this.txtApprovedBySignature.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtApprovedBySignature.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.txtApprovedBySignature.ForeColor = System.Drawing.Color.Black;
            this.txtApprovedBySignature.Location = new System.Drawing.Point(540, 374);
            this.txtApprovedBySignature.Name = "txtApprovedBySignature";
            this.txtApprovedBySignature.Size = new System.Drawing.Size(116, 23);
            this.txtApprovedBySignature.TabIndex = 18;
            this.txtApprovedBySignature.Text = "Signature";
            // 
            // txtApprovedByProjectRole
            // 
            this.txtApprovedByProjectRole.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtApprovedByProjectRole.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.txtApprovedByProjectRole.ForeColor = System.Drawing.Color.Black;
            this.txtApprovedByProjectRole.Location = new System.Drawing.Point(540, 344);
            this.txtApprovedByProjectRole.Name = "txtApprovedByProjectRole";
            this.txtApprovedByProjectRole.Size = new System.Drawing.Size(116, 23);
            this.txtApprovedByProjectRole.TabIndex = 17;
            this.txtApprovedByProjectRole.Text = "Project Role";
            // 
            // txtApprovedByName
            // 
            this.txtApprovedByName.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtApprovedByName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.txtApprovedByName.ForeColor = System.Drawing.Color.Black;
            this.txtApprovedByName.Location = new System.Drawing.Point(540, 314);
            this.txtApprovedByName.Name = "txtApprovedByName";
            this.txtApprovedByName.Size = new System.Drawing.Size(116, 23);
            this.txtApprovedByName.TabIndex = 16;
            this.txtApprovedByName.Text = "Name";
            // 
            // lblApprovedBy
            // 
            this.lblApprovedBy.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblApprovedBy.AutoSize = true;
            this.lblApprovedBy.Location = new System.Drawing.Point(448, 314);
            this.lblApprovedBy.Name = "lblApprovedBy";
            this.lblApprovedBy.Size = new System.Drawing.Size(106, 16);
            this.lblApprovedBy.TabIndex = 15;
            this.lblApprovedBy.Text = "Approved By: ";
            // 
            // dateTimePickerSubmittedBy
            // 
            this.dateTimePickerSubmittedBy.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dateTimePickerSubmittedBy.Location = new System.Drawing.Point(226, 378);
            this.dateTimePickerSubmittedBy.Name = "dateTimePickerSubmittedBy";
            this.dateTimePickerSubmittedBy.Size = new System.Drawing.Size(233, 23);
            this.dateTimePickerSubmittedBy.TabIndex = 14;
            // 
            // txtSignature
            // 
            this.txtSignature.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtSignature.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.txtSignature.ForeColor = System.Drawing.Color.Black;
            this.txtSignature.Location = new System.Drawing.Point(103, 378);
            this.txtSignature.Name = "txtSignature";
            this.txtSignature.Size = new System.Drawing.Size(116, 23);
            this.txtSignature.TabIndex = 13;
            this.txtSignature.Text = "Signature";
            // 
            // txtProjectRole
            // 
            this.txtProjectRole.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtProjectRole.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.txtProjectRole.ForeColor = System.Drawing.Color.Black;
            this.txtProjectRole.Location = new System.Drawing.Point(103, 348);
            this.txtProjectRole.Name = "txtProjectRole";
            this.txtProjectRole.Size = new System.Drawing.Size(116, 23);
            this.txtProjectRole.TabIndex = 12;
            this.txtProjectRole.Text = "Project Role";
            // 
            // txtName
            // 
            this.txtName.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.txtName.ForeColor = System.Drawing.Color.Black;
            this.txtName.Location = new System.Drawing.Point(103, 317);
            this.txtName.Name = "txtName";
            this.txtName.Size = new System.Drawing.Size(116, 23);
            this.txtName.TabIndex = 11;
            this.txtName.Text = "Name";
            // 
            // lblSubmittedBy
            // 
            this.lblSubmittedBy.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblSubmittedBy.AutoSize = true;
            this.lblSubmittedBy.Location = new System.Drawing.Point(10, 317);
            this.lblSubmittedBy.Name = "lblSubmittedBy";
            this.lblSubmittedBy.Size = new System.Drawing.Size(108, 16);
            this.lblSubmittedBy.TabIndex = 10;
            this.lblSubmittedBy.Text = "Submitted By: ";
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 86);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(244, 16);
            this.label2.TabIndex = 9;
            this.label2.Text = "Enter Your Team Members Name: ";
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 56);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(262, 16);
            this.label1.TabIndex = 8;
            this.label1.Text = "Enter Your Project Managers Name: ";
            // 
            // lblTimesheetProjectName
            // 
            this.lblTimesheetProjectName.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblTimesheetProjectName.AutoSize = true;
            this.lblTimesheetProjectName.Location = new System.Drawing.Point(8, 26);
            this.lblTimesheetProjectName.Name = "lblTimesheetProjectName";
            this.lblTimesheetProjectName.Size = new System.Drawing.Size(188, 16);
            this.lblTimesheetProjectName.TabIndex = 7;
            this.lblTimesheetProjectName.Text = "Enter Your Project Name: ";
            // 
            // dataGridViewTimesheetForm
            // 
            this.dataGridViewTimesheetForm.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewTimesheetForm.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridViewTimesheetForm.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewTimesheetForm.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.dataGridViewTimesheetForm.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewTimesheetForm.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Date,
            this.StartTime,
            this.EndTime,
            this.Duration,
            this.Activity,
            this.Task,
            this.StartComplete,
            this.EndComplete,
            this.Result});
            this.dataGridViewTimesheetForm.EnableHeadersVisualStyles = false;
            this.dataGridViewTimesheetForm.Location = new System.Drawing.Point(7, 113);
            this.dataGridViewTimesheetForm.Name = "dataGridViewTimesheetForm";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewTimesheetForm.RowHeadersDefaultCellStyle = dataGridViewCellStyle5;
            this.dataGridViewTimesheetForm.RowHeadersWidth = 51;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Arial", 10.8F);
            this.dataGridViewTimesheetForm.RowsDefaultCellStyle = dataGridViewCellStyle6;
            this.dataGridViewTimesheetForm.Size = new System.Drawing.Size(888, 196);
            this.dataGridViewTimesheetForm.TabIndex = 6;
            // 
            // Date
            // 
            this.Date.HeaderText = "Date";
            this.Date.MinimumWidth = 6;
            this.Date.Name = "Date";
            // 
            // StartTime
            // 
            this.StartTime.HeaderText = "StartTime";
            this.StartTime.MinimumWidth = 6;
            this.StartTime.Name = "StartTime";
            // 
            // EndTime
            // 
            this.EndTime.HeaderText = "End Time";
            this.EndTime.MinimumWidth = 6;
            this.EndTime.Name = "EndTime";
            // 
            // Duration
            // 
            this.Duration.HeaderText = "Duration";
            this.Duration.MinimumWidth = 6;
            this.Duration.Name = "Duration";
            // 
            // Activity
            // 
            this.Activity.HeaderText = "Activity";
            this.Activity.MinimumWidth = 6;
            this.Activity.Name = "Activity";
            // 
            // Task
            // 
            this.Task.HeaderText = "Task";
            this.Task.MinimumWidth = 6;
            this.Task.Name = "Task";
            // 
            // StartComplete
            // 
            this.StartComplete.HeaderText = "Start % Complete";
            this.StartComplete.MinimumWidth = 6;
            this.StartComplete.Name = "StartComplete";
            // 
            // EndComplete
            // 
            this.EndComplete.HeaderText = "End % Complete";
            this.EndComplete.MinimumWidth = 6;
            this.EndComplete.Name = "EndComplete";
            // 
            // Result
            // 
            this.Result.HeaderText = "Result";
            this.Result.MinimumWidth = 6;
            this.Result.Name = "Result";
            // 
            // txtTimesheetFormTeamMember
            // 
            this.txtTimesheetFormTeamMember.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtTimesheetFormTeamMember.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.txtTimesheetFormTeamMember.ForeColor = System.Drawing.Color.Black;
            this.txtTimesheetFormTeamMember.Location = new System.Drawing.Point(226, 84);
            this.txtTimesheetFormTeamMember.Name = "txtTimesheetFormTeamMember";
            this.txtTimesheetFormTeamMember.Size = new System.Drawing.Size(116, 23);
            this.txtTimesheetFormTeamMember.TabIndex = 5;
            this.txtTimesheetFormTeamMember.Text = "Team Member";
            // 
            // txtTimesheetFormProjectManager
            // 
            this.txtTimesheetFormProjectManager.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtTimesheetFormProjectManager.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.txtTimesheetFormProjectManager.ForeColor = System.Drawing.Color.Black;
            this.txtTimesheetFormProjectManager.Location = new System.Drawing.Point(226, 53);
            this.txtTimesheetFormProjectManager.Name = "txtTimesheetFormProjectManager";
            this.txtTimesheetFormProjectManager.Size = new System.Drawing.Size(116, 23);
            this.txtTimesheetFormProjectManager.TabIndex = 4;
            this.txtTimesheetFormProjectManager.Text = "Project Manager";
            // 
            // txtTimesheetFormProjectName
            // 
            this.txtTimesheetFormProjectName.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtTimesheetFormProjectName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.txtTimesheetFormProjectName.ForeColor = System.Drawing.Color.Black;
            this.txtTimesheetFormProjectName.Location = new System.Drawing.Point(226, 23);
            this.txtTimesheetFormProjectName.Name = "txtTimesheetFormProjectName";
            this.txtTimesheetFormProjectName.Size = new System.Drawing.Size(116, 23);
            this.txtTimesheetFormProjectName.TabIndex = 3;
            this.txtTimesheetFormProjectName.Text = "Project Name";
            // 
            // btnSaveProjectName
            // 
            this.btnSaveProjectName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnSaveProjectName.ForeColor = System.Drawing.Color.White;
            this.btnSaveProjectName.Location = new System.Drawing.Point(520, 9);
            this.btnSaveProjectName.Name = "btnSaveProjectName";
            this.btnSaveProjectName.Size = new System.Drawing.Size(122, 27);
            this.btnSaveProjectName.TabIndex = 3;
            this.btnSaveProjectName.Text = "Complete";
            this.btnSaveProjectName.UseVisualStyleBackColor = false;
            this.btnSaveProjectName.Click += new System.EventHandler(this.btnSaveProjectName_Click);
            // 
            // btnExportToWord
            // 
            this.btnExportToWord.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnExportToWord.ForeColor = System.Drawing.Color.White;
            this.btnExportToWord.Location = new System.Drawing.Point(804, 9);
            this.btnExportToWord.Name = "btnExportToWord";
            this.btnExportToWord.Size = new System.Drawing.Size(110, 27);
            this.btnExportToWord.TabIndex = 22;
            this.btnExportToWord.Text = "Export to Word";
            this.btnExportToWord.UseVisualStyleBackColor = false;
            this.btnExportToWord.Click += new System.EventHandler(this.btnExportToWord_Click);
            // 
            // btnSaveProgress
            // 
            this.btnSaveProgress.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnSaveProgress.ForeColor = System.Drawing.Color.White;
            this.btnSaveProgress.Location = new System.Drawing.Point(667, 9);
            this.btnSaveProgress.Name = "btnSaveProgress";
            this.btnSaveProgress.Size = new System.Drawing.Size(122, 27);
            this.btnSaveProgress.TabIndex = 23;
            this.btnSaveProgress.Text = "Save Progress";
            this.btnSaveProgress.UseVisualStyleBackColor = false;
            this.btnSaveProgress.Click += new System.EventHandler(this.btnSaveProgress_Click);
            // 
            // TimesheetFormDocumentForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.ClientSize = new System.Drawing.Size(933, 484);
            this.Controls.Add(this.btnSaveProgress);
            this.Controls.Add(this.btnExportToWord);
            this.Controls.Add(this.btnSaveProjectName);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.txtProjectName);
            this.Controls.Add(this.lblProjectName);
            this.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Black;
            this.Name = "TimesheetFormDocumentForm";
            this.Text = "TimesheetFormDocumentForm";
            this.Load += new System.EventHandler(this.TimesheetFormDocumentForm_Load_1);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewTimesheetForm)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblProjectName;
        private System.Windows.Forms.TextBox txtProjectName;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dateTimePickerApprovedBy;
        private System.Windows.Forms.TextBox txtApprovedBySignature;
        private System.Windows.Forms.TextBox txtApprovedByProjectRole;
        private System.Windows.Forms.TextBox txtApprovedByName;
        private System.Windows.Forms.Label lblApprovedBy;
        private System.Windows.Forms.DateTimePicker dateTimePickerSubmittedBy;
        private System.Windows.Forms.TextBox txtSignature;
        private System.Windows.Forms.TextBox txtProjectRole;
        private System.Windows.Forms.TextBox txtName;
        private System.Windows.Forms.Label lblSubmittedBy;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblTimesheetProjectName;
        private System.Windows.Forms.DataGridView dataGridViewTimesheetForm;
        private System.Windows.Forms.TextBox txtTimesheetFormTeamMember;
        private System.Windows.Forms.TextBox txtTimesheetFormProjectManager;
        private System.Windows.Forms.TextBox txtTimesheetFormProjectName;
        private System.Windows.Forms.Button btnSaveProjectName;
        private System.Windows.Forms.Button btnExportToWord;
        private System.Windows.Forms.DataGridViewTextBoxColumn Date;
        private System.Windows.Forms.DataGridViewTextBoxColumn StartTime;
        private System.Windows.Forms.DataGridViewTextBoxColumn EndTime;
        private System.Windows.Forms.DataGridViewTextBoxColumn Duration;
        private System.Windows.Forms.DataGridViewTextBoxColumn Activity;
        private System.Windows.Forms.DataGridViewTextBoxColumn Task;
        private System.Windows.Forms.DataGridViewTextBoxColumn StartComplete;
        private System.Windows.Forms.DataGridViewTextBoxColumn EndComplete;
        private System.Windows.Forms.DataGridViewTextBoxColumn Result;
        private System.Windows.Forms.Button btnSaveProgress;
    }
}