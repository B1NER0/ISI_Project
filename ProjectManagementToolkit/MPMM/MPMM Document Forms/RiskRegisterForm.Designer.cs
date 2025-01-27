﻿namespace ProjectManagementToolkit.MPMM.MPMM_Document_Forms
{
    partial class RiskRegisterForm
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
            this.dgvRiskRegister = new System.Windows.Forms.DataGridView();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Date_Raised = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Raised_By = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Received_By = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Description_of_Issue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Impact = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Likelihood_Rating = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Impact_Rating = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Priority_Rating = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Preventative_Actions = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PreventativeActionsOwner = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PreventativeActions_Date = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Contingency_Actions = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Contingency_ActionsOwner = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Contingency_ActionsData = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.txtRiskRegisterProjectManager = new System.Windows.Forms.TextBox();
            this.label27 = new System.Windows.Forms.Label();
            this.txtRiskRegisterProjectName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnSaveProgress = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRiskRegister)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvRiskRegister
            // 
            this.dgvRiskRegister.AllowUserToOrderColumns = true;
            this.dgvRiskRegister.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvRiskRegister.BackgroundColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvRiskRegister.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.dgvRiskRegister.ColumnHeadersHeight = 29;
            this.dgvRiskRegister.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ID,
            this.Date_Raised,
            this.Raised_By,
            this.Received_By,
            this.Description_of_Issue,
            this.Impact,
            this.Likelihood_Rating,
            this.Impact_Rating,
            this.Priority_Rating,
            this.Preventative_Actions,
            this.PreventativeActionsOwner,
            this.PreventativeActions_Date,
            this.Contingency_Actions,
            this.Contingency_ActionsOwner,
            this.Contingency_ActionsData});
            this.dgvRiskRegister.EnableHeadersVisualStyles = false;
            this.dgvRiskRegister.Location = new System.Drawing.Point(15, 66);
            this.dgvRiskRegister.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dgvRiskRegister.Name = "dgvRiskRegister";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvRiskRegister.RowHeadersDefaultCellStyle = dataGridViewCellStyle5;
            this.dgvRiskRegister.RowHeadersWidth = 51;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle6.ForeColor = System.Drawing.Color.Black;
            this.dgvRiskRegister.RowsDefaultCellStyle = dataGridViewCellStyle6;
            this.dgvRiskRegister.Size = new System.Drawing.Size(1283, 505);
            this.dgvRiskRegister.TabIndex = 36;
            this.dgvRiskRegister.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridViewSolutionRaiseRaised_CellContentClick);
            // 
            // ID
            // 
            this.ID.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.ID.HeaderText = "ID";
            this.ID.MinimumWidth = 6;
            this.ID.Name = "ID";
            this.ID.Width = 55;
            // 
            // Date_Raised
            // 
            this.Date_Raised.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Date_Raised.HeaderText = "Date_Raised";
            this.Date_Raised.MinimumWidth = 6;
            this.Date_Raised.Name = "Date_Raised";
            this.Date_Raised.Width = 131;
            // 
            // Raised_By
            // 
            this.Raised_By.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Raised_By.HeaderText = "Raised_By";
            this.Raised_By.MinimumWidth = 6;
            this.Raised_By.Name = "Raised_By";
            this.Raised_By.Width = 117;
            // 
            // Received_By
            // 
            this.Received_By.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Received_By.HeaderText = "Received_By";
            this.Received_By.MinimumWidth = 6;
            this.Received_By.Name = "Received_By";
            this.Received_By.Width = 134;
            // 
            // Description_of_Issue
            // 
            this.Description_of_Issue.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Description_of_Issue.HeaderText = "Description_of_Risk";
            this.Description_of_Issue.MinimumWidth = 6;
            this.Description_of_Issue.Name = "Description_of_Issue";
            this.Description_of_Issue.Width = 184;
            // 
            // Impact
            // 
            this.Impact.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Impact.HeaderText = "Description_of_Impact";
            this.Impact.MinimumWidth = 6;
            this.Impact.Name = "Impact";
            this.Impact.Width = 202;
            // 
            // Likelihood_Rating
            // 
            this.Likelihood_Rating.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Likelihood_Rating.HeaderText = "Likelihood_Rating";
            this.Likelihood_Rating.MinimumWidth = 6;
            this.Likelihood_Rating.Name = "Likelihood_Rating";
            this.Likelihood_Rating.Width = 166;
            // 
            // Impact_Rating
            // 
            this.Impact_Rating.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Impact_Rating.HeaderText = "Impact_Rating";
            this.Impact_Rating.MinimumWidth = 6;
            this.Impact_Rating.Name = "Impact_Rating";
            this.Impact_Rating.Width = 142;
            // 
            // Priority_Rating
            // 
            this.Priority_Rating.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Priority_Rating.HeaderText = "Priority_Rating";
            this.Priority_Rating.MinimumWidth = 6;
            this.Priority_Rating.Name = "Priority_Rating";
            this.Priority_Rating.Width = 146;
            // 
            // Preventative_Actions
            // 
            this.Preventative_Actions.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Preventative_Actions.HeaderText = "Preventative_Actions";
            this.Preventative_Actions.MinimumWidth = 6;
            this.Preventative_Actions.Name = "Preventative_Actions";
            this.Preventative_Actions.Width = 189;
            // 
            // PreventativeActionsOwner
            // 
            this.PreventativeActionsOwner.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.PreventativeActionsOwner.HeaderText = "Preventative_ActionsOwner";
            this.PreventativeActionsOwner.MinimumWidth = 6;
            this.PreventativeActionsOwner.Name = "PreventativeActionsOwner";
            this.PreventativeActionsOwner.Width = 236;
            // 
            // PreventativeActions_Date
            // 
            this.PreventativeActions_Date.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.PreventativeActions_Date.HeaderText = "Preventative_ActionsDate";
            this.PreventativeActions_Date.MinimumWidth = 6;
            this.PreventativeActions_Date.Name = "PreventativeActions_Date";
            this.PreventativeActions_Date.Width = 223;
            // 
            // Contingency_Actions
            // 
            this.Contingency_Actions.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Contingency_Actions.HeaderText = "Contingency_Actions";
            this.Contingency_Actions.MinimumWidth = 6;
            this.Contingency_Actions.Name = "Contingency_Actions";
            this.Contingency_Actions.Width = 193;
            // 
            // Contingency_ActionsOwner
            // 
            this.Contingency_ActionsOwner.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Contingency_ActionsOwner.HeaderText = "Contingency_ActionsOwner";
            this.Contingency_ActionsOwner.MinimumWidth = 6;
            this.Contingency_ActionsOwner.Name = "Contingency_ActionsOwner";
            this.Contingency_ActionsOwner.Width = 240;
            // 
            // Contingency_ActionsData
            // 
            this.Contingency_ActionsData.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.ColumnHeader;
            this.Contingency_ActionsData.HeaderText = "Contingency_ActionsData";
            this.Contingency_ActionsData.MinimumWidth = 6;
            this.Contingency_ActionsData.Name = "Contingency_ActionsData";
            this.Contingency_ActionsData.Width = 227;
            // 
            // btnExport
            // 
            this.btnExport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnExport.Location = new System.Drawing.Point(898, 3);
            this.btnExport.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(137, 55);
            this.btnExport.TabIndex = 35;
            this.btnExport.Text = "Export to Excel";
            this.btnExport.UseVisualStyleBackColor = false;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSave.Location = new System.Drawing.Point(710, 6);
            this.btnSave.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(165, 52);
            this.btnSave.TabIndex = 34;
            this.btnSave.Text = "Complete";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // txtRiskRegisterProjectManager
            // 
            this.txtRiskRegisterProjectManager.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.txtRiskRegisterProjectManager.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRiskRegisterProjectManager.ForeColor = System.Drawing.Color.Black;
            this.txtRiskRegisterProjectManager.Location = new System.Drawing.Point(165, 36);
            this.txtRiskRegisterProjectManager.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.txtRiskRegisterProjectManager.Name = "txtRiskRegisterProjectManager";
            this.txtRiskRegisterProjectManager.Size = new System.Drawing.Size(323, 23);
            this.txtRiskRegisterProjectManager.TabIndex = 32;
            this.txtRiskRegisterProjectManager.Text = "Project Manager";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label27.ForeColor = System.Drawing.Color.Black;
            this.label27.Location = new System.Drawing.Point(27, 14);
            this.label27.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(105, 16);
            this.label27.TabIndex = 31;
            this.label27.Text = "Project Name:";
            // 
            // txtRiskRegisterProjectName
            // 
            this.txtRiskRegisterProjectName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.txtRiskRegisterProjectName.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtRiskRegisterProjectName.ForeColor = System.Drawing.Color.Black;
            this.txtRiskRegisterProjectName.Location = new System.Drawing.Point(167, 6);
            this.txtRiskRegisterProjectName.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.txtRiskRegisterProjectName.Name = "txtRiskRegisterProjectName";
            this.txtRiskRegisterProjectName.Size = new System.Drawing.Size(323, 23);
            this.txtRiskRegisterProjectName.TabIndex = 30;
            this.txtRiskRegisterProjectName.Text = "Project Name";
            this.txtRiskRegisterProjectName.TextChanged += new System.EventHandler(this.txtRiskRegisterProjectName_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(27, 37);
            this.label1.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(128, 16);
            this.label1.TabIndex = 33;
            this.label1.Text = "Project Manager:";
            // 
            // btnSaveProgress
            // 
            this.btnSaveProgress.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnSaveProgress.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSaveProgress.Location = new System.Drawing.Point(525, 6);
            this.btnSaveProgress.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnSaveProgress.Name = "btnSaveProgress";
            this.btnSaveProgress.Size = new System.Drawing.Size(165, 52);
            this.btnSaveProgress.TabIndex = 34;
            this.btnSaveProgress.Text = "Save Progress";
            this.btnSaveProgress.UseVisualStyleBackColor = false;
            this.btnSaveProgress.Click += new System.EventHandler(this.btnSaveProgress_Click);
            // 
            // RiskRegisterForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.ClientSize = new System.Drawing.Size(1313, 624);
            this.Controls.Add(this.dgvRiskRegister);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnSaveProgress);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.txtRiskRegisterProjectManager);
            this.Controls.Add(this.label27);
            this.Controls.Add(this.txtRiskRegisterProjectName);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "RiskRegisterForm";
            this.Text = "RiskRegisterForm";
            this.Load += new System.EventHandler(this.RiskRegisterForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvRiskRegister)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvRiskRegister;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.TextBox txtRiskRegisterProjectManager;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.TextBox txtRiskRegisterProjectName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridViewTextBoxColumn ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn Date_Raised;
        private System.Windows.Forms.DataGridViewTextBoxColumn Raised_By;
        private System.Windows.Forms.DataGridViewTextBoxColumn Received_By;
        private System.Windows.Forms.DataGridViewTextBoxColumn Description_of_Issue;
        private System.Windows.Forms.DataGridViewTextBoxColumn Impact;
        private System.Windows.Forms.DataGridViewTextBoxColumn Likelihood_Rating;
        private System.Windows.Forms.DataGridViewTextBoxColumn Impact_Rating;
        private System.Windows.Forms.DataGridViewTextBoxColumn Priority_Rating;
        private System.Windows.Forms.DataGridViewTextBoxColumn Preventative_Actions;
        private System.Windows.Forms.DataGridViewTextBoxColumn PreventativeActionsOwner;
        private System.Windows.Forms.DataGridViewTextBoxColumn PreventativeActions_Date;
        private System.Windows.Forms.DataGridViewTextBoxColumn Contingency_Actions;
        private System.Windows.Forms.DataGridViewTextBoxColumn Contingency_ActionsOwner;
        private System.Windows.Forms.DataGridViewTextBoxColumn Contingency_ActionsData;
        private System.Windows.Forms.Button btnSaveProgress;
    }
}