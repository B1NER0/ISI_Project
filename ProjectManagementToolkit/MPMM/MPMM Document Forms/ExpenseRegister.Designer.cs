namespace ProjectManagementToolkit.MPMM.MPMM_Document_Forms
{
    partial class ExpenseRegister
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
            this.dataGridViewExpenseRegister = new System.Windows.Forms.DataGridView();
            this.ActivityID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ActivityDescription = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TaskIID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TaskDesc = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ExpenseID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ExpenseType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ExpenseDescription = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ExpenseAmount = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ApprovalStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ApprovalDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Approver = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PaymentStatus = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PaymentDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Payee = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Method = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnSave = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtIssueRegisterProjectManager = new System.Windows.Forms.TextBox();
            this.txtIssueRegisterProjectName = new System.Windows.Forms.TextBox();
            this.label27 = new System.Windows.Forms.Label();
            this.btnExport = new System.Windows.Forms.Button();
            this.btnSaveProgress = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExpenseRegister)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewExpenseRegister
            // 
            this.dataGridViewExpenseRegister.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            dataGridViewCellStyle4.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle4.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle4.SelectionForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewExpenseRegister.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle4;
            this.dataGridViewExpenseRegister.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewExpenseRegister.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ActivityID,
            this.ActivityDescription,
            this.TaskIID,
            this.TaskDesc,
            this.ExpenseID,
            this.ExpenseType,
            this.ExpenseDescription,
            this.ExpenseAmount,
            this.ApprovalStatus,
            this.ApprovalDate,
            this.Approver,
            this.PaymentStatus,
            this.PaymentDate,
            this.Payee,
            this.Method});
            this.dataGridViewExpenseRegister.EnableHeadersVisualStyles = false;
            this.dataGridViewExpenseRegister.Location = new System.Drawing.Point(12, 78);
            this.dataGridViewExpenseRegister.Name = "dataGridViewExpenseRegister";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle5.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            dataGridViewCellStyle5.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle5.ForeColor = System.Drawing.Color.Black;
            dataGridViewCellStyle5.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle5.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle5.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewExpenseRegister.RowHeadersDefaultCellStyle = dataGridViewCellStyle5;
            this.dataGridViewExpenseRegister.RowHeadersWidth = 51;
            dataGridViewCellStyle6.Font = new System.Drawing.Font("Arial", 10.8F);
            this.dataGridViewExpenseRegister.RowsDefaultCellStyle = dataGridViewCellStyle6;
            this.dataGridViewExpenseRegister.Size = new System.Drawing.Size(904, 370);
            this.dataGridViewExpenseRegister.TabIndex = 4;
            // 
            // ActivityID
            // 
            this.ActivityID.HeaderText = "Activity ID";
            this.ActivityID.MinimumWidth = 6;
            this.ActivityID.Name = "ActivityID";
            this.ActivityID.Width = 125;
            // 
            // ActivityDescription
            // 
            this.ActivityDescription.HeaderText = "Activity Description";
            this.ActivityDescription.MinimumWidth = 6;
            this.ActivityDescription.Name = "ActivityDescription";
            this.ActivityDescription.Width = 125;
            // 
            // TaskIID
            // 
            this.TaskIID.HeaderText = "Task ID";
            this.TaskIID.MinimumWidth = 6;
            this.TaskIID.Name = "TaskIID";
            this.TaskIID.Width = 125;
            // 
            // TaskDesc
            // 
            this.TaskDesc.HeaderText = "Task Description ";
            this.TaskDesc.MinimumWidth = 6;
            this.TaskDesc.Name = "TaskDesc";
            this.TaskDesc.Width = 125;
            // 
            // ExpenseID
            // 
            this.ExpenseID.HeaderText = "Expense ID";
            this.ExpenseID.MinimumWidth = 6;
            this.ExpenseID.Name = "ExpenseID";
            this.ExpenseID.Width = 125;
            // 
            // ExpenseType
            // 
            this.ExpenseType.HeaderText = "Expense Type";
            this.ExpenseType.MinimumWidth = 6;
            this.ExpenseType.Name = "ExpenseType";
            this.ExpenseType.Width = 125;
            // 
            // ExpenseDescription
            // 
            this.ExpenseDescription.HeaderText = "Expense Description";
            this.ExpenseDescription.MinimumWidth = 6;
            this.ExpenseDescription.Name = "ExpenseDescription";
            this.ExpenseDescription.Width = 125;
            // 
            // ExpenseAmount
            // 
            this.ExpenseAmount.HeaderText = "Expense Amount";
            this.ExpenseAmount.MinimumWidth = 6;
            this.ExpenseAmount.Name = "ExpenseAmount";
            this.ExpenseAmount.Width = 125;
            // 
            // ApprovalStatus
            // 
            this.ApprovalStatus.HeaderText = "Approval Status";
            this.ApprovalStatus.MinimumWidth = 6;
            this.ApprovalStatus.Name = "ApprovalStatus";
            this.ApprovalStatus.Width = 125;
            // 
            // ApprovalDate
            // 
            this.ApprovalDate.HeaderText = "Approval Date";
            this.ApprovalDate.MinimumWidth = 6;
            this.ApprovalDate.Name = "ApprovalDate";
            this.ApprovalDate.Width = 125;
            // 
            // Approver
            // 
            this.Approver.HeaderText = "Approver";
            this.Approver.MinimumWidth = 6;
            this.Approver.Name = "Approver";
            this.Approver.Width = 125;
            // 
            // PaymentStatus
            // 
            this.PaymentStatus.HeaderText = "Payment Status";
            this.PaymentStatus.MinimumWidth = 6;
            this.PaymentStatus.Name = "PaymentStatus";
            this.PaymentStatus.Width = 125;
            // 
            // PaymentDate
            // 
            this.PaymentDate.HeaderText = "Payment Date";
            this.PaymentDate.MinimumWidth = 6;
            this.PaymentDate.Name = "PaymentDate";
            this.PaymentDate.Width = 125;
            // 
            // Payee
            // 
            this.Payee.HeaderText = "Payee";
            this.Payee.MinimumWidth = 6;
            this.Payee.Name = "Payee";
            this.Payee.Width = 125;
            // 
            // Method
            // 
            this.Method.HeaderText = "Method";
            this.Method.MinimumWidth = 6;
            this.Method.Name = "Method";
            this.Method.Width = 125;
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSave.Location = new System.Drawing.Point(483, 24);
            this.btnSave.Margin = new System.Windows.Forms.Padding(2);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(120, 37);
            this.btnSave.TabIndex = 37;
            this.btnSave.Text = "Complete";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(11, 41);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(128, 16);
            this.label1.TabIndex = 41;
            this.label1.Text = "Project Manager:";
            // 
            // txtIssueRegisterProjectManager
            // 
            this.txtIssueRegisterProjectManager.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.txtIssueRegisterProjectManager.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtIssueRegisterProjectManager.ForeColor = System.Drawing.Color.Black;
            this.txtIssueRegisterProjectManager.Location = new System.Drawing.Point(116, 41);
            this.txtIssueRegisterProjectManager.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.txtIssueRegisterProjectManager.Name = "txtIssueRegisterProjectManager";
            this.txtIssueRegisterProjectManager.Size = new System.Drawing.Size(274, 23);
            this.txtIssueRegisterProjectManager.TabIndex = 40;
            this.txtIssueRegisterProjectManager.Text = "Project Manager";
            // 
            // txtIssueRegisterProjectName
            // 
            this.txtIssueRegisterProjectName.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.txtIssueRegisterProjectName.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtIssueRegisterProjectName.ForeColor = System.Drawing.Color.Black;
            this.txtIssueRegisterProjectName.Location = new System.Drawing.Point(116, 17);
            this.txtIssueRegisterProjectName.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.txtIssueRegisterProjectName.Name = "txtIssueRegisterProjectName";
            this.txtIssueRegisterProjectName.Size = new System.Drawing.Size(274, 23);
            this.txtIssueRegisterProjectName.TabIndex = 38;
            this.txtIssueRegisterProjectName.Text = "Project Name";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label27.ForeColor = System.Drawing.Color.Black;
            this.label27.Location = new System.Drawing.Point(11, 22);
            this.label27.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(105, 16);
            this.label27.TabIndex = 39;
            this.label27.Text = "Project Name:";
            // 
            // btnExport
            // 
            this.btnExport.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnExport.Location = new System.Drawing.Point(796, 24);
            this.btnExport.Margin = new System.Windows.Forms.Padding(2);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(120, 37);
            this.btnExport.TabIndex = 43;
            this.btnExport.Text = "Export to Excel";
            this.btnExport.UseVisualStyleBackColor = false;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // btnSaveProgress
            // 
            this.btnSaveProgress.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnSaveProgress.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnSaveProgress.Location = new System.Drawing.Point(636, 24);
            this.btnSaveProgress.Margin = new System.Windows.Forms.Padding(2);
            this.btnSaveProgress.Name = "btnSaveProgress";
            this.btnSaveProgress.Size = new System.Drawing.Size(133, 37);
            this.btnSaveProgress.TabIndex = 44;
            this.btnSaveProgress.Text = "Save Progress";
            this.btnSaveProgress.UseVisualStyleBackColor = false;
            this.btnSaveProgress.Click += new System.EventHandler(this.btnSaveProgress_Click);
            // 
            // ExpenseRegister
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.ClientSize = new System.Drawing.Size(933, 484);
            this.Controls.Add(this.btnSaveProgress);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtIssueRegisterProjectManager);
            this.Controls.Add(this.txtIssueRegisterProjectName);
            this.Controls.Add(this.label27);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.dataGridViewExpenseRegister);
            this.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.Black;
            this.Name = "ExpenseRegister";
            this.Text = "ExpenseRegister";
            this.Load += new System.EventHandler(this.ExpenseRegister_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewExpenseRegister)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.DataGridView dataGridViewExpenseRegister;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtIssueRegisterProjectManager;
        private System.Windows.Forms.TextBox txtIssueRegisterProjectName;
        private System.Windows.Forms.Label label27;
        private System.Windows.Forms.DataGridViewTextBoxColumn ActivityID;
        private System.Windows.Forms.DataGridViewTextBoxColumn ActivityDescription;
        private System.Windows.Forms.DataGridViewTextBoxColumn TaskIID;
        private System.Windows.Forms.DataGridViewTextBoxColumn TaskDesc;
        private System.Windows.Forms.DataGridViewTextBoxColumn ExpenseID;
        private System.Windows.Forms.DataGridViewTextBoxColumn ExpenseType;
        private System.Windows.Forms.DataGridViewTextBoxColumn ExpenseDescription;
        private System.Windows.Forms.DataGridViewTextBoxColumn ExpenseAmount;
        private System.Windows.Forms.DataGridViewTextBoxColumn ApprovalStatus;
        private System.Windows.Forms.DataGridViewTextBoxColumn ApprovalDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn Approver;
        private System.Windows.Forms.DataGridViewTextBoxColumn PaymentStatus;
        private System.Windows.Forms.DataGridViewTextBoxColumn PaymentDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn Payee;
        private System.Windows.Forms.DataGridViewTextBoxColumn Method;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnSaveProgress;
    }
}