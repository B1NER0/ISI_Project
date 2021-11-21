
namespace ProjectManagementToolkit.MPMM.MPMM_Forms.Project_Management
{
    partial class frmContractConclusion
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
            this.btnContractConcWConcludeStage = new System.Windows.Forms.Button();
            this.btnContractConcWUpdateProjectImplementationPlan = new System.Windows.Forms.Button();
            this.btnContractConcWCompileBalanceOfEnquiries = new System.Windows.Forms.Button();
            this.ContractConcDGV = new System.Windows.Forms.DataGridView();
            this.btnBackToContractConclusion = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.ContractConcDGV)).BeginInit();
            this.SuspendLayout();
            // 
            // btnContractConcWConcludeStage
            // 
            this.btnContractConcWConcludeStage.Location = new System.Drawing.Point(475, 51);
            this.btnContractConcWConcludeStage.Name = "btnContractConcWConcludeStage";
            this.btnContractConcWConcludeStage.Size = new System.Drawing.Size(210, 35);
            this.btnContractConcWConcludeStage.TabIndex = 5;
            this.btnContractConcWConcludeStage.Text = "Conclude Stage";
            this.btnContractConcWConcludeStage.UseVisualStyleBackColor = true;
            this.btnContractConcWConcludeStage.Click += new System.EventHandler(this.btnContractConcWConcludeStage_Click);
            // 
            // btnContractConcWUpdateProjectImplementationPlan
            // 
            this.btnContractConcWUpdateProjectImplementationPlan.Location = new System.Drawing.Point(246, 51);
            this.btnContractConcWUpdateProjectImplementationPlan.Name = "btnContractConcWUpdateProjectImplementationPlan";
            this.btnContractConcWUpdateProjectImplementationPlan.Size = new System.Drawing.Size(210, 35);
            this.btnContractConcWUpdateProjectImplementationPlan.TabIndex = 4;
            this.btnContractConcWUpdateProjectImplementationPlan.Text = "Update Project Implementation Plan";
            this.btnContractConcWUpdateProjectImplementationPlan.UseVisualStyleBackColor = true;
            this.btnContractConcWUpdateProjectImplementationPlan.Click += new System.EventHandler(this.btnContractConcWUpdateProjectImplementationPlan_Click);
            // 
            // btnContractConcWCompileBalanceOfEnquiries
            // 
            this.btnContractConcWCompileBalanceOfEnquiries.Location = new System.Drawing.Point(12, 50);
            this.btnContractConcWCompileBalanceOfEnquiries.Name = "btnContractConcWCompileBalanceOfEnquiries";
            this.btnContractConcWCompileBalanceOfEnquiries.Size = new System.Drawing.Size(210, 36);
            this.btnContractConcWCompileBalanceOfEnquiries.TabIndex = 3;
            this.btnContractConcWCompileBalanceOfEnquiries.Text = "Compile Balance of Enquiries";
            this.btnContractConcWCompileBalanceOfEnquiries.UseVisualStyleBackColor = true;
            this.btnContractConcWCompileBalanceOfEnquiries.Click += new System.EventHandler(this.btnContractConcWCompileBalanceOfEnquiries_Click);
            // 
            // ContractConcDGV
            // 
            this.ContractConcDGV.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.ContractConcDGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ContractConcDGV.Location = new System.Drawing.Point(13, 102);
            this.ContractConcDGV.Name = "ContractConcDGV";
            this.ContractConcDGV.Size = new System.Drawing.Size(672, 332);
            this.ContractConcDGV.TabIndex = 6;
            // 
            // btnBackToContractConclusion
            // 
            this.btnBackToContractConclusion.Location = new System.Drawing.Point(13, 474);
            this.btnBackToContractConclusion.Name = "btnBackToContractConclusion";
            this.btnBackToContractConclusion.Size = new System.Drawing.Size(189, 23);
            this.btnBackToContractConclusion.TabIndex = 7;
            this.btnBackToContractConclusion.Text = "Back to Contract Conclusion";
            this.btnBackToContractConclusion.UseVisualStyleBackColor = true;
            this.btnBackToContractConclusion.Click += new System.EventHandler(this.btnBackToContractConclusion_Click);
            // 
            // frmContractConclusion
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1068, 595);
            this.Controls.Add(this.btnBackToContractConclusion);
            this.Controls.Add(this.ContractConcDGV);
            this.Controls.Add(this.btnContractConcWConcludeStage);
            this.Controls.Add(this.btnContractConcWUpdateProjectImplementationPlan);
            this.Controls.Add(this.btnContractConcWCompileBalanceOfEnquiries);
            this.Name = "frmContractConclusion";
            this.Text = "frmContractConclusion";
            ((System.ComponentModel.ISupportInitialize)(this.ContractConcDGV)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnContractConcWConcludeStage;
        private System.Windows.Forms.Button btnContractConcWUpdateProjectImplementationPlan;
        private System.Windows.Forms.Button btnContractConcWCompileBalanceOfEnquiries;
        public System.Windows.Forms.DataGridView ContractConcDGV;
        private System.Windows.Forms.Button btnBackToContractConclusion;
    }
}