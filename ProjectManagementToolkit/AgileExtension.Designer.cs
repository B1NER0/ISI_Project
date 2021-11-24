
namespace ProjectManagementToolkit
{
    partial class AgileExtension
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
            this.lblIndicator = new System.Windows.Forms.Label();
            this.btnRun = new System.Windows.Forms.Button();
            this.btnSetPath = new System.Windows.Forms.Button();
            this.openAgile = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // lblIndicator
            // 
            this.lblIndicator.AutoSize = true;
            this.lblIndicator.Location = new System.Drawing.Point(19, 74);
            this.lblIndicator.Name = "lblIndicator";
            this.lblIndicator.Size = new System.Drawing.Size(0, 13);
            this.lblIndicator.TabIndex = 17;
            // 
            // btnRun
            // 
            this.btnRun.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnRun.Enabled = false;
            this.btnRun.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRun.ForeColor = System.Drawing.Color.Black;
            this.btnRun.Location = new System.Drawing.Point(181, 30);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(144, 28);
            this.btnRun.TabIndex = 16;
            this.btnRun.Text = "Launch Agile Extension";
            this.btnRun.UseVisualStyleBackColor = false;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // btnSetPath
            // 
            this.btnSetPath.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(73)))), ((int)(((byte)(173)))), ((int)(((byte)(252)))));
            this.btnSetPath.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSetPath.ForeColor = System.Drawing.Color.Black;
            this.btnSetPath.Location = new System.Drawing.Point(22, 30);
            this.btnSetPath.Name = "btnSetPath";
            this.btnSetPath.Size = new System.Drawing.Size(144, 28);
            this.btnSetPath.TabIndex = 15;
            this.btnSetPath.Text = "Set Path";
            this.btnSetPath.UseVisualStyleBackColor = false;
            this.btnSetPath.Click += new System.EventHandler(this.btnSetPath_Click);
            // 
            // openAgile
            // 
            this.openAgile.FileName = "openFileDialog1";
            // 
            // AgileExtension
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(209)))), ((int)(((byte)(237)))), ((int)(((byte)(242)))));
            this.ClientSize = new System.Drawing.Size(347, 95);
            this.Controls.Add(this.lblIndicator);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.btnSetPath);
            this.Name = "AgileExtension";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "AgileExtension";
            this.Load += new System.EventHandler(this.AgileExtension_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label lblIndicator;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Button btnSetPath;
        private System.Windows.Forms.OpenFileDialog openAgile;
    }
}