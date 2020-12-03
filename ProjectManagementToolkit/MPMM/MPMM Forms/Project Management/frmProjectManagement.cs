﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProjectManagementToolkit
{
    public partial class frmProjectManagement : Form
    {
        public frmProjectManagement(string type)
        {
            InitializeComponent();
            lblType.Text = type;
        }

        private void frmProjectManagement_Activated(object sender, EventArgs e)
        {
            pnlHideTabControl.Location = new Point(6, 106);
            tbp_Project_Management.SelectTab(lblType.Text);
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime datetime = DateTime.Now;
            this.lblReportCenterDate.Text = datetime.ToString("dddd, MMMM dd, yyyy");
            this.lblReportCenterTime.Text = datetime.ToString("hh:mm tt");
        }

        private void btnDirectingAProject_Click(object sender, EventArgs e)
        {
            this.Hide();
            PLSM_ProcessGroupingGovernance_Interfaces plsm = new PLSM_ProcessGroupingGovernance_Interfaces("tbpDirectingAProject");
            plsm.Show();
        }

        private void btnPlanningAProject_Click(object sender, EventArgs e)
        {
            this.Hide();
            PLSM_ProcessGroupingGovernance_Interfaces plsm = new PLSM_ProcessGroupingGovernance_Interfaces("tbpPlanningAProject");
            plsm.Show();
        }

        private void btnStartingUpAProject_Click(object sender, EventArgs e)
        {
            this.Hide();
            PLSM_ProcessGroupingGovernance_Interfaces plsm = new PLSM_ProcessGroupingGovernance_Interfaces("tbpStartingUpAProject");
            plsm.Show();
        }

        private void btnInitiatingAProject_Click(object sender, EventArgs e)
        {
            this.Hide();
            PLSM_ProcessGroupingGovernance_Interfaces plsm = new PLSM_ProcessGroupingGovernance_Interfaces("tbpInitiatingAProject");
            plsm.Show();
        }

        private void btnManagingProductDelivery_Click(object sender, EventArgs e)
        {
            this.Hide();
            PLSM_ProcessGroupingGovernance_Interfaces plsm = new PLSM_ProcessGroupingGovernance_Interfaces("tbpManagingProductDelivery");
            plsm.Show();
        }

        private void btnManagingStageBoundaries_Click(object sender, EventArgs e)
        {
            this.Hide();
            PLSM_ProcessGroupingGovernance_Interfaces plsm = new PLSM_ProcessGroupingGovernance_Interfaces("tbpManagingAStageBoundary");
            plsm.Show();
        }

        private void btnControllingAStage_Click(object sender, EventArgs e)
        {
            this.Hide();
            PLSM_ProcessGroupingGovernance_Interfaces plsm = new PLSM_ProcessGroupingGovernance_Interfaces("tbpControllingAStage");
            plsm.Show();
        }

        private void btnClosingAProject_Click(object sender, EventArgs e)
        {
            this.Hide();
            PLSM_ProcessGroupingGovernance_Interfaces plsm = new PLSM_ProcessGroupingGovernance_Interfaces("tbpClosingAProject");
            plsm.Show();
        }

        private void linkOpenProjects_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lblSelectingReport.Text = linkOpenProjects.Text;
        }

        private void linkProjectsBalanceSheet_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lblSelectingReport.Text = linkProjectsBalanceSheet.Text;
        }

        private void btnGovernanceProcess_Click(object sender, EventArgs e)
        {
            tbcProcesses.SelectTab(0);
        }

        private void btnRiskManagement_Click(object sender, EventArgs e)
        {
            tbcProcesses.SelectTab(9);
        }

        private void btnChnageManagement_Click(object sender, EventArgs e)
        {
            tbcProcesses.SelectTab(2);
        }

        private void btnCostAndExpenseManagement_Click(object sender, EventArgs e)
        {
            tbcProcesses.SelectTab(1);
        }

        private void btnCommunicationAndStakeholderManagement_Click(object sender, EventArgs e)
        {
            tbcProcesses.SelectTab(3);
        }

        private void btnIssueManagement_Click(object sender, EventArgs e)
        {
            tbcProcesses.SelectTab(4);
        }

        private void btnQualityManagement_Click(object sender, EventArgs e)
        {
            tbcProcesses.SelectTab(10);
        }

        private void btnAcceptanceManagement_Click(object sender, EventArgs e)
        {
            tbcProcesses.SelectTab(5);
        }

        private void btnProcurementManagement_Click(object sender, EventArgs e)
        {
            tbcProcesses.SelectTab(6);
        }

        private void btnTimeAllocation_Click(object sender, EventArgs e)
        {
            tbcProcesses.SelectTab(7);
        }

        private void btnExceptionManagement_Click(object sender, EventArgs e)
        {
            tbcProcesses.SelectTab(8);
        }

        private void btnBackToPLCMFrontEND_Click(object sender, EventArgs e)
        {
            tbp_Project_Management.SelectTab("tbp_PLSMs");
        }

        private void btnRisk_Click(object sender, EventArgs e)
        {
            tbp_Project_Management.SelectTab("tbp_Processes");
            btnRiskManagement_Click(sender, e);
        }

        private void btnAcceptance_Click(object sender, EventArgs e)
        {
            tbp_Project_Management.SelectTab("tbp_Processes");
            btnAcceptanceManagement_Click(sender, e);
        }

        private void btnTimeAllocation2_Click(object sender, EventArgs e)
        {
            tbp_Project_Management.SelectTab("tbp_Processes");
            btnTimeAllocation_Click(sender, e);
        }

        private void btnChanges_Click(object sender, EventArgs e)
        {
            tbp_Project_Management.SelectTab("tbp_Processes");
            btnChnageManagement_Click(sender, e);
        }

        private void btnProcurement_Click(object sender, EventArgs e)
        {
            tbp_Project_Management.SelectTab("tbp_Processes");
            btnProcurementManagement_Click(sender, e);
        }

        private void btnIssues_Click(object sender, EventArgs e)
        {
            tbp_Project_Management.SelectTab("tbp_Processes");
            btnIssueManagement_Click(sender, e);
        }

        private void btnCost_Click(object sender, EventArgs e)
        {
            tbp_Project_Management.SelectTab("tbp_Processes");
            btnCostAndExpenseManagement_Click(sender, e);
        }

        private void btnComms_Click(object sender, EventArgs e)
        {
            tbp_Project_Management.SelectTab("tbp_Processes");
            btnCommunicationAndStakeholderManagement_Click(sender, e);
        }

        private void btnQuality_Click(object sender, EventArgs e)
        {
            tbp_Project_Management.SelectTab("tbp_Processes");
            btnQualityManagement_Click(sender, e);
        }

        private void btnStakeholders_Click(object sender, EventArgs e)
        {
            tbp_Project_Management.SelectTab("tbp_Processes");
            btnCommunicationAndStakeholderManagement_Click(sender, e);
        }
    }
}
