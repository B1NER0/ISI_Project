﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ProjectManagementToolkit.MPMM.MPMM_Document_Forms;
using ProjectManagementToolkit.MPMM.MPMM_Forms.Project_Management;

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

        private void btnRiskManagementTemplate_Click(object sender, EventArgs e)
        {
            this.Hide();
            RiskManagementProcessDocumentForm RiskForm = new RiskManagementProcessDocumentForm();
            RiskForm.Show();
        }

        private void btnRiskManagementRegister_Click(object sender, EventArgs e)
        {
            this.Hide();
            RiskRegisterForm RiskRegister = new RiskRegisterForm();
            RiskRegister.Show();
        }

        private void btnChangeManagementTemplate_Click(object sender, EventArgs e)
        {
            this.Hide();
            ChangeManagementProcessDocumentForm ChangeForm = new ChangeManagementProcessDocumentForm();
            ChangeForm.Show();
        }

        private void btnChangeManagementRegister_Click(object sender, EventArgs e)
        {
            this.Hide();
            ChangeRegister ChangeRegister = new ChangeRegister();
            ChangeRegister.Show();
        }

        private void btnCostAndExpenseManagementManagementTemplate_Click(object sender, EventArgs e)
        {
            this.Hide();
            CostManagementProcessDocumentForm CostForm = new CostManagementProcessDocumentForm();
            CostForm.Show();
        }

        private void btnCostAndExpenseManagementRegister_Click(object sender, EventArgs e)
        {
            this.Hide();
            ExpenseRegister ExpenseRegister = new ExpenseRegister();
            ExpenseRegister.Show();
        }

        private void btnCaSManagementTemplate_Click(object sender, EventArgs e)
        {
            this.Hide();
            CommunicationsManagementProcessDocumentForm CommunicationForm = new CommunicationsManagementProcessDocumentForm();
            CommunicationForm.Show();
        }

        private void btnCaSManagementRegister_Click(object sender, EventArgs e)
        {
            this.Hide();
            CommunicationsRegister ChangeRegister = new CommunicationsRegister();
            ChangeRegister.Show();
        }

        private void btnIssueManagementTemplate_Click(object sender, EventArgs e)
        {
            this.Hide();
            IssueManagementProcessDocumentForm IssueForm = new IssueManagementProcessDocumentForm(); 
            IssueForm.Show();
        }

        private void btnIssueManagementRegister_Click(object sender, EventArgs e)
        {
            this.Hide();
            IssueRegisterForm IssueRegister = new IssueRegisterForm();
            IssueRegister.Show();
        }

        private void btnAcceptanceManagementRegister_Click(object sender, EventArgs e)
        {
            this.Hide();
            AcceptanceRegister AcceptanceForm = new AcceptanceRegister();
            AcceptanceForm.Show();
        }

        private void btnAcceptanceManagementTemplate_Click(object sender, EventArgs e)
        {
            this.Hide();
            AcceptanceManagementProcessDocumentForm AcceptanceForm = new AcceptanceManagementProcessDocumentForm();
            AcceptanceForm.Show();
        }

        private void btnProcurementManagementTemplate_Click(object sender, EventArgs e)
        {
            this.Hide();
            ProcurementManagementProcess ProcureForm = new ProcurementManagementProcess();
            ProcureForm.Show();
        }

        private void btnProcurementManagementRegister_Click(object sender, EventArgs e)
        {
            this.Hide();
            ProcurementRegister ProcureRegister = new ProcurementRegister();
            ProcureRegister.Show();
        }

        private void btnTimeManagementTemplate_Click(object sender, EventArgs e)
        {
            this.Hide();
            TimeMangementProcessDocumentForm TimeForm = new TimeMangementProcessDocumentForm();
            TimeForm.Show();
        }

        private void btnTimeManagementRegister_Click(object sender, EventArgs e)
        {
            this.Hide();
            TimesheetRegister TimeRegister = new TimesheetRegister();
            TimeRegister.Show();
        }

        private void btnQualityManagementTemplate_Click(object sender, EventArgs e)
        {
            this.Hide();
            QualityManagementProcessDocumentForm QualityForm = new QualityManagementProcessDocumentForm();
            QualityForm.Show();
        }

        private void btnQualityManagementRegister_Click(object sender, EventArgs e)
        {
            this.Hide();
            QualityRegister QualRegister = new QualityRegister();
            QualRegister.Show();
        }

        private void btnTemplates_Click(object sender, EventArgs e)
        {
            MPMM.MPMM_Forms.Project_Management.frmTemplates Templates = new MPMM.MPMM_Forms.Project_Management.frmTemplates();
            Templates.Show();
        }

        private void btnProcessFlowOverview_Click(object sender, EventArgs e)
        {
            MPMM.MPMM_Forms.Project_Management.frmProcessFlowOverview ProcessFlowOverview = new MPMM.MPMM_Forms.Project_Management.frmProcessFlowOverview();
            ProcessFlowOverview.Show();
        }

        private void btnRoleDescription_Click(object sender, EventArgs e)
        {
            MPMM.MPMM_Forms.Project_Management.frmRoleDescriptions RoleDescriptions = new MPMM.MPMM_Forms.Project_Management.frmRoleDescriptions();
            RoleDescriptions.Show();
        }

        private void btnToolsAndTechniques_Click(object sender, EventArgs e)
        {
            MPMM.MPMM_Forms.Project_Management.frmToolsAndTechniques ToolsResource = new MPMM.MPMM_Forms.Project_Management.frmToolsAndTechniques();
            ToolsResource.Show();
        }

        private void btnKnowledgeAndResources_Click(object sender, EventArgs e)
        {
            MPMM.MPMM_Forms.Project_Management.frmKnowledgeAndResources KnowResource = new MPMM.MPMM_Forms.Project_Management.frmKnowledgeAndResources();
            KnowResource.Show();
        }
               

        private void btnCRA_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpCRA");
            Warn.Show();
        }

        private void btnDRA_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpDRA");
            Warn.Show();
        }

        private void btnBackableProject_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpBackable");
            Warn.Show();
        }

        private void btnERA_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpERA");
            Warn.Show();
        }

        private void btnAuthorizedProject_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpAuthorized");
            Warn.Show();
        }

        private void btnFinancialClosure_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpFinacial");
            Warn.Show();
        }

        private void btnExecutionAgreement_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpExecution");
            Warn.Show();
        }

        private void btnHandOverDocument1_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpHandOver");
            Warn.Show();
        }

        private void btnClosureContracts_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpClosure");
            Warn.Show();
        }

        private void btnFRA_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpFRA");
            Warn.Show();
        }

        private void btnHandOverDocument2_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpHandOver2");
            Warn.Show();
        }

        private void btnStageGateRequirements_Click(object sender, EventArgs e)
        {
            frmWarnings Warn = new frmWarnings();
            Warn.tbcWarning.SelectTab("tbpCRA");
            Warn.Show();
        }

        private void btnWorkPakagesOpportunityScreening_Click(object sender, EventArgs e)
        {
            frmPLSM_Workpackages frmWorkpackages = new frmPLSM_Workpackages();
            frmWorkpackages.Show();
            frmWorkpackages.WorkPackagesTabControl.SelectedIndex = 1;

        }

        private void btnWorkPakagesPrefeasibility_Click(object sender, EventArgs e)
        {
            frmPLSM_Workpackages frmWorkpackages = new frmPLSM_Workpackages();
            frmWorkpackages.Show();
            frmWorkpackages.WorkPackagesTabControl.SelectedIndex = 2;

        }

        private void btnWorkPakagesFeasibility_Click(object sender, EventArgs e)
        {
            frmPLSM_Workpackages frmWorkpackages = new frmPLSM_Workpackages();
            frmWorkpackages.Show();
            frmWorkpackages.WorkPackagesTabControl.SelectedIndex = 3;

        }

        private void btnWorkPakagesBusinessPlan_Click(object sender, EventArgs e)
        {
            frmPLSM_Workpackages frmWorkpackages = new frmPLSM_Workpackages();
            frmWorkpackages.Show();
            frmWorkpackages.WorkPackagesTabControl.SelectedIndex = 4;
        }

        private void btnWorkPakagesProjectExecutionPlanning_Click(object sender, EventArgs e)
        {
            frmPLSM_Workpackages frmWorkpackages = new frmPLSM_Workpackages();
            frmWorkpackages.Show();
            frmWorkpackages.WorkPackagesTabControl.SelectedIndex = 5;
        }

        private void btnWorkPakagesContractConcluding_Click(object sender, EventArgs e)
        {
            frmPLSM_Workpackages frmWorkpackages = new frmPLSM_Workpackages();
            frmWorkpackages.Show();
            frmWorkpackages.WorkPackagesTabControl.SelectedIndex = 6;
        }

        private void btnWorkPakagesDetailedDesign_Click(object sender, EventArgs e)
        {
            frmPLSM_Workpackages frmWorkpackages = new frmPLSM_Workpackages();
            frmWorkpackages.Show();
            frmWorkpackages.WorkPackagesTabControl.SelectedIndex = 7;
        }

        private void btnWorkPakagesImplementation_Click(object sender, EventArgs e)
        {
            frmPLSM_Workpackages frmWorkpackages = new frmPLSM_Workpackages();
            frmWorkpackages.Show();
            frmWorkpackages.WorkPackagesTabControl.SelectedIndex = 9;
        }

        private void btnWorkPakagesTransfer_Click(object sender, EventArgs e)
        {
            frmPLSM_Workpackages frmWorkpackages = new frmPLSM_Workpackages();
            frmWorkpackages.Show();
            frmWorkpackages.WorkPackagesTabControl.SelectedIndex = 10;
        }

        private void btnWorkPakagesCloseOut_Click(object sender, EventArgs e)
        {
            frmPLSM_Workpackages frmWorkpackages = new frmPLSM_Workpackages();
            frmWorkpackages.Show();
            frmWorkpackages.WorkPackagesTabControl.SelectedIndex = 11;
        }

        private void btnWorkPakagesEvaluate_Click(object sender, EventArgs e)
        {
            frmPLSM_Workpackages frmWorkpackages = new frmPLSM_Workpackages();
            frmWorkpackages.Show();
            frmWorkpackages.WorkPackagesTabControl.SelectedIndex = 12;
        }

        private void btnWorkStreamsOpportunityScreening_Click(object sender, EventArgs e)
        {
            frmPLSM_Workstreams frmWorkStreams = new frmPLSM_Workstreams();
            frmWorkStreams.Show();
            frmWorkStreams.PLSMWorkstreamsTabControl.SelectedIndex = 1;
        }

        private void btnWorkStreamsPrefeasibility_Click(object sender, EventArgs e)
        {
            frmPLSM_Workstreams frmWorkStreams = new frmPLSM_Workstreams();
            frmWorkStreams.Show();
            frmWorkStreams.PLSMWorkstreamsTabControl.SelectedIndex = 1;
        }

        private void btnWorkStreamsImplementation_Click(object sender, EventArgs e)
        {
            frmPLSM_Workstreams frmWorkStreams = new frmPLSM_Workstreams();
            frmWorkStreams.Show();
            frmWorkStreams.PLSMWorkstreamsTabControl.SelectedIndex = 1;
        }

        private void btnWorkStreamsFeasibility_Click(object sender, EventArgs e)
        {
            frmPLSM_Workstreams frmWorkStreams = new frmPLSM_Workstreams();
            frmWorkStreams.Show();
            frmWorkStreams.PLSMWorkstreamsTabControl.SelectedIndex = 1;
        }

        private void btnWorkStreamsBusinessPlan_Click(object sender, EventArgs e)
        {
            frmPLSM_Workstreams frmWorkStreams = new frmPLSM_Workstreams();
            frmWorkStreams.Show();
            frmWorkStreams.PLSMWorkstreamsTabControl.SelectedIndex = 1;
        }

        private void btnWorkStreamsProjectExecutionPlanning_Click(object sender, EventArgs e)
        {
            frmPLSM_Workstreams frmWorkStreams = new frmPLSM_Workstreams();
            frmWorkStreams.Show();
            frmWorkStreams.PLSMWorkstreamsTabControl.SelectedIndex = 1;
        }

        private void btnWorkStreamsContractConcluding_Click(object sender, EventArgs e)
        {
            frmPLSM_Workstreams frmWorkStreams = new frmPLSM_Workstreams();
            frmWorkStreams.Show();
            frmWorkStreams.PLSMWorkstreamsTabControl.SelectedIndex = 1;
        }

        private void panel41_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnWorkStreamsDetailedDesign_Click(object sender, EventArgs e)
        {
            frmPLSM_Workstreams frmWorkStreams = new frmPLSM_Workstreams();
            frmWorkStreams.Show();
            frmWorkStreams.PLSMWorkstreamsTabControl.SelectedIndex = 1;
        }

        private void btnWorkStreamsTransfer_Click(object sender, EventArgs e)
        {
            frmPLSM_Workstreams frmWorkStreams = new frmPLSM_Workstreams();
            frmWorkStreams.Show();
            frmWorkStreams.PLSMWorkstreamsTabControl.SelectedIndex = 1;
        }

        private void btnWorkStreamsCloseOut_Click(object sender, EventArgs e)
        {
            frmPLSM_Workstreams frmWorkStreams = new frmPLSM_Workstreams();
            frmWorkStreams.Show();
            frmWorkStreams.PLSMWorkstreamsTabControl.SelectedIndex = 1;
        }

        private void btnWorkStreamsEvaluate_Click(object sender, EventArgs e)
        {
            frmPLSM_Workstreams frmWorkStreams = new frmPLSM_Workstreams();
            frmWorkStreams.Show();
            frmWorkStreams.PLSMWorkstreamsTabControl.SelectedIndex = 1;
        }

        private void BtnCorporateProgramManagement_Click(object sender, EventArgs e)
        {
            this.Hide();
            PLSM_ProcessGroupingGovernance_Interfaces plsm = new PLSM_ProcessGroupingGovernance_Interfaces("tbpControllingAStage");
            plsm.Show();
        }

        private void BtnPortfolioManagementAndExecution_Click(object sender, EventArgs e)
        {
            this.Hide();
            PLSM_ProcessGroupingGovernance_Interfaces plsm = new PLSM_ProcessGroupingGovernance_Interfaces("tbpControllingAStage");
            plsm.Show();
        }

        private void btnOppScreenObj_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Identify business and technical deficiencies, needs and oppertunities. Develop, screen & prioritise possible investments to address these, and plan the way forward for the most promising options remaining.");
        }

        private void btnPreFeasObj_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Establish the value of alternatives features of the solution to the business, and specify the business requiremetns. Determine the nature of the solution and assess benefits, cost, risk and external factors, to narrow down the options.");
        }

        private void btnFeasabilityObjective_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Apply high-level systems engineering process to define the list of alternative solutions, verify technical feasibility, business viability and architectural alignment of at least one solution. Quantify technical parameters, costs and benefits.");
        }

        private void btnBusinessPlanObjective_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Complete the systems engineering design and produce a set of specific requirements for hand-over to potential suppliers or specific design disciplines.");
        }

        private void btnProjectExeObj_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Identify business and technical deficiencies, needs and oppertunities. Develop, screen & prioritise possible investments to address these, and plan the way forward for the most promising options remaining. ");
        }

        private void btnContractConcludingObj_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Establish equitable and administratable contract(s) for all aspects of the project, in accordance with predetermined contracting strategies.");
        }

        private void btnDeteailedDesgin_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Design components in detail to meet their specifications.");
        }

        private void btnImpObj_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Close the project in terms of contracts and contractors' liabilities. Close out any transactions after the guarantee period. Evaluate and report on the execution and management of the project.");
        }

        private void btnTransferObj_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Put the asset or solution into service, demonstrate that contractual commitments were met, and hand over responsibility for the asset to the Client.");
        }

        private void btnCloseObj_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Close the project in terms of contracts and contractors' liabilities. Close out any transactions after the guarantee period. Evaluate and report on the execution and management of the project.");
        }

        private void btnEvalObj_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Close the project in terms of contracts and contractors' liabilities. Close out any transactions after the guarantee period. Evaluate and report on the execution and management of the project.");
        }
    }
}
