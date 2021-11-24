using Newtonsoft.Json;
using ProjectManagementToolkit.Properties;
using ProjectManagementToolkit.Utility;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using ProjectManagementToolkit.MPMM.MPMM_Document_Forms;
using ProjectManagementToolkit.MPMM.MPMM_Document_Models;
using ProjectManagementToolkit.Classes;
using System.Drawing.Drawing2D;
using System.Globalization;


namespace ProjectManagementToolkit
{
    public partial class ProjectDashboard : Form
    {
        ProjectModel projectModel = new ProjectModel();
        public ProjectDashboard()
        {
            InitializeComponent();
        }

        // Made variables public for 'Complete' buttons to work
        double initationProgressVal = 0;
        double planningProgressVal = 0;
        double executionProgressVal = 0;
        double closingProgressVal = 0;
        double initationPercentage = 0;
        double planningPercentage = 0;
        double executionPercentage = 0;
        double closingPercentage = 0;

        // Made lists public for 'Complete' buttons to work
        List<string> closingDocuments = new List<string>();
        List<string> initiationDocuments = new List<string>();
        List<string> planningDocuments = new List<string>();
        List<string> executionDocuments = new List<string>();


        string[] xValues1 = new string[4];
        double[] yValues1 = new double[4];
        double[] yValues2 = new double[4];

        BusinessCaseModel currentBusinessCaseModel;
        VersionControl<BusinessCaseModel> versionControl;

        FeasibilityStudyModel currentFeasibilityStudyModel;
        VersionControl<FeasibilityStudyModel> versionControl1;

        ProjectCharterModel currentProjectCharter;
        VersionControl<ProjectCharterModel> versionControl2;

        JobDescriptionModel currentJobDescription;
        VersionControl<JobDescriptionModel> versionControl3;

        ProjectOfficeChecklistModel currentProjectOfficeChecklist;
        VersionControl<ProjectOfficeChecklistModel> versionControl4;

        PhaseReviewFormInitiationModel currentPhaseReviewFormInitiation;
        VersionControl<PhaseReviewFormInitiationModel> versionControl5;

        TermsOfReferenceModel currentTermOfReference;
        VersionControl<TermsOfReferenceModel> versionControl6;

        List<string> initDocsListStatus = new List<string>();

        // Create Lists for the dates and budgets

        List<(string dueDate, string budget, string startDate, string plannedBudget, string completedDate)> initDocsListDueDate = new List<(string dueDate, string budget, string startDate, string plannedBudget, string completedDate)>();
        List<(string dueDate, string budget, string startDate, string plannedBudget, string completedDate)> planningDocsListDueDate = new List<(string dueDate, string budget, string startDate, string plannedBudget, string completedDate)>();
        List<(string dueDate, string budget, string startDate, string plannedBudget, string completedDate)> executeDocsListDueDate = new List<(string dueDate, string budget, string startDate, string plannedBudget, string completedDate)>();
        List<(string dueDate, string budget, string startDate, string plannedBudget, string completedDate)> closingListDueDate = new List<(string dueDate, string budget, string startDate, string plannedBudget, string completedDate)>();

        private void ProjectDashboard_Load(object sender, EventArgs e)
        {
            pbarOverall.Hide();
            lblOverallProgress.Hide();


            List<string> planningDocsListStatus = new List<string>();
            List<string> executionDocsListStatus = new List<string>();
            List<string> closingDocsListStatus = new List<string>();

            // Lists for keeping track of the completion dates of the forms

            List<string> initDocsCompleteDate = new List<string>();
            List<string> planningDocsCompleteDate = new List<string>();
            List<string> executionDocsCompleteDate = new List<string>();
            List<string> closingDocsCompleteDate = new List<string>();

            string json = JsonHelper.loadProjectInfo(Settings.Default.Username);
            List<ProjectModel> projectListModel = JsonConvert.DeserializeObject<List<ProjectModel>>(json);
            projectModel = projectModel.getProjectModel(Settings.Default.ProjectID, projectListModel);


            /////////////////////////////////////////////////////////////////INITIATION PHASE/////////////////////////////////////////////////////////////////////////////////////////////////////////

            // Load the budget/date model

            string jsonInitDue = JsonHelper.loadDocument(Settings.Default.ProjectID, "InitDueDateModel");
            InitDueDateModel tInit = JsonConvert.DeserializeObject<InitDueDateModel>(jsonInitDue);
            

            if (tInit != null)
            {
                initDocsListDueDate.Add((tInit.BusinessCaseDD, tInit.BusinessCaseBudget, tInit.BusinessCaseSD, tInit.BusinessCasePlannedBudget, tInit.BusinessCaseCD));
                initDocsListDueDate.Add((tInit.FeasibilityStudyDD, tInit.FeasibilityStudyBudget, tInit.FeasibilityStudySD, tInit.FeasibilityStudyPlannedBudget, tInit.FeasibilityStudyCD));
                initDocsListDueDate.Add((tInit.ProjectCharterDD, tInit.ProjectCharterBudget, tInit.ProjectCharterSD, tInit.ProjectCharterPlannedBudget, tInit.ProjectCharterCD));
                initDocsListDueDate.Add((tInit.JobDescriptionDD, tInit.JobDescriptionBudget, tInit.JobDescriptionSD, tInit.JobDescriptionPlannedBudget, tInit.JobDescriptionCD));
                initDocsListDueDate.Add((tInit.ProjectOfficeCheckListDD, tInit.ProjectOfficeCheckListBudget, tInit.ProjectOfficeCheckListSD, tInit.ProjectOfficeCheckListPlannedBudget, tInit.ProjectOfficeCheckListCD));
                initDocsListDueDate.Add((tInit.PhaseRevieFormInitiationDD, tInit.PhaseRevieFormInitiationBudget, tInit.PhaseRevieFormInitiationSD, tInit.PhaseRevieFormInitiationPlannedBudget, tInit.PhaseRevieFormInitiationCD));
                initDocsListDueDate.Add((tInit.TermOfReferenceDocumentDD, tInit.TermOfReferenceDocumentBudget, tInit.TermOfReferenceDocumentSD, tInit.TermOfReferenceDocumentPlannedBudget, tInit.TermOfReferenceDocumentCD));
            }


            string jsonPlanningDue = JsonHelper.loadDocument(Settings.Default.ProjectID, "PlanningDueDateModel");
            PlanningDueDateModel tPlanning = JsonConvert.DeserializeObject<PlanningDueDateModel>(jsonPlanningDue);

            if(tPlanning != null)
            {
                planningDocsListDueDate.Add((tPlanning.ProjectPlanDD, tPlanning.ProjectPlanBudget, tPlanning.ProjectPlanSD, tPlanning.ProjectPlanPlannedBudget, tPlanning.ProjectPlanCD));
                planningDocsListDueDate.Add((tPlanning.ResourcePlanDD, tPlanning.ResourcePlanBudget, tPlanning.ResourcePlanSD, tPlanning.ResourcePlanPlannedBudget, tPlanning.ResourcePlanCD));
                planningDocsListDueDate.Add((tPlanning.FinancialPlanDD, tPlanning.FinancialPlanBudget, tPlanning.FinancialPlanSD, tPlanning.FinancialPlanPlannedBudget, tPlanning.FinancialPlanCD));
                planningDocsListDueDate.Add((tPlanning.QualityPlanDD, tPlanning.QualityPlanBudget, tPlanning.QualityPlanSD, tPlanning.QualityPlanPlannedBudget, tPlanning.QualityPlanCD));
                planningDocsListDueDate.Add((tPlanning.RiskPlanDD, tPlanning.RiskPlanBudget, tPlanning.RiskPlanSD, tPlanning.RiskPlanPlannedBudget, tPlanning.RiskPlanCD));
                planningDocsListDueDate.Add((tPlanning.AcceptancePlanDD, tPlanning.AcceptancePlanBudget, tPlanning.AcceptancePlanSD, tPlanning.AcceptancePlanPlannedBudget, tPlanning.AcceptancePlanCD));
                planningDocsListDueDate.Add((tPlanning.CommunicationPlanDD, tPlanning.CommunicationPlanBudget, tPlanning.CommunicationPlanSD, tPlanning.CommunicationPlanPlannedBudget, tPlanning.CommunicationPlanCD));
                planningDocsListDueDate.Add((tPlanning.ProcurementPlanDD, tPlanning.ProcurementPlanBudget, tPlanning.ProcurementPlanSD, tPlanning.ProcurementPlanPlannedBudget, tPlanning.ProcurementPlanCD));
                planningDocsListDueDate.Add((tPlanning.StatementOfWorkDD, tPlanning.StatementOfWorkBudget, tPlanning.StatementOfWorkSD, tPlanning.StatementOfWorkPlannedBudget, tPlanning.StatementOfWorkCD));
                planningDocsListDueDate.Add((tPlanning.RequestForInformationDD, tPlanning.RequestForInformationBudget, tPlanning.RequestForInformationSD, tPlanning.RequestForInformationPlannedBudget, tPlanning.RequestForInformationCD));
                planningDocsListDueDate.Add((tPlanning.SupplierContractDD, tPlanning.SupplierContractBudget, tPlanning.SupplierContractSD, tPlanning.SupplierContractPlannedBudget, tPlanning.SupplierContractCD));
                planningDocsListDueDate.Add((tPlanning.RequestForProposalDD, tPlanning.RequestForProposalBudget, tPlanning.RequestForProposalSD, tPlanning.RequestForProposalPlannedBudget, tPlanning.RequestForProposalCD));
                planningDocsListDueDate.Add((tPlanning.PhaseReviewPlanningDD, tPlanning.PhaseReviewPlanningBudget, tPlanning.PhaseReviewPlanningSD, tPlanning.PhaseReviewPlanningPlannedBudget, tPlanning.PhaseReviewPlanningCD));
            }
            

            string jsonExecuteDue = JsonHelper.loadDocument(Settings.Default.ProjectID, "ExecutionDueDateModel");
            ExecutionDueDateModel tExecute = JsonConvert.DeserializeObject<ExecutionDueDateModel>(jsonExecuteDue);

            if(tExecute != null)
            {
                executeDocsListDueDate.Add((tExecute.TimeMangementDD, tExecute.TimeMangementBudget, tExecute.TimeMangementSD, tExecute.TimeMangementPlannedBudget, tExecute.TimeMangementCD));
                executeDocsListDueDate.Add((tExecute.TimeSheetDD, tExecute.TimeSheetBudget, tExecute.TimeSheetSD, tExecute.TimeSheetPlannedBudget, tExecute.TimeSheetCD));
                executeDocsListDueDate.Add((tExecute.TimeSheetRegisterDD, tExecute.TimeSheetRegisterBudget, tExecute.TimeSheetRegisterSD, tExecute.TimeSheetRegisterPlannedBudget, tExecute.TimeSheetRegisterCD));
                executeDocsListDueDate.Add((tExecute.CostManagementProcessDD, tExecute.CostManagementProcessBudget, tExecute.CostManagementProcessSD, tExecute.CostManagementProcessPlannedBudget, tExecute.CostManagementProcessCD));
                executeDocsListDueDate.Add((tExecute.ExpenseFormDD, tExecute.ExpenseFormBudget, tExecute.ExpenseFormSD, tExecute.ExpenseFormPlannedBudget, tExecute.ExpenseFormCD));
                executeDocsListDueDate.Add((tExecute.ExpenseRegisterDD, tExecute.ExpenseRegisterBudget, tExecute.ExpenseRegisterSD, tExecute.ExpenseRegisterPlannedBudget, tExecute.ExpenseRegisterCD));
                executeDocsListDueDate.Add((tExecute.QualityManagementDD, tExecute.QualityManagementBudget, tExecute.QualityManagementSD, tExecute.QualityManagementPlannedBudget, tExecute.QualityManagementCD));
                executeDocsListDueDate.Add((tExecute.QualityReviewPlanDD, tExecute.QualityReviewPlanBudget, tExecute.QualityReviewPlanSD, tExecute.QualityReviewPlanPlannedBudget, tExecute.QualityReviewPlanCD));
                executeDocsListDueDate.Add((tExecute.QualityReviewFormDD, tExecute.QualityReviewFormBudget, tExecute.QualityReviewFormSD, tExecute.QualityReviewFormPlannedBudget, tExecute.QualityReviewFormCD));
                executeDocsListDueDate.Add((tExecute.ChangeManagementProcessDD, tExecute.ChangeManagementProcessBudget, tExecute.ChangeManagementProcessSD, tExecute.ChangeManagementProcessPlannedBudget, tExecute.ChangeManagementProcessCD));
                executeDocsListDueDate.Add((tExecute.ChangeRequestFormDD, tExecute.ChangeRequestFormBudget, tExecute.ChangeRequestFormSD, tExecute.ChangeRequestFormPlannedBudget, tExecute.ChangeRequestFormCD));
                executeDocsListDueDate.Add((tExecute.ChangeRequestRegisterDD, tExecute.ChangeRequestRegisterBudget, tExecute.ChangeRequestRegisterSD, tExecute.ChangeRequestRegisterPlannedBudget, tExecute.ChangeRequestRegisterCD));
                executeDocsListDueDate.Add((tExecute.RiskManagamentProcessDD, tExecute.RiskManagamentProcessBudget, tExecute.RiskManagamentProcessSD, tExecute.RiskManagamentProcessPlannedBudget, tExecute.RiskManagamentProcessCD));
                executeDocsListDueDate.Add((tExecute.RiskFormDD, tExecute.RiskFormBudget, tExecute.RiskFormSD, tExecute.RiskFormPlannedBudget, tExecute.RiskFormCD));
                executeDocsListDueDate.Add((tExecute.RiskRegisterDD, tExecute.RiskRegisterBudget, tExecute.RiskRegisterSD, tExecute.RiskRegisterPlannedBudget, tExecute.RiskRegisterCD));
                executeDocsListDueDate.Add((tExecute.IssueManagementProcessDD, tExecute.IssueManagementProcessBudget, tExecute.IssueManagementProcessSD, tExecute.IssueManagementProcessPlannedBudget, tExecute.IssueManagementProcessCD));
                executeDocsListDueDate.Add((tExecute.IssueFormDD, tExecute.IssueFormBudget, tExecute.IssueFormSD, tExecute.IssueFormPlannedBudget, tExecute.IssueFormCD));
                executeDocsListDueDate.Add((tExecute.IssueRegisterDD, tExecute.IssueRegisterBudget, tExecute.IssueRegisterSD, tExecute.IssueRegisterPlannedBudget, tExecute.IssueRegisterCD));
                executeDocsListDueDate.Add((tExecute.PurchaseOrderDD, tExecute.PurchaseOrderBudget, tExecute.PurchaseOrderSD, tExecute.PurchaseOrderPlannedBudget, tExecute.PurchaseOrderCD));
                executeDocsListDueDate.Add((tExecute.ProcurementRegisterDD, tExecute.ProcurementRegisterBudget, tExecute.ProcurementRegisterSD, tExecute.ProcurementRegisterPlannedBudget, tExecute.ProcurementRegisterCD));
                executeDocsListDueDate.Add((tExecute.AcceptanceManagementProcessDD, tExecute.AcceptanceManagementProcessBudget, tExecute.AcceptanceManagementProcessSD, tExecute.AcceptanceManagementProcessPlannedBudget, tExecute.AcceptanceManagementProcessCD));
                executeDocsListDueDate.Add((tExecute.AcceptanceFormDD, tExecute.AcceptanceFormBudget, tExecute.AcceptanceFormSD, tExecute.AcceptanceFormPlannedBudget, tExecute.AcceptanceFormCD));
                executeDocsListDueDate.Add((tExecute.AcceptanceRegisterDD, tExecute.AcceptanceRegisterBudget, tExecute.AcceptanceRegisterSD, tExecute.AcceptanceRegisterPlannedBudget, tExecute.AcceptanceRegisterCD));
                executeDocsListDueDate.Add((tExecute.CommunicationsManagementProcessDD, tExecute.CommunicationsManagementProcessBudget, tExecute.CommunicationsManagementProcessSD, tExecute.CommunicationsManagementProcessPlannedBudget, tExecute.CommunicationsManagementProcessCD));
                executeDocsListDueDate.Add((tExecute.ProjectStatusReportDD, tExecute.ProjectStatusReportBudget, tExecute.ProjectStatusReportSD, tExecute.ProjectStatusReportPlannedBudget, tExecute.ProjectStatusReportCD));
                executeDocsListDueDate.Add((tExecute.CommunicationsRegisterDD, tExecute.CommunicationsRegisterBudget, tExecute.CommunicationsRegisterSD, tExecute.CommunicationsRegisterPlannedBudget, tExecute.CommunicationsRegisterCD));
                executeDocsListDueDate.Add((tExecute.PhaseReviewExeDD, tExecute.PhaseReviewExeBudget, tExecute.PhaseReviewExeSD, tExecute.PhaseReviewExePlannedBudget, tExecute.PhaseReviewExeCD));
            }

           

            string jsonClosingDue = JsonHelper.loadDocument(Settings.Default.ProjectID, "ClosingDueDateModel");
            ClosingDueDateModel tClosing = JsonConvert.DeserializeObject<ClosingDueDateModel>(jsonClosingDue);

            if(tClosing != null)
            {
                closingListDueDate.Add((tClosing.ProjectClosureReportDD, tClosing.ProjectClosureReportBudget, tClosing.ProjectClosureReportSD, tClosing.ProjectClosureReportPlannedBudget, tClosing.ProjectClosureReportCD)); ;
                closingListDueDate.Add((tClosing.PostImplementationReviewDD, tClosing.PostImplementationReviewBudget, tClosing.PostImplementationReviewSD, tClosing.PostImplementationReviewPlannedBudget, tClosing.PostImplementationReviewCD));
            }

           
            ////////BUSINESSCASE////////
            //Verander Json
            string json1 = JsonHelper.loadDocument(Settings.Default.ProjectID, "BusinessCase");

            //Check versions
            versionControl = JsonConvert.DeserializeObject<VersionControl<BusinessCaseModel>>(json1);
            //Get current businesscaseModel
            // BusinessCaseModel currentBusinessCaseModel;

            if (versionControl != null)
            {
                currentBusinessCaseModel = JsonConvert.DeserializeObject<BusinessCaseModel>(versionControl.getLatest(versionControl.DocumentModels));
                initDocsListStatus.Add(currentBusinessCaseModel.Progress);
                initDocsCompleteDate.Add(currentBusinessCaseModel.completeDate);
            }
            else
            {
                //IsBusinessCaseModelDone = "";
                initDocsListStatus.Add("");
                initDocsCompleteDate.Add("");
            }

            //////FEASIBILITY STUDY/////////
            string json2 = JsonHelper.loadDocument(Settings.Default.ProjectID, "FeasibilityStudy");
            versionControl1 = JsonConvert.DeserializeObject<VersionControl<FeasibilityStudyModel>>(json2);

            if (versionControl1 != null)
            {
                currentFeasibilityStudyModel = JsonConvert.DeserializeObject<FeasibilityStudyModel>(versionControl1.getLatest(versionControl1.DocumentModels));
                initDocsListStatus.Add(currentFeasibilityStudyModel.FeasibilityStudyProgress);
                initDocsCompleteDate.Add(currentFeasibilityStudyModel.completedDate);
            }
            else
            {
                initDocsListStatus.Add("");
                initDocsCompleteDate.Add("");
            }

            //////PROJECT CHARTER/////////
            string json3 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectCharter");
            versionControl2 = JsonConvert.DeserializeObject<VersionControl<ProjectCharterModel>>(json3);


            if (versionControl2 != null)
            {
                currentProjectCharter = JsonConvert.DeserializeObject<ProjectCharterModel>(versionControl2.getLatest(versionControl2.DocumentModels));
                initDocsListStatus.Add(currentProjectCharter.ProjectCharterProgress);
                initDocsCompleteDate.Add(currentProjectCharter.completedDate);
            }
            else
            {
                initDocsListStatus.Add("");
                initDocsCompleteDate.Add("");
            }

            //////JOB DESCRIPTION/////////
            string json4 = JsonHelper.loadDocument(Settings.Default.ProjectID, "JobDescription");
            versionControl3 = JsonConvert.DeserializeObject<VersionControl<JobDescriptionModel>>(json4);

            if (versionControl3 != null)
            {
                currentJobDescription = JsonConvert.DeserializeObject<JobDescriptionModel>(versionControl3.getLatest(versionControl3.DocumentModels));
                initDocsListStatus.Add(currentJobDescription.JobDescriptionProgress);
                initDocsCompleteDate.Add(currentJobDescription.completedDate);

            }
            else
            {
                initDocsListStatus.Add("");
                initDocsCompleteDate.Add("");
            }

            //////PROJECT OFFICE CHECKLIST/////////
            string json5 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectOfficeCheckList");
            versionControl4 = JsonConvert.DeserializeObject<VersionControl<ProjectOfficeChecklistModel>>(json5);


            if (versionControl4 != null)
            {
                currentProjectOfficeChecklist = JsonConvert.DeserializeObject<ProjectOfficeChecklistModel>(versionControl4.getLatest(versionControl4.DocumentModels));
                initDocsListStatus.Add(currentProjectOfficeChecklist.ProjectOfficeCheckListProgress);
                initDocsCompleteDate.Add(currentProjectOfficeChecklist.completedDate);
            }
            else
            {
                initDocsListStatus.Add("");
                initDocsCompleteDate.Add("");
            }
                


            //////PHASE REVIEW FORM INITIATION/////////
            string json6 = JsonHelper.loadDocument(Settings.Default.ProjectID, "PhaseReviewFormInitiation");
            versionControl5 = JsonConvert.DeserializeObject<VersionControl<PhaseReviewFormInitiationModel>>(json6);


            if (versionControl5 != null)
            {
                currentPhaseReviewFormInitiation = JsonConvert.DeserializeObject<PhaseReviewFormInitiationModel>(versionControl5.getLatest(versionControl5.DocumentModels));
                initDocsListStatus.Add(currentPhaseReviewFormInitiation.PhaseReviewFormInitiationProgress);
                initDocsCompleteDate.Add(currentPhaseReviewFormInitiation.completedDate);

            }
            else
            {
                initDocsListStatus.Add("");
                initDocsCompleteDate.Add("");
            }


            //////TERMS OF REFERENCE/////////
            string json7 = JsonHelper.loadDocument(Settings.Default.ProjectID, "TermOfReferenceDocument");
            versionControl6 = JsonConvert.DeserializeObject<VersionControl<TermsOfReferenceModel>>(json7);


            if (versionControl6 != null)
            {
                currentTermOfReference = JsonConvert.DeserializeObject<TermsOfReferenceModel>(versionControl6.getLatest(versionControl6.DocumentModels));
                initDocsListStatus.Add(currentTermOfReference.TermOfReferenceProgress);
                initDocsCompleteDate.Add(currentTermOfReference.completedDate);
            }
            else
            {
                initDocsListStatus.Add("");
                initDocsCompleteDate.Add("");
            }

            /////////////////////////////////////////////////////////////////PLANNING PHASE/////////////////////////////////////////////////////////////////////////////////////////////////////////

            //////ProjectPlan/////////
            string json8 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectPlan");
            VersionControl<ProjectPlanModel> versionControl7 = JsonConvert.DeserializeObject<VersionControl<ProjectPlanModel>>(json8);
            ProjectPlanModel currentProjectPlan;

            if (versionControl7 != null)
            {
                currentProjectPlan = JsonConvert.DeserializeObject<ProjectPlanModel>(versionControl7.getLatest(versionControl7.DocumentModels));
                planningDocsListStatus.Add(currentProjectPlan.projectPlanProgress);
                planningDocsCompleteDate.Add(currentProjectPlan.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }
                
            

            //////ResourcePlan/////////
            string json9 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ResourcePlan");
            VersionControl<ResourcePlanModel> versionControl8 = JsonConvert.DeserializeObject<VersionControl<ResourcePlanModel>>(json9);
            ResourcePlanModel currentResourcePlan;

            if (versionControl8 != null)
            {
                currentResourcePlan = JsonConvert.DeserializeObject<ResourcePlanModel>(versionControl8.getLatest(versionControl8.DocumentModels));
                planningDocsListStatus.Add(currentResourcePlan.ResourcePlanProgress);
                planningDocsCompleteDate.Add(currentResourcePlan.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }

                //////FinancialPlan/////////
                string json10 = JsonHelper.loadDocument(Settings.Default.ProjectID, "FinancialPlan");
            VersionControl<FinancialPlanModel> versionControl9 = JsonConvert.DeserializeObject<VersionControl<FinancialPlanModel>>(json10);
            FinancialPlanModel currentFinancialPlan;

            if (versionControl9 != null)
            {
                currentFinancialPlan = JsonConvert.DeserializeObject<FinancialPlanModel>(versionControl9.getLatest(versionControl9.DocumentModels));
                planningDocsListStatus.Add(currentFinancialPlan.FinancialPlanProgress);
                planningDocsCompleteDate.Add(currentFinancialPlan.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }

            //////QualityPlan/////////
            string json11 = JsonHelper.loadDocument(Settings.Default.ProjectID, "QualityPlan");
            VersionControl<QualityPlanModel> versionControl10 = JsonConvert.DeserializeObject<VersionControl<QualityPlanModel>>(json11);
            QualityPlanModel currentQualityPlan;

            if (versionControl10 != null)
            {
                currentQualityPlan = JsonConvert.DeserializeObject<QualityPlanModel>(versionControl10.getLatest(versionControl10.DocumentModels));
                planningDocsListStatus.Add(currentQualityPlan.QualityPlanProgress);
                planningDocsCompleteDate.Add(currentQualityPlan.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }

            //////RiskPlan/////////
            string json12 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RiskPlan");
            VersionControl<RiskPlanModel> versionControl11 = JsonConvert.DeserializeObject<VersionControl<RiskPlanModel>>(json12);
            RiskPlanModel currentRiskPlan;

            if (versionControl11 != null)
            {
                currentRiskPlan = JsonConvert.DeserializeObject<RiskPlanModel>(versionControl11.getLatest(versionControl11.DocumentModels));
                planningDocsListStatus.Add(currentRiskPlan.RiskPlanProgress);
                planningDocsCompleteDate.Add(currentRiskPlan.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }

                //////AcceptancePlan/////////
                string json13 = JsonHelper.loadDocument(Settings.Default.ProjectID, "AcceptancePlan");
            VersionControl<AcceptancePlanModel> versionControl12 = JsonConvert.DeserializeObject<VersionControl<AcceptancePlanModel>>(json13);
            AcceptancePlanModel currentAcceptancePlan;

            if (versionControl12 != null)
            {
                currentAcceptancePlan = JsonConvert.DeserializeObject<AcceptancePlanModel>(versionControl12.getLatest(versionControl12.DocumentModels));
                planningDocsListStatus.Add(currentAcceptancePlan.AcceptancePlanProgress);
                planningDocsCompleteDate.Add(currentAcceptancePlan.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }

                //////CommunicationPlan/////////
                string json14 = JsonHelper.loadDocument(Settings.Default.ProjectID, "CommunicationPlan");
            VersionControl<CommunicationsPlanModel> versionControl13 = JsonConvert.DeserializeObject<VersionControl<CommunicationsPlanModel>>(json14);
            CommunicationsPlanModel currentCommunicationPlan;

            if (versionControl13 != null)
            {
                currentCommunicationPlan = JsonConvert.DeserializeObject<CommunicationsPlanModel>(versionControl13.getLatest(versionControl13.DocumentModels));
                planningDocsListStatus.Add(currentCommunicationPlan.CommunicationPlanProgress);
                planningDocsCompleteDate.Add(currentCommunicationPlan.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }

                //////ProcurementPlan/////////
                string json15 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProcurementPlan");
            VersionControl<ProcurementPlanModel> versionControl14 = JsonConvert.DeserializeObject<VersionControl<ProcurementPlanModel>>(json15);
            ProcurementPlanModel currentProcurementPlan;

            if (versionControl14 != null)
            {
                currentProcurementPlan = JsonConvert.DeserializeObject<ProcurementPlanModel>(versionControl14.getLatest(versionControl14.DocumentModels));
                planningDocsListStatus.Add(currentProcurementPlan.ProcurementPlanProgress);
                planningDocsCompleteDate.Add(currentProcurementPlan.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }

                //////StatementOfWork/////////
                string json16 = JsonHelper.loadDocument(Settings.Default.ProjectID, "StatementOfWork");
            VersionControl<StatementOfWorkModel> versionControl15 = JsonConvert.DeserializeObject<VersionControl<StatementOfWorkModel>>(json16);
            StatementOfWorkModel currentStatementOfWork;

            if (versionControl15 != null)
            {
                currentStatementOfWork = JsonConvert.DeserializeObject<StatementOfWorkModel>(versionControl15.getLatest(versionControl15.DocumentModels));
                planningDocsListStatus.Add(currentStatementOfWork.StatementOfWorkProgress);
                planningDocsCompleteDate.Add(currentStatementOfWork.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }

                //////RequestForInformation/////////
                string json17 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RequestForInformation");
            VersionControl<RequestForInformationModel> versionControl16 = JsonConvert.DeserializeObject<VersionControl<RequestForInformationModel>>(json17);
            RequestForInformationModel currentRequestForInformation;

            if (versionControl16 != null)
            {
                currentRequestForInformation = JsonConvert.DeserializeObject<RequestForInformationModel>(versionControl16.getLatest(versionControl16.DocumentModels));
                planningDocsListStatus.Add(currentRequestForInformation.RequestForInformationProgress);
                planningDocsCompleteDate.Add(currentRequestForInformation.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }

                //////SupplierContract/////////
                string json18 = JsonHelper.loadDocument(Settings.Default.ProjectID, "SupplierContract");
            VersionControl<SupplierContractModel> versionControl17 = JsonConvert.DeserializeObject<VersionControl<SupplierContractModel>>(json18);
            SupplierContractModel currentSupplierContract;

            if (versionControl17 != null)
            {
                currentSupplierContract = JsonConvert.DeserializeObject<SupplierContractModel>(versionControl17.getLatest(versionControl17.DocumentModels));
                planningDocsListStatus.Add(currentSupplierContract.SupplierContractProgress);
                planningDocsCompleteDate.Add(currentSupplierContract.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }

                //////RequestForProposal/////////
                string json19 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RequestForProposal");
            VersionControl<RequestForProposalModel> versionControl18 = JsonConvert.DeserializeObject<VersionControl<RequestForProposalModel>>(json19);
            RequestForProposalModel currentRequestForProposal;

            if (versionControl18 != null)
            {
                currentRequestForProposal = JsonConvert.DeserializeObject<RequestForProposalModel>(versionControl18.getLatest(versionControl18.DocumentModels));
                planningDocsListStatus.Add(currentRequestForProposal.RequestForProposalProgress);
                planningDocsCompleteDate.Add(currentRequestForProposal.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }

                //////PhaseReviewPlanning/////////
                string json20 = JsonHelper.loadDocument(Settings.Default.ProjectID, "PhaseReviewPlanning");
            VersionControl<PhaseReviewPlanningModel> versionControl19 = JsonConvert.DeserializeObject<VersionControl<PhaseReviewPlanningModel>>(json20);
            PhaseReviewPlanningModel currentPhaseReviewPlanning;

            if (versionControl19 != null)
            {
                currentPhaseReviewPlanning = JsonConvert.DeserializeObject<PhaseReviewPlanningModel>(versionControl19.getLatest(versionControl19.DocumentModels));
                planningDocsListStatus.Add(currentPhaseReviewPlanning.PhaseReviewPlanningProgress);
                planningDocsCompleteDate.Add(currentPhaseReviewPlanning.completedDate);
            }
            else
            {
                planningDocsListStatus.Add("");
                planningDocsCompleteDate.Add("");
            }


                /////////////////////////////////////////////////////////////////EXECUTION PHASE/////////////////////////////////////////////////////////////////////////////////////////////////////////

                //////TimeManagement Process/////////
                string json21 = JsonHelper.loadDocument(Settings.Default.ProjectID, "TimeMangement");
            VersionControl<TimeMangementProcessModel> versionControl20 = JsonConvert.DeserializeObject<VersionControl<TimeMangementProcessModel>>(json21);
            TimeMangementProcessModel currentTimeManagementProcess;

            if (versionControl20 != null)
            {
                currentTimeManagementProcess = JsonConvert.DeserializeObject<TimeMangementProcessModel>(versionControl20.getLatest(versionControl20.DocumentModels));
                executionDocsListStatus.Add(currentTimeManagementProcess.TimeManagementProcessProgress);
                executionDocsCompleteDate.Add(currentTimeManagementProcess.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


                //////TimeSheet Process/////////
            string json22 = JsonHelper.loadDocument(Settings.Default.ProjectID, "TimeSheet");
            VersionControl<TimeSheetModel> versionControl21 = JsonConvert.DeserializeObject<VersionControl<TimeSheetModel>>(json22);
            TimeSheetModel currentTimeSheet;

            if (versionControl21 != null)
            {
                currentTimeSheet = JsonConvert.DeserializeObject<TimeSheetModel>(versionControl21.getLatest(versionControl21.DocumentModels));
                executionDocsListStatus.Add(currentTimeSheet.TimeSheetProgress);
                executionDocsCompleteDate.Add(currentTimeSheet.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


                //////TimeSheet Register/////////
                string json23 = JsonHelper.loadDocument(Settings.Default.ProjectID, "TimeSheetRegister");
            VersionControl<TimesheetRegisterModel.TimesheetEntry> versionControl22 = JsonConvert.DeserializeObject<VersionControl<TimesheetRegisterModel.TimesheetEntry>>(json23);
            TimesheetRegisterModel.TimesheetEntry currentTimeSheetRegister;

            if (versionControl22 != null)
            {
                currentTimeSheetRegister = JsonConvert.DeserializeObject<TimesheetRegisterModel.TimesheetEntry>(versionControl22.getLatest(versionControl22.DocumentModels));
                executionDocsListStatus.Add(currentTimeSheetRegister.TimeSheetRegisterProgress);
                executionDocsCompleteDate.Add(currentTimeSheetRegister.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


                //////CostManagement Process/////////
                string json24 = JsonHelper.loadDocument(Settings.Default.ProjectID, "CostManagementProcess");
            VersionControl<CostManagementProcessModel> versionControl23 = JsonConvert.DeserializeObject<VersionControl<CostManagementProcessModel>>(json24);
            CostManagementProcessModel currentCostManagementProcess;

            if (versionControl23 != null)
            {
                currentCostManagementProcess = JsonConvert.DeserializeObject<CostManagementProcessModel>(versionControl23.getLatest(versionControl23.DocumentModels));
                executionDocsListStatus.Add(currentCostManagementProcess.CostManagementProcessProgress);
                executionDocsCompleteDate.Add(currentCostManagementProcess.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


                //////ExpenseForm Process/////////
                string json25 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ExpenseForm");
            VersionControl<ExpenseFormModel> versionControl24 = JsonConvert.DeserializeObject<VersionControl<ExpenseFormModel>>(json25);
            ExpenseFormModel currentExpenseFormProcess;

            if (versionControl24 != null)
            {
                currentExpenseFormProcess = JsonConvert.DeserializeObject<ExpenseFormModel>(versionControl24.getLatest(versionControl24.DocumentModels));
                executionDocsListStatus.Add(currentExpenseFormProcess.ExpenseFormProgress);
                executionDocsCompleteDate.Add(currentExpenseFormProcess.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


                //////ExpenseRegister Process/////////
                string json26 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ExpenseRegister");
            VersionControl<ProjectManagementToolkit.MPMM.MPMM_Document_Models.ExpenseRegister.ExpenseEntry> versionControl25 = JsonConvert.DeserializeObject<VersionControl<ProjectManagementToolkit.MPMM.MPMM_Document_Models.ExpenseRegister.ExpenseEntry>>(json26);
            ProjectManagementToolkit.MPMM.MPMM_Document_Models.ExpenseRegister.ExpenseEntry currentExpenseRegister;

            if (versionControl25 != null)
            {
                currentExpenseRegister = JsonConvert.DeserializeObject<ProjectManagementToolkit.MPMM.MPMM_Document_Models.ExpenseRegister.ExpenseEntry>(versionControl25.getLatest(versionControl25.DocumentModels));
                executionDocsListStatus.Add(currentExpenseRegister.ExpenseRegisterProgress);
                executionDocsCompleteDate.Add(currentExpenseRegister.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


                //////QualityManagement Process/////////
                string json27 = JsonHelper.loadDocument(Settings.Default.ProjectID, "QualityManagement");
            VersionControl<QualityManagementProcessModel> versionControl26 = JsonConvert.DeserializeObject<VersionControl<QualityManagementProcessModel>>(json27);
            QualityManagementProcessModel currentQualityMnagementProcess;

            if (versionControl26 != null)
            {
                currentQualityMnagementProcess = JsonConvert.DeserializeObject<QualityManagementProcessModel>(versionControl26.getLatest(versionControl26.DocumentModels));
                executionDocsListStatus.Add(currentQualityMnagementProcess.QualityManagementProcessProgress);
                executionDocsCompleteDate.Add(currentQualityMnagementProcess.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }



            //////QualityReviewPlan Process/////////
            string json28 = JsonHelper.loadDocument(Settings.Default.ProjectID, "QualityReviewPlan");
            VersionControl<QualityReviewPlanModel> versionControl27 = JsonConvert.DeserializeObject<VersionControl<QualityReviewPlanModel>>(json28);
            QualityReviewPlanModel currentQualityReviewPlan;

            if (versionControl27 != null)
            {
                currentQualityReviewPlan = JsonConvert.DeserializeObject<QualityReviewPlanModel>(versionControl27.getLatest(versionControl27.DocumentModels));
                executionDocsListStatus.Add(currentQualityReviewPlan.QualityReviewPlanProgress);
                executionDocsCompleteDate.Add(currentQualityReviewPlan.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


            //////QualityReviewForm Process/////////
            string json29 = JsonHelper.loadDocument(Settings.Default.ProjectID, "QualityReviewForm");
            VersionControl<QualityRegisterModel.ConformanceOfProcess> versionControl28 = JsonConvert.DeserializeObject<VersionControl<QualityRegisterModel.ConformanceOfProcess>>(json29);
            QualityRegisterModel.ConformanceOfProcess currentQualityReviewForm;

            if (versionControl28 != null)
            {
                currentQualityReviewForm = JsonConvert.DeserializeObject<QualityRegisterModel.ConformanceOfProcess>(versionControl28.getLatest(versionControl28.DocumentModels));
                executionDocsListStatus.Add(currentQualityReviewForm.QualityRegisterProgress);
                executionDocsCompleteDate.Add(currentQualityReviewForm.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


            //////ChangeManagementProcess Process/////////
            string json30 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ChangeManagementProcess");
            VersionControl<ChangeManagementProcessModel> versionControl29 = JsonConvert.DeserializeObject<VersionControl<ChangeManagementProcessModel>>(json30);
            ChangeManagementProcessModel currentChangeManagementProcess;

            if (versionControl29 != null)
            {
                currentChangeManagementProcess = JsonConvert.DeserializeObject<ChangeManagementProcessModel>(versionControl29.getLatest(versionControl29.DocumentModels));
                executionDocsListStatus.Add(currentChangeManagementProcess.ChangeManagementProcessProgress);
                executionDocsCompleteDate.Add(currentChangeManagementProcess.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


            //////ChangeRequestForm Process/////////
            string json31 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ChangeRequestForm");
            VersionControl<ChangeRequestModel> versionControl30 = JsonConvert.DeserializeObject<VersionControl<ChangeRequestModel>>(json31);
            ChangeRequestModel currentChangeRequestForm;

            if (versionControl30 != null)
            {
                currentChangeRequestForm = JsonConvert.DeserializeObject<ChangeRequestModel>(versionControl30.getLatest(versionControl30.DocumentModels));
                executionDocsListStatus.Add(currentChangeRequestForm.ChangeRequestProgress);
                executionDocsCompleteDate.Add(currentChangeRequestForm.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


            //////ChangeRequestRegister Process/////////
            string json32 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ChangeRequestRegister");
            VersionControl<ChangeRegisterModel> versionControl31 = JsonConvert.DeserializeObject<VersionControl<ChangeRegisterModel>>(json32);
            ChangeRegisterModel currentChangeRegister;

            if (versionControl31 != null)
            {
                currentChangeRegister = JsonConvert.DeserializeObject<ChangeRegisterModel>(versionControl31.getLatest(versionControl31.DocumentModels));
                executionDocsListStatus.Add(currentChangeRegister.ChangeRegisterProgress);
                executionDocsCompleteDate.Add(currentChangeRegister.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


            //////RiskManagamentProcess Process/////////
            string json33 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RiskManagamentProcess");
            VersionControl<RiskManagmentProcessModel> versionControl32 = JsonConvert.DeserializeObject<VersionControl<RiskManagmentProcessModel>>(json33);
            RiskManagmentProcessModel currentRiskManagementProcess;

            if (versionControl32 != null)
            {
                currentRiskManagementProcess = JsonConvert.DeserializeObject<RiskManagmentProcessModel>(versionControl32.getLatest(versionControl32.DocumentModels));
                executionDocsListStatus.Add(currentRiskManagementProcess.RiskManagementProcessProgress);
                executionDocsCompleteDate.Add(currentRiskManagementProcess.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }


            //////RiskForm/////////
            string json34 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RiskForm");
            VersionControl<RiskFormModel> versionControl33 = JsonConvert.DeserializeObject<VersionControl<RiskFormModel>>(json34);
            RiskFormModel currentRiskForm;

            if (versionControl33 != null)
            {
                currentRiskForm = JsonConvert.DeserializeObject<RiskFormModel>(versionControl33.getLatest(versionControl33.DocumentModels));
                executionDocsListStatus.Add(currentRiskForm.RiskFormProgress);
                executionDocsCompleteDate.Add(currentRiskForm.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////RiskRegister/////////
            string json35 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RiskRegister");
            VersionControl<RiskRegisterModel> versionControl34 = JsonConvert.DeserializeObject<VersionControl<RiskRegisterModel>>(json35);
            RiskRegisterModel currentRiskRegister;

            if (versionControl34 != null)
            {
                currentRiskRegister = JsonConvert.DeserializeObject<RiskRegisterModel>(versionControl34.getLatest(versionControl34.DocumentModels));
                executionDocsListStatus.Add(currentRiskRegister.RiskRegisterProgress);
                executionDocsCompleteDate.Add(currentRiskRegister.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////IssueManagementProcess/////////
            string json36 = JsonHelper.loadDocument(Settings.Default.ProjectID, "IssueManagementProcess");
            VersionControl<IssueManagementProcessModel> versionControl35 = JsonConvert.DeserializeObject<VersionControl<IssueManagementProcessModel>>(json36);
            IssueManagementProcessModel currentIssueManagementProcess;

            if (versionControl35 != null)
            {
                currentIssueManagementProcess = JsonConvert.DeserializeObject<IssueManagementProcessModel>(versionControl35.getLatest(versionControl35.DocumentModels));
                executionDocsListStatus.Add(currentIssueManagementProcess.IssueManagementProcessProgress);
                executionDocsCompleteDate.Add(currentIssueManagementProcess.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////IssueForm/////////
            string json37 = JsonHelper.loadDocument(Settings.Default.ProjectID, "IssueForm");
            VersionControl<IssueFormModel> versionControl36 = JsonConvert.DeserializeObject<VersionControl<IssueFormModel>>(json37);
            IssueFormModel currentIssueForm;

            if (versionControl36 != null)
            {
                currentIssueForm = JsonConvert.DeserializeObject<IssueFormModel>(versionControl36.getLatest(versionControl36.DocumentModels));
                executionDocsListStatus.Add(currentIssueForm.IssueFormProgress);
                executionDocsCompleteDate.Add(currentIssueForm.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////IssueRegister/////////
            string json38 = JsonHelper.loadDocument(Settings.Default.ProjectID, "IssueRegister");
            VersionControl<IssueRegisterModel> versionControl37 = JsonConvert.DeserializeObject<VersionControl<IssueRegisterModel>>(json38);
            IssueRegisterModel currentIssueRegister;

            if (versionControl37 != null)
            {
                currentIssueRegister = JsonConvert.DeserializeObject<IssueRegisterModel>(versionControl37.getLatest(versionControl37.DocumentModels));
                executionDocsListStatus.Add(currentIssueRegister.IssueRegisterProgress);
                executionDocsCompleteDate.Add(currentIssueRegister.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////PurchaseOrder/////////
            string json39 = JsonHelper.loadDocument(Settings.Default.ProjectID, "PurchaseOrder");
            VersionControl<PurchaseOrderModel> versionControl38 = JsonConvert.DeserializeObject<VersionControl<PurchaseOrderModel>>(json39);
            PurchaseOrderModel currentPurchaseOrder;

            if (versionControl38 != null)
            {
                currentPurchaseOrder = JsonConvert.DeserializeObject<PurchaseOrderModel>(versionControl38.getLatest(versionControl38.DocumentModels));
                executionDocsListStatus.Add(currentPurchaseOrder.PurchaseOrderProgress);
                executionDocsCompleteDate.Add(currentPurchaseOrder.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////ProcurementRegister/////////
            string json40 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProcurementRegister");
            VersionControl<ProcurementRegisterModel> versionControl39 = JsonConvert.DeserializeObject<VersionControl<ProcurementRegisterModel>>(json40);
            ProcurementRegisterModel currentProcurementRegister;

            if (versionControl39 != null)
            {
                currentProcurementRegister = JsonConvert.DeserializeObject<ProcurementRegisterModel>(versionControl39.getLatest(versionControl39.DocumentModels));
                executionDocsListStatus.Add(currentProcurementRegister.ProcurementRegisterProgress);
                executionDocsCompleteDate.Add(currentProcurementRegister.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////AcceptanceManagementProcess/////////
            string json41 = JsonHelper.loadDocument(Settings.Default.ProjectID, "AcceptanceManagementProcess");
            VersionControl<AcceptanceManagementProcessModel> versionControl40 = JsonConvert.DeserializeObject<VersionControl<AcceptanceManagementProcessModel>>(json41);
            AcceptanceManagementProcessModel currentAcceptanceManagementProcess;

            if (versionControl40 != null)
            {
                currentAcceptanceManagementProcess = JsonConvert.DeserializeObject<AcceptanceManagementProcessModel>(versionControl40.getLatest(versionControl40.DocumentModels));
                executionDocsListStatus.Add(currentAcceptanceManagementProcess.AcceptanceManagementProcessProgress);
                executionDocsCompleteDate.Add(currentAcceptanceManagementProcess.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////AcceptanceForm/////////
            string json42 = JsonHelper.loadDocument(Settings.Default.ProjectID, "AcceptanceForm");
            VersionControl<AcceptanceFormModel> versionControl41 = JsonConvert.DeserializeObject<VersionControl<AcceptanceFormModel>>(json42);
            AcceptanceFormModel currentAcceptanceForm;

            if (versionControl41 != null)
            {
                currentAcceptanceForm = JsonConvert.DeserializeObject<AcceptanceFormModel>(versionControl41.getLatest(versionControl41.DocumentModels));
                executionDocsListStatus.Add(currentAcceptanceForm.AcceptanceFormProgress);
                executionDocsCompleteDate.Add(currentAcceptanceForm.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////AcceptanceRegister/////////
            string json43 = JsonHelper.loadDocument(Settings.Default.ProjectID, "AcceptanceRegister");
            VersionControl<AcceptanceRegisterModel> versionControl42 = JsonConvert.DeserializeObject<VersionControl<AcceptanceRegisterModel>>(json43);
            AcceptanceRegisterModel currentAcceptanceRegister;

            if (versionControl42 != null)
            {
                currentAcceptanceRegister = JsonConvert.DeserializeObject<AcceptanceRegisterModel>(versionControl42.getLatest(versionControl42.DocumentModels));
                executionDocsListStatus.Add(currentAcceptanceRegister.AcceptanceRegisterProgress);
                executionDocsCompleteDate.Add(currentAcceptanceRegister.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////CommunicationsManagementProcess/////////
            string json44 = JsonHelper.loadDocument(Settings.Default.ProjectID, "CommunicationsManagementProcess");
            VersionControl<CommunicationsManagementProcessModel> versionControl43 = JsonConvert.DeserializeObject<VersionControl<CommunicationsManagementProcessModel>>(json44);
            CommunicationsManagementProcessModel currentCommunicationsManagementProcess;

            if (versionControl43 != null)
            {
                currentCommunicationsManagementProcess = JsonConvert.DeserializeObject<CommunicationsManagementProcessModel>(versionControl43.getLatest(versionControl43.DocumentModels));
                executionDocsListStatus.Add(currentCommunicationsManagementProcess.CommunicationsManagementProcessProgress);
                executionDocsCompleteDate.Add(currentCommunicationsManagementProcess.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////ProjectStatusReport/////////
            string json45 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectStatusReport");
            VersionControl<ProjectStatusReportModel> versionControl44 = JsonConvert.DeserializeObject<VersionControl<ProjectStatusReportModel>>(json45);
            ProjectStatusReportModel currentProjectStatusReport;

            if (versionControl44 != null)
            {
                currentProjectStatusReport = JsonConvert.DeserializeObject<ProjectStatusReportModel>(versionControl44.getLatest(versionControl44.DocumentModels));
                executionDocsListStatus.Add(currentProjectStatusReport.ProjectStatusReportProgress);
                executionDocsCompleteDate.Add(currentProjectStatusReport.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////CommunicationsRegister/////////
            string json46 = JsonHelper.loadDocument(Settings.Default.ProjectID, "CommunicationsRegister");
            VersionControl<CommunicationRegisterModel> versionControl45 = JsonConvert.DeserializeObject<VersionControl<CommunicationRegisterModel>>(json46);
            CommunicationRegisterModel currentCommunicationsRegister;

            if (versionControl45 != null)
            {
                currentCommunicationsRegister = JsonConvert.DeserializeObject<CommunicationRegisterModel>(versionControl45.getLatest(versionControl45.DocumentModels));
                executionDocsListStatus.Add(currentCommunicationsRegister.CommunicationsRegisterProgress);
                executionDocsCompleteDate.Add(currentCommunicationsRegister.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////PhaseReviewExe/////////
            string json47 = JsonHelper.loadDocument(Settings.Default.ProjectID, "PhaseReviewExe");
            VersionControl<PhaseReviewFormExecutionModel> versionControl46 = JsonConvert.DeserializeObject<VersionControl<PhaseReviewFormExecutionModel>>(json47);
            PhaseReviewFormExecutionModel currentPhaseReviewExe;

            if (versionControl46 != null)
            {
                currentPhaseReviewExe = JsonConvert.DeserializeObject<PhaseReviewFormExecutionModel>(versionControl46.getLatest(versionControl46.DocumentModels));
                executionDocsListStatus.Add(currentPhaseReviewExe.PhaseReviewExeProgress);
                executionDocsCompleteDate.Add(currentPhaseReviewExe.completedDate);
            }
            else
            {
                executionDocsListStatus.Add("");
                executionDocsCompleteDate.Add("");
            }

            //////////////////////////////////////////////////////CLOSING PHASE///////////////////////////////////////////////////////////////////////////////////
            //////ProjectClosureReport/////////
            string json48 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectClosureReport");
            VersionControl<ProjectClosureReportModel> versionControl47 = JsonConvert.DeserializeObject<VersionControl<ProjectClosureReportModel>>(json48);
            ProjectClosureReportModel currentProjectClosureReport;

            if (versionControl47 != null)
            {
                currentProjectClosureReport = JsonConvert.DeserializeObject<ProjectClosureReportModel>(versionControl47.getLatest(versionControl47.DocumentModels));
                closingDocsListStatus.Add(currentProjectClosureReport.ProjectClosureReportProgress);
                closingDocsCompleteDate.Add(currentProjectClosureReport.completedDate);
            }
            else
            {
                closingDocsListStatus.Add("");
                closingDocsCompleteDate.Add("");
            }

            //////PostImplementationReview/////////
            string json49 = JsonHelper.loadDocument(Settings.Default.ProjectID, "PostImplementationReview");
            VersionControl<PostImplementationReviewModel> versionControl48 = JsonConvert.DeserializeObject<VersionControl<PostImplementationReviewModel>>(json49);
            PostImplementationReviewModel currentPostImplementationReview;

            if (versionControl48 != null)
            {
                currentPostImplementationReview = JsonConvert.DeserializeObject<PostImplementationReviewModel>(versionControl48.getLatest(versionControl48.DocumentModels));
                closingDocsListStatus.Add(currentPostImplementationReview.PostImplementationReviewProgress);
                closingDocsCompleteDate.Add(currentPostImplementationReview.completedDate);
            }
            else
            {
                closingDocsListStatus.Add("");
                closingDocsCompleteDate.Add("");
            }

            //Get localdocs
            List<string> localDocuments = getLocalDocuments();

            lblProjectName.Text = projectModel.ProjectName;

            chart1.ChartAreas[0].BackColor = Color.Transparent;
            chart1.Legends[0].BackColor = Color.Transparent;
            chart2.Legends[0].BackColor = Color.Transparent;



            //Counters for completed, uncompleted, and in progress tasks
            int comp = 0, uncomp = 0, inprog = 0, behind = 0;
            int compPlanning = 0, uncompPlanning = 0, inprogPlanning = 0, behindPlanning = 0;
            int compExecution = 0, uncompExecution = 0, inprogExecution = 0, behindExecution = 0;
            int compClosing = 0, uncompClosing = 0, inprogClosing = 0, behindClosing = 0;

            if (localDocuments == null)
            {
                MessageBox.Show("No documents have been added yet.\nPlease save some documents before opening the dashboard", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tabInitiation.Visible = false;
                chart1.Visible = false;
                chart2.Visible = false;
                this.Close();
            }
            else
            {
                tabInitiation.Visible = true;
                chart1.Visible = true;
                chart2.Visible = true;



                ///////////////INITIATION PHASE/////////////////
                initiationDocuments.Add("BusinessCase");
                initiationDocuments.Add("FeasibilityStudy");
                initiationDocuments.Add("ProjectCharter");
                initiationDocuments.Add("JobDescription");
                initiationDocuments.Add("ProjectOfficeCheckList");
                initiationDocuments.Add("PhaseReviewFormInitiation");
                initiationDocuments.Add("TermOfReferenceDocument");

                // executionDocsListStatus.Add("TimeMangement");
                initDocsListStatus.Add("BusinessCase");



                int k = 0;

                for (int i = 0; i < initiationDocuments.Count; i++)
                {
                    dgvInitiation.Rows.Add();
                    dgvInitiation.Rows[i].Cells[0].Value = initiationDocuments[i];                  
                    

                    if (initDocsListDueDate.Count > 0)
                    {
                        if (initDocsListDueDate[i].dueDate != null)
                            dgvInitiation.Rows[i].Cells[3].Value = initDocsListDueDate[i].dueDate;
                        else
                            dgvInitiation.Rows[i].Cells[3].Value = "";

                        if (initDocsListDueDate[i].budget != null)
                            dgvInitiation.Rows[i].Cells[5].Value = initDocsListDueDate[i].budget;
                        else
                            dgvInitiation.Rows[i].Cells[5].Value = "";

                        if (initDocsListDueDate[i].startDate!= null)
                            dgvInitiation.Rows[i].Cells[1].Value = initDocsListDueDate[i].startDate;
                        else
                            dgvInitiation.Rows[i].Cells[1].Value = "";

                        if (initDocsListDueDate[i].completedDate != null)
                            dgvInitiation.Rows[i].Cells[2].Value = initDocsListDueDate[i].completedDate;
                        else
                            dgvInitiation.Rows[i].Cells[2].Value = "";

                        if (initDocsListDueDate[i].plannedBudget != null)
                            dgvInitiation.Rows[i].Cells[4].Value = initDocsListDueDate[i].plannedBudget;
                        else
                            dgvInitiation.Rows[i].Cells[4].Value = "";


                    }
                    else
                    {
                        dgvInitiation.Rows[i].Cells[1].Value = "";
                        dgvInitiation.Rows[i].Cells[2].Value = "";
                        dgvInitiation.Rows[i].Cells[3].Value = "";
                        dgvInitiation.Rows[i].Cells[4].Value = "";
                        dgvInitiation.Rows[i].Cells[5].Value = "";
                    }
                    //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    if (initDocsCompleteDate[i] != "")
                    {
                        dgvInitiation.Rows[i].Cells[2].Value = initDocsCompleteDate[i];
                    }
                    else
                    {
                        dgvInitiation.Rows[i].Cells[2].Value = "";
                    }
                    

                    if (initDocsListStatus[i] == "UNDONE")
                    {
                        dgvInitiation.Rows[i].Cells[6].Style.BackColor = Color.Orange;
                        inprog++;
                    }
                    else if (initDocsListStatus[i] == "DONE")
                    {
                        initationProgressVal++;
                        ///////////////////////////AL DIE CODE OM TE CHECK OF IETS VOOR IETS ANDERS GEDOEN IS/////////////////////
                        k = i;
                        comp++;
                        dgvInitiation.Rows[i].Cells[6].Style.BackColor = Color.LimeGreen;
                        // pbarInitiation.Value = (int)initationProgressVal;
                        initationPercentage = ((initationProgressVal) / initiationDocuments.Count) * 100;
                        //lblInitiationProgress.Text = "Progress: " + Math.Round(initationPercentage, 2) + "%";

                        xValues1[0] = "Initiation";
                        yValues1[0] = initationPercentage;

                        yValues2[0] = 100 - initationPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        xValues1[0] = "Initiation";
                        yValues1[0] = initationPercentage;

                        yValues2[0] = 100;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);

                        uncomp++;
                        dgvInitiation.Rows[i].Cells[6].Style.BackColor = Color.Gray;
                    }


                }

                for (int j = 0; j < k; j++)
                {
                    if (initDocsListStatus[j] == "") //Check if the previous tasks are not done or in progress, because then they are behind schedule
                    {
                        //Increment the behind schedule tasks
                        behind++;
                        //Set all the tasks that are behind schedule to display red
                        dgvInitiation.Rows[j].Cells[6].Style.BackColor = Color.Red;
                    }
                }

                k = 0;

                if (inprog == initiationDocuments.Count)
                {
                    xValues1[0] = "Initiation";
                    yValues1[0] = initationPercentage;

                    yValues2[0] = 100;

                    chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                    chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                }

                lblInitNumTasks.Text = (initationPercentage / 100).ToString("p");

                chartInit.ChartAreas[0].BackColor = Color.Transparent;
                chartInit.Legends[0].BackColor = Color.Transparent;
                chartInit.Legends[0].BackColor = Color.Transparent;
                string[] xInit = { "Completed Tasks  " + comp, "Not started Tasks  " + uncomp, "In Progress Tasks " + inprog, "Behind Schedule Tasks " + behind };

                double[] yInit = { comp, uncomp, inprog, behind };

                chartInit.Series["Series1"].Points.DataBindXY(xInit, yInit);
                chartInit.Series["Series1"].ChartType = SeriesChartType.Doughnut;

                chartInit.Series["Series1"]["PieLabelStyle"] = "Disabled";
                chartInit.Legends[0].Enabled = true;

                chartInit.Text = "Test";

                chartInit.Series["Series1"].Points[0].Color = Color.LimeGreen;
                chartInit.Series["Series1"].Points[1].Color = Color.Gray;
                chartInit.Series["Series1"].Points[2].Color = Color.Orange;
                chartInit.Series["Series1"].Points[3].Color = Color.Red;
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

                ///////////////PLANNING PHASE/////////////////
                planningDocuments.Add("ProjectPlan");
                planningDocuments.Add("ResourcePlan");
                planningDocuments.Add("FinancialPlan");
                planningDocuments.Add("QualityPlan");
                planningDocuments.Add("RiskPlan");
                planningDocuments.Add("AcceptancePlan");
                planningDocuments.Add("CommunicationPlan");
                planningDocuments.Add("ProcurementPlan");
                planningDocuments.Add("StatementOfWork");
                planningDocuments.Add("RequestForInformation");
                planningDocuments.Add("SupplierContract");
                planningDocuments.Add("RequestForProposal");
                planningDocuments.Add("PhaseReviewPlanning");

                planningDocsListStatus.Add("ProjectPlan");

                for (int i = 0; i < planningDocuments.Count; i++)
                {
                    dgvPlanning.Rows.Add();
                    dgvPlanning.Rows[i].Cells[0].Value = planningDocuments[i];

                    if (planningDocsListDueDate.Count > 0)
                    {
                        if (planningDocsListDueDate[i].dueDate != null)
                            dgvPlanning.Rows[i].Cells[3].Value = planningDocsListDueDate[i].dueDate;
                        else
                            dgvPlanning.Rows[i].Cells[3].Value = "";

                        if (planningDocsListDueDate[i].budget != null)
                            dgvPlanning.Rows[i].Cells[5].Value = planningDocsListDueDate[i].budget;
                        else
                            dgvPlanning.Rows[i].Cells[5].Value = "";

                        if (planningDocsListDueDate[i].startDate != null)
                            dgvPlanning.Rows[i].Cells[1].Value = planningDocsListDueDate[i].startDate;
                        else
                            dgvPlanning.Rows[i].Cells[1].Value = "";

                        if (planningDocsListDueDate[i].completedDate != null)
                            dgvPlanning.Rows[i].Cells[2].Value = planningDocsListDueDate[i].completedDate;
                        else
                            dgvPlanning.Rows[i].Cells[2].Value = "";

                        if (planningDocsListDueDate[i].plannedBudget != null)
                            dgvPlanning.Rows[i].Cells[4].Value = planningDocsListDueDate[i].plannedBudget;
                        else
                            dgvPlanning.Rows[i].Cells[4].Value = "";
                    }
                    else
                    {
                        dgvPlanning.Rows[i].Cells[1].Value = "";
                        dgvPlanning.Rows[i].Cells[2].Value = "";
                        dgvPlanning.Rows[i].Cells[3].Value = "";
                        dgvPlanning.Rows[i].Cells[4].Value = "";
                        dgvPlanning.Rows[i].Cells[5].Value = "";
                    }

                    if (planningDocsCompleteDate[i] != "")
                    {
                        dgvPlanning.Rows[i].Cells[2].Value = planningDocsCompleteDate[i];
                    }
                    else
                    {
                        dgvPlanning.Rows[i].Cells[2].Value = "";
                    }


                    if (planningDocsListStatus[i] == "UNDONE")
                    {
                        dgvPlanning.Rows[i].Cells[6].Style.BackColor = Color.Orange;
                        inprogPlanning++;
                    }
                    else if (planningDocsListStatus[i] == "DONE")
                    {
                        planningProgressVal++;

                        k = i;

                        compPlanning++;
                        dgvPlanning.Rows[i].Cells[6].Style.BackColor = Color.LimeGreen;
                        planningPercentage = ((planningProgressVal) / planningDocuments.Count) * 100;

                        xValues1[1] = "Planning";
                        yValues1[1] = planningPercentage;

                        yValues2[1] = 100 - planningPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        xValues1[1] = "Planning";
                        yValues1[1] = planningPercentage;

                        yValues2[1] = 100;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);

                        uncompPlanning++;
                        dgvPlanning.Rows[i].Cells[6].Style.BackColor = Color.Gray;
                    }
                }

                for (int j = 0; j < k; j++)
                {
                    if (planningDocsListStatus[j] == "") //Check if the previous tasks are not done or in progress, because then they are behind schedule
                    {
                        //Increment the behind schedule tasks
                        behindPlanning++;
                        //Set all the tasks that are behind schedule to display red
                        dgvPlanning.Rows[j].Cells[6].Style.BackColor = Color.Red;
                    }
                }

                k = 0;

                if (inprogPlanning == planningDocuments.Count)
                {
                    xValues1[1] = "Planning";
                    yValues1[1] = planningPercentage;

                    yValues2[1] = 100;

                    chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                    chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                }

                lblPlanNumTasks.Text = (planningPercentage / 100).ToString("p");


                chartPlanning.ChartAreas[0].BackColor = Color.Transparent;
                chartPlanning.Legends[0].BackColor = Color.Transparent;
                chartPlanning.Legends[0].BackColor = Color.Transparent;
                string[] xPlan = { "Completed Tasks  " + compPlanning, "Not started Tasks  " + uncompPlanning, "In Progress Tasks " + inprogPlanning, "Behind Schedule Tasks " + behindPlanning };

                double[] yPlan = { compPlanning, uncompPlanning, inprogPlanning, behindPlanning };

                chartPlanning.Series["Series1"].Points.DataBindXY(xPlan, yPlan);
                chartPlanning.Series["Series1"].ChartType = SeriesChartType.Doughnut;

                chartPlanning.Series["Series1"]["PieLabelStyle"] = "Disabled";
                chartPlanning.Legends[0].Enabled = true;

                chartPlanning.Text = "Test";

                chartPlanning.Series["Series1"].Points[0].Color = Color.LimeGreen;
                chartPlanning.Series["Series1"].Points[1].Color = Color.Gray;
                chartPlanning.Series["Series1"].Points[2].Color = Color.Orange;
                chartPlanning.Series["Series1"].Points[3].Color = Color.Red;
                ///////////////////////////////////////////////////////////////////////////////////////////////////

                ///////////////EXECUTION PHASE/////////////////
                //executionDocuments.Add("BuildDeliverables"); Guides
                //executionDocuments.Add("MonitorAndControl"); Guides
                executionDocuments.Add("TimeMangement");
                executionDocuments.Add("TimeSheet");
                executionDocuments.Add("TimeSheetRegister");
                executionDocuments.Add("CostManagementProcess");
                executionDocuments.Add("ExpenseForm");
                executionDocuments.Add("ExpenseRegister");

                ////////////////////
                ///Sal jy dalk net ook kyk na die Quality goed, daar is 4 goed hier maar net 3 goed in die mainform onder quality (Nico)
                executionDocuments.Add("QualityManagement");
                executionDocuments.Add("QualityReviewPlan");
                executionDocuments.Add("QualityReviewForm");
                //executionDocuments.Add("QualityReviewRegister"); Die kan ons eers los dink ek, ek sien dit ook nie in die main nie (Rickus)
                ///////////////////


                executionDocuments.Add("ChangeManagementProcess");
                executionDocuments.Add("ChangeRequestForm");
                executionDocuments.Add("ChangeRequestRegister");
                executionDocuments.Add("RiskManagamentProcess");
                executionDocuments.Add("RiskForm");
                executionDocuments.Add("RiskRegister");
                executionDocuments.Add("IssueManagementProcess");
                executionDocuments.Add("IssueForm");
                executionDocuments.Add("IssueRegister");
                executionDocuments.Add("PurchaseOrder");
                executionDocuments.Add("ProcurementRegister");
                executionDocuments.Add("AcceptanceManagementProcess");
                executionDocuments.Add("AcceptanceForm");
                executionDocuments.Add("AcceptanceRegister");
                executionDocuments.Add("CommunicationsManagementProcess");
                executionDocuments.Add("ProjectStatusReport");
                executionDocuments.Add("CommunicationsRegister");
                executionDocuments.Add("PhaseReviewExe");

                executionDocsListStatus.Add("TimeMangement");

                for (int i = 0; i < executionDocuments.Count; i++)
                {

                    dgvExecution.Rows.Add();
                    dgvExecution.Rows[i].Cells[0].Value = executionDocuments[i];

                    if (executeDocsListDueDate.Count > 0)
                    {
                        if (executeDocsListDueDate[i].dueDate != null)
                            dgvExecution.Rows[i].Cells[3].Value = executeDocsListDueDate[i].dueDate;
                        else
                            dgvExecution.Rows[i].Cells[3].Value = "";

                        if (executeDocsListDueDate[i].budget != null)
                            dgvExecution.Rows[i].Cells[5].Value = executeDocsListDueDate[i].budget;
                        else
                            dgvExecution.Rows[i].Cells[5].Value = "";

                        if (executeDocsListDueDate[i].startDate != null)
                            dgvExecution.Rows[i].Cells[1].Value = executeDocsListDueDate[i].startDate;
                        else
                            dgvExecution.Rows[i].Cells[1].Value = "";

                        if (executeDocsListDueDate[i].completedDate != null)
                            dgvExecution.Rows[i].Cells[2].Value = executeDocsListDueDate[i].completedDate;
                        else
                            dgvExecution.Rows[i].Cells[2].Value = "";

                        if (executeDocsListDueDate[i].plannedBudget != null)
                            dgvExecution.Rows[i].Cells[4].Value = executeDocsListDueDate[i].plannedBudget;
                        else
                            dgvExecution.Rows[i].Cells[4].Value = "";
                    }
                    else
                    {
                        dgvExecution.Rows[i].Cells[1].Value = "";
                        dgvExecution.Rows[i].Cells[2].Value = "";
                        dgvExecution.Rows[i].Cells[3].Value = "";
                        dgvExecution.Rows[i].Cells[4].Value = "";
                        dgvExecution.Rows[i].Cells[5].Value = "";
                    }

                    if (executionDocsCompleteDate[i] != "")
                    {
                        dgvExecution.Rows[i].Cells[2].Value = executionDocsCompleteDate[i];
                    }
                    else
                    {
                        dgvExecution.Rows[i].Cells[2].Value = "";
                    }


                    if (executionDocsListStatus[i] == "UNDONE")
                    {
                        dgvExecution.Rows[i].Cells[6].Style.BackColor = Color.Orange;
                        inprogExecution++;
                    }
                    else if (executionDocsListStatus[i] == "DONE")
                    {
                        executionProgressVal++;

                        k = i;

                        compExecution++;
                        dgvExecution.Rows[i].Cells[6].Style.BackColor = Color.LimeGreen;
                        executionPercentage = ((executionProgressVal) / executionDocuments.Count) * 100;

                        xValues1[2] = "Execution";
                        yValues1[2] = executionPercentage;

                        yValues1[2] = 100;
                        yValues2[2] = 100 - executionPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        xValues1[2] = "Execution";
                        yValues1[2] = executionPercentage;

                        yValues2[2] = 100;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);

                        uncompExecution++;
                        dgvExecution.Rows[i].Cells[6].Style.BackColor = Color.Gray;
                    }
                }

                for (int j = 0; j < k; j++)
                {
                    if (executionDocsListStatus[j] == "") //Check if the previous tasks are not done or in progress, because then they are behind schedule
                    {
                        //Increment the behind schedule tasks
                        behindExecution++;
                        //Set all the tasks that are behind schedule to display red
                        dgvExecution.Rows[j].Cells[6].Style.BackColor = Color.Red;
                    }
                }

                k = 0;

                if (inprogExecution == executionDocuments.Count)
                {
                    xValues1[2] = "Execution";
                    yValues1[2] = executionPercentage;

                    yValues2[2] = 100;

                    chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                    chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                }

                lblExecNumTasks.Text = (executionPercentage / 100).ToString("p");

                chartExecution.ChartAreas[0].BackColor = Color.Transparent;
                chartExecution.Legends[0].BackColor = Color.Transparent;
                chartExecution.Legends[0].BackColor = Color.Transparent;
                string[] xExec = { "Completed Tasks  " + compExecution, "Not started Tasks  " + uncompExecution, "In Progress Tasks " + inprogExecution, "Behind Schedule Tasks " + behindExecution };

                double[] yExec = { compExecution, uncompExecution, inprogExecution, behindExecution };

                chartExecution.Series["Series1"].Points.DataBindXY(xExec, yExec);
                chartExecution.Series["Series1"].ChartType = SeriesChartType.Doughnut;

                chartExecution.Series["Series1"]["PieLabelStyle"] = "Disabled";
                chartExecution.Legends[0].Enabled = true;

                chartExecution.Text = "Test";

                chartExecution.Series["Series1"].Points[0].Color = Color.LimeGreen;
                chartExecution.Series["Series1"].Points[1].Color = Color.Gray;
                chartExecution.Series["Series1"].Points[2].Color = Color.Orange;
                chartExecution.Series["Series1"].Points[3].Color = Color.Red;
                ///////////////////////////////////////////////////////////////////////////////////////////////////

                ////////////////////////////////////////CLOSING PHASE/////////////////////////////////////////////
                closingDocuments.Add("ProjectClosureReport");
                closingDocuments.Add("PostImplementationReview");

                closingDocsListStatus.Add("ProjectClosureReport");

                for (int i = 0; i < closingDocuments.Count; i++)
                {
                    dgvClosing.Rows.Add();
                    dgvClosing.Rows[i].Cells[0].Value = closingDocuments[i];

                    if (closingListDueDate.Count > 0)
                    {
                        if (closingListDueDate[i].dueDate != null)
                            dgvClosing.Rows[i].Cells[1].Value = closingListDueDate[i].startDate;
                        else
                            dgvClosing.Rows[i].Cells[1].Value = "";

                        if (closingListDueDate[i].completedDate != null)
                            dgvClosing.Rows[i].Cells[2].Value = closingListDueDate[i].completedDate;
                        else
                            dgvClosing.Rows[i].Cells[2].Value = "";

                        if (closingListDueDate[i].dueDate != null)
                            dgvClosing.Rows[i].Cells[3].Value = closingListDueDate[i].dueDate;
                        else
                            dgvClosing.Rows[i].Cells[3].Value = "";

                        if (closingListDueDate[i].plannedBudget != null)
                            dgvClosing.Rows[i].Cells[4].Value = closingListDueDate[i].plannedBudget;
                        else
                            dgvClosing.Rows[i].Cells[4].Value = "";

                        if (closingListDueDate[i].budget != null)
                            dgvClosing.Rows[i].Cells[5].Value = closingListDueDate[i].budget;
                        else
                            dgvClosing.Rows[i].Cells[5].Value = "";
                    }
                    else
                    {
                        dgvClosing.Rows[i].Cells[1].Value = "";
                        dgvClosing.Rows[i].Cells[2].Value = "";
                        dgvClosing.Rows[i].Cells[3].Value = "";
                        dgvClosing.Rows[i].Cells[4].Value = "";
                        dgvClosing.Rows[i].Cells[5].Value = "";
                    }

                    if (closingDocsCompleteDate[i] != "")
                    {
                        dgvClosing.Rows[i].Cells[2].Value = closingDocsCompleteDate[i];
                    }
                    else
                    {
                        dgvClosing.Rows[i].Cells[2].Value = "";
                    }

                    if (closingDocsListStatus[i] == "UNDONE")
                    {
                        dgvClosing.Rows[i].Cells[6].Style.BackColor = Color.Orange;
                        inprogClosing++;
                    }
                    else if (closingDocsListStatus[i] == "DONE")
                    {
                        closingProgressVal++;

                        k = i;

                        compClosing++;
                        dgvClosing.Rows[i].Cells[6].Style.BackColor = Color.LimeGreen;
                        closingPercentage = ((closingProgressVal) / closingDocuments.Count) * 100;

                        xValues1[3] = "Closing";
                        yValues1[3] = closingPercentage;

                        yValues2[3] = 100 - closingPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        xValues1[3] = "Closing";
                        yValues1[3] = closingPercentage;

                        yValues2[3] = 100;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);

                        uncompClosing++;
                        dgvClosing.Rows[i].Cells[6].Style.BackColor = Color.Gray;
                    }
                }

                for (int j = 0; j < k; j++)
                {
                    if (closingDocsListStatus[j] == "") //Check if the previous tasks are not done or in progress, because then they are behind schedule
                    {
                        //Increment the behind schedule tasks
                        behindClosing++;
                        //Set all the tasks that are behind schedule to display red
                        dgvClosing.Rows[j].Cells[6].Style.BackColor = Color.Red;
                    }
                }

                k = 0;

                if (inprogClosing == closingDocuments.Count)
                {
                    xValues1[3] = "Closing";
                    yValues1[3] = closingPercentage;

                    yValues2[3] = 100;

                    chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                    chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                }

                lblClosingNumTasks.Text = (closingPercentage / 100).ToString("p");


                chartClosing.ChartAreas[0].BackColor = Color.Transparent;
                chartClosing.Legends[0].BackColor = Color.Transparent;
                chartClosing.Legends[0].BackColor = Color.Transparent;
                string[] xClose = { "Completed Tasks  " + compClosing, "Not started Tasks  " + uncompClosing, "In Progress Tasks " + inprogClosing, "Behind Schedule Tasks " + behindClosing };

                double[] yClose = { compClosing, uncompClosing, inprogClosing, behindClosing };

                chartClosing.Series["Series1"].Points.DataBindXY(xClose, yClose);
                chartClosing.Series["Series1"].ChartType = SeriesChartType.Doughnut;

                chartClosing.Series["Series1"]["PieLabelStyle"] = "Disabled";
                chartClosing.Legends[0].Enabled = true;

                chartClosing.Text = "Test";

                chartClosing.Series["Series1"].Points[0].Color = Color.LimeGreen;
                chartClosing.Series["Series1"].Points[1].Color = Color.Gray;
                chartClosing.Series["Series1"].Points[2].Color = Color.Orange;
                chartClosing.Series["Series1"].Points[3].Color = Color.Red;
                ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                double overallProgressVal = (initationProgressVal + planningProgressVal + executionProgressVal + closingProgressVal);
                pbarOverall.Value = (int)overallProgressVal;
                pbarOverall.Maximum = initiationDocuments.Count + planningDocuments.Count + executionDocuments.Count + closingDocuments.Count;
                double overallPercentage = ((overallProgressVal) / pbarOverall.Maximum) * 100;
                lblOverallProgress.Text = "Overall Progress: " + Math.Round(overallPercentage, 2) + "%";

                string[] xValues = { "Completed Tasks", "Not Started Tasks" };
                double[] yValues = { overallPercentage, 100 - overallPercentage };

                chart1.Series["Series1"].Points.DataBindXY(xValues, yValues);
                chart1.Series["Series1"].ChartType = SeriesChartType.Doughnut;

                chart1.Series["Series1"]["PieLabelStyle"] = "Disabled";
                chart1.Legends[0].Enabled = true;

                chart1.Series["Series1"].Points[0].Color = Color.LimeGreen;
                chart1.Series["Series1"].Points[1].Color = Color.Gray;


                foreach (DataPoint p in chart1.Series["Series1"].Points)
                {
                    p.Label = "#PERCENT\n#VALX";
                }
            }

            canChange = true;

            //Calling earnedValueAnalysis method to update the data grid view

            earnedValueAnalysis(dgvInitiation, daysSpent, daysAhead, daysBehind, budgetSpent, budgetAhead, budgetBehind, totalDaysInitlbl, lblTotalInitialBudget);
            earnedValueAnalysis(dgvClosing, daysSpent, daysAhead, daysBehind, budgetSpent, budgetAhead, budgetBehind, lblClosingDays, lblClosingBudget);
            earnedValueAnalysis(dgvPlanning, daysSpent, daysAhead, daysBehind, budgetSpent, budgetAhead, budgetBehind, lblPlanningSchedule, lblPlanningBudget);
            earnedValueAnalysis(dgvExecution, daysSpent, daysAhead, daysBehind, budgetSpent, budgetAhead, budgetBehind, lblExecutionSchedule, lblExecutionBudget);
        }

        private List<string> getLocalDocuments()
        {
            List<string> localDocuments = new List<string>();
            string projectPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ProjectManagementToolkit", Settings.Default.ProjectID);

            if (Directory.Exists(projectPath))
            {
                foreach (string documentPath in Directory.GetFiles(projectPath))
                {
                    string documentName = Path.GetFileNameWithoutExtension(documentPath);
                    localDocuments.Add(documentName);
                }
            }
            else
            {
                return null;
            }

            return localDocuments;
        }


        private void chart3_Click(object sender, EventArgs e)
        {

        }

        DateTimePicker InitDateTimePicker;
        DateTimePicker PlanDateTimePicker;
        DateTimePicker ExecuteDateTimePicker;
        DateTimePicker CloseDateTimePicker;

        InitDueDateModel currentInit = new InitDueDateModel();
        PlanningDueDateModel currentPlan = new PlanningDueDateModel();
        ExecutionDueDateModel currentExecute = new ExecutionDueDateModel();
        ClosingDueDateModel currentClose = new ClosingDueDateModel();

        //On Schedule, Behind Schedule, or Ahead of schedule
        int daysSpent = 0;
        int daysAhead = 0;
        int daysBehind = 0;

        //Over Spent, In Budget, On Budget
        double budgetSpent = 0;
        double budgetAhead = 0;
        double budgetBehind = 0;



        private void saveAllDueDate(int phase)
        {
            if (phase == 1)
            {
                //Start date
                currentInit.BusinessCaseSD = dgvInitiation.Rows[0].Cells[1].Value.ToString();
                currentInit.FeasibilityStudySD = dgvInitiation.Rows[1].Cells[1].Value.ToString();
                currentInit.ProjectCharterSD = dgvInitiation.Rows[2].Cells[1].Value.ToString();
                currentInit.JobDescriptionSD = dgvInitiation.Rows[3].Cells[1].Value.ToString();
                currentInit.ProjectOfficeCheckListSD = dgvInitiation.Rows[4].Cells[1].Value.ToString();
                currentInit.PhaseRevieFormInitiationSD = dgvInitiation.Rows[5].Cells[1].Value.ToString();
                currentInit.TermOfReferenceDocumentSD = dgvInitiation.Rows[6].Cells[1].Value.ToString();

                //Completed date
                currentInit.BusinessCaseCD = dgvInitiation.Rows[0].Cells[2].Value.ToString();
                currentInit.FeasibilityStudyCD = dgvInitiation.Rows[1].Cells[2].Value.ToString();
                currentInit.ProjectCharterCD = dgvInitiation.Rows[2].Cells[2].Value.ToString();
                currentInit.JobDescriptionCD = dgvInitiation.Rows[3].Cells[2].Value.ToString();
                currentInit.ProjectOfficeCheckListCD = dgvInitiation.Rows[4].Cells[2].Value.ToString();
                currentInit.PhaseRevieFormInitiationCD = dgvInitiation.Rows[5].Cells[2].Value.ToString();
                currentInit.TermOfReferenceDocumentCD = dgvInitiation.Rows[6].Cells[2].Value.ToString();

                //Due date
                currentInit.BusinessCaseDD = dgvInitiation.Rows[0].Cells[3].Value.ToString();
                currentInit.FeasibilityStudyDD = dgvInitiation.Rows[1].Cells[3].Value.ToString();
                currentInit.ProjectCharterDD = dgvInitiation.Rows[2].Cells[3].Value.ToString();
                currentInit.JobDescriptionDD = dgvInitiation.Rows[3].Cells[3].Value.ToString();
                currentInit.ProjectOfficeCheckListDD = dgvInitiation.Rows[4].Cells[3].Value.ToString();
                currentInit.PhaseRevieFormInitiationDD = dgvInitiation.Rows[5].Cells[3].Value.ToString();
                currentInit.TermOfReferenceDocumentDD = dgvInitiation.Rows[6].Cells[3].Value.ToString();

                //Planned Budget
                currentInit.BusinessCasePlannedBudget = dgvInitiation.Rows[0].Cells[4].Value.ToString();
                currentInit.FeasibilityStudyPlannedBudget = dgvInitiation.Rows[1].Cells[4].Value.ToString();
                currentInit.ProjectCharterPlannedBudget = dgvInitiation.Rows[2].Cells[4].Value.ToString();
                currentInit.JobDescriptionPlannedBudget = dgvInitiation.Rows[3].Cells[4].Value.ToString();
                currentInit.ProjectOfficeCheckListPlannedBudget = dgvInitiation.Rows[4].Cells[4].Value.ToString();
                currentInit.PhaseRevieFormInitiationPlannedBudget = dgvInitiation.Rows[5].Cells[4].Value.ToString();
                currentInit.TermOfReferenceDocumentPlannedBudget = dgvInitiation.Rows[6].Cells[4].Value.ToString();

                //Actual Budget Used
                currentInit.BusinessCaseBudget = dgvInitiation.Rows[0].Cells[5].Value.ToString();
                currentInit.FeasibilityStudyBudget = dgvInitiation.Rows[1].Cells[5].Value.ToString();
                currentInit.ProjectCharterBudget = dgvInitiation.Rows[2].Cells[5].Value.ToString();
                currentInit.JobDescriptionBudget = dgvInitiation.Rows[3].Cells[5].Value.ToString();
                currentInit.ProjectOfficeCheckListBudget = dgvInitiation.Rows[4].Cells[5].Value.ToString();
                currentInit.PhaseRevieFormInitiationBudget = dgvInitiation.Rows[5].Cells[5].Value.ToString();
                currentInit.TermOfReferenceDocumentBudget = dgvInitiation.Rows[6].Cells[5].Value.ToString();

                earnedValueAnalysis(dgvInitiation, daysSpent, daysAhead, daysBehind, budgetSpent, budgetAhead, budgetBehind, totalDaysInitlbl, lblTotalInitialBudget);

                string jsong = JsonConvert.SerializeObject(currentInit);
                JsonHelper.saveDocument(jsong, Settings.Default.ProjectID, "InitDueDateModel");
            }
            else if (phase == 2)
            {
                //Start Date
                currentPlan.ProjectPlanSD = dgvPlanning.Rows[0].Cells[1].Value.ToString();
                currentPlan.ResourcePlanSD = dgvPlanning.Rows[1].Cells[1].Value.ToString();
                currentPlan.FinancialPlanSD = dgvPlanning.Rows[2].Cells[1].Value.ToString();
                currentPlan.QualityPlanSD = dgvPlanning.Rows[3].Cells[1].Value.ToString();
                currentPlan.RiskPlanSD = dgvPlanning.Rows[4].Cells[1].Value.ToString();
                currentPlan.AcceptancePlanSD = dgvPlanning.Rows[5].Cells[1].Value.ToString();
                currentPlan.CommunicationPlanSD = dgvPlanning.Rows[6].Cells[1].Value.ToString();
                currentPlan.ProcurementPlanSD = dgvPlanning.Rows[7].Cells[1].Value.ToString();
                currentPlan.StatementOfWorkSD = dgvPlanning.Rows[8].Cells[1].Value.ToString();
                currentPlan.RequestForInformationSD = dgvPlanning.Rows[9].Cells[1].Value.ToString();
                currentPlan.SupplierContractSD = dgvPlanning.Rows[10].Cells[1].Value.ToString();
                currentPlan.RequestForProposalSD = dgvPlanning.Rows[11].Cells[1].Value.ToString();
                currentPlan.PhaseReviewPlanningSD = dgvPlanning.Rows[12].Cells[1].Value.ToString();

                //Completed Date
                currentPlan.ProjectPlanCD = dgvPlanning.Rows[0].Cells[2].Value.ToString();
                currentPlan.ResourcePlanCD = dgvPlanning.Rows[1].Cells[2].Value.ToString();
                currentPlan.FinancialPlanCD = dgvPlanning.Rows[2].Cells[2].Value.ToString();
                currentPlan.QualityPlanCD = dgvPlanning.Rows[3].Cells[2].Value.ToString();
                currentPlan.RiskPlanCD = dgvPlanning.Rows[4].Cells[2].Value.ToString();
                currentPlan.AcceptancePlanCD = dgvPlanning.Rows[5].Cells[2].Value.ToString();
                currentPlan.CommunicationPlanCD = dgvPlanning.Rows[6].Cells[2].Value.ToString();
                currentPlan.ProcurementPlanCD = dgvPlanning.Rows[7].Cells[2].Value.ToString();
                currentPlan.StatementOfWorkCD = dgvPlanning.Rows[8].Cells[2].Value.ToString();
                currentPlan.RequestForInformationCD = dgvPlanning.Rows[9].Cells[2].Value.ToString();
                currentPlan.SupplierContractCD = dgvPlanning.Rows[10].Cells[2].Value.ToString();
                currentPlan.RequestForProposalCD = dgvPlanning.Rows[11].Cells[2].Value.ToString();
                currentPlan.PhaseReviewPlanningCD = dgvPlanning.Rows[12].Cells[2].Value.ToString();

                //Due Date
                currentPlan.ProjectPlanDD = dgvPlanning.Rows[0].Cells[3].Value.ToString();
                currentPlan.ResourcePlanDD = dgvPlanning.Rows[1].Cells[3].Value.ToString();
                currentPlan.FinancialPlanDD = dgvPlanning.Rows[2].Cells[3].Value.ToString();
                currentPlan.QualityPlanDD = dgvPlanning.Rows[3].Cells[3].Value.ToString();
                currentPlan.RiskPlanDD = dgvPlanning.Rows[4].Cells[3].Value.ToString();
                currentPlan.AcceptancePlanDD = dgvPlanning.Rows[5].Cells[3].Value.ToString();
                currentPlan.CommunicationPlanDD = dgvPlanning.Rows[6].Cells[3].Value.ToString();
                currentPlan.ProcurementPlanDD = dgvPlanning.Rows[7].Cells[3].Value.ToString();
                currentPlan.StatementOfWorkDD = dgvPlanning.Rows[8].Cells[3].Value.ToString();
                currentPlan.RequestForInformationDD = dgvPlanning.Rows[9].Cells[3].Value.ToString();
                currentPlan.SupplierContractDD = dgvPlanning.Rows[10].Cells[3].Value.ToString();
                currentPlan.RequestForProposalDD = dgvPlanning.Rows[11].Cells[3].Value.ToString();
                currentPlan.PhaseReviewPlanningDD = dgvPlanning.Rows[12].Cells[3].Value.ToString();

                //Planned Budget
                currentPlan.ProjectPlanPlannedBudget = dgvPlanning.Rows[0].Cells[4].Value.ToString();
                currentPlan.ResourcePlanPlannedBudget = dgvPlanning.Rows[1].Cells[4].Value.ToString();
                currentPlan.FinancialPlanPlannedBudget = dgvPlanning.Rows[2].Cells[4].Value.ToString();
                currentPlan.QualityPlanPlannedBudget = dgvPlanning.Rows[3].Cells[4].Value.ToString();
                currentPlan.RiskPlanPlannedBudget = dgvPlanning.Rows[4].Cells[4].Value.ToString();
                currentPlan.AcceptancePlanPlannedBudget = dgvPlanning.Rows[5].Cells[4].Value.ToString();
                currentPlan.CommunicationPlanPlannedBudget = dgvPlanning.Rows[6].Cells[4].Value.ToString();
                currentPlan.ProcurementPlanPlannedBudget = dgvPlanning.Rows[7].Cells[4].Value.ToString();
                currentPlan.StatementOfWorkPlannedBudget = dgvPlanning.Rows[8].Cells[4].Value.ToString();
                currentPlan.RequestForInformationPlannedBudget = dgvPlanning.Rows[9].Cells[4].Value.ToString();
                currentPlan.SupplierContractPlannedBudget = dgvPlanning.Rows[10].Cells[4].Value.ToString();
                currentPlan.RequestForProposalPlannedBudget = dgvPlanning.Rows[11].Cells[4].Value.ToString();
                currentPlan.PhaseReviewPlanningPlannedBudget = dgvPlanning.Rows[12].Cells[4].Value.ToString();

                //Actual Budget Used
                currentPlan.ProjectPlanBudget = dgvPlanning.Rows[0].Cells[5].Value.ToString();
                currentPlan.ResourcePlanBudget = dgvPlanning.Rows[1].Cells[5].Value.ToString();
                currentPlan.FinancialPlanBudget = dgvPlanning.Rows[2].Cells[5].Value.ToString();
                currentPlan.QualityPlanBudget = dgvPlanning.Rows[3].Cells[5].Value.ToString();
                currentPlan.RiskPlanBudget = dgvPlanning.Rows[4].Cells[5].Value.ToString();
                currentPlan.AcceptancePlanBudget = dgvPlanning.Rows[5].Cells[5].Value.ToString();
                currentPlan.CommunicationPlanBudget = dgvPlanning.Rows[6].Cells[5].Value.ToString();
                currentPlan.ProcurementPlanBudget = dgvPlanning.Rows[7].Cells[5].Value.ToString();
                currentPlan.StatementOfWorkBudget = dgvPlanning.Rows[8].Cells[5].Value.ToString();
                currentPlan.RequestForInformationBudget = dgvPlanning.Rows[9].Cells[5].Value.ToString();
                currentPlan.SupplierContractBudget = dgvPlanning.Rows[10].Cells[5].Value.ToString();
                currentPlan.RequestForProposalBudget = dgvPlanning.Rows[11].Cells[5].Value.ToString();
                currentPlan.PhaseReviewPlanningBudget = dgvPlanning.Rows[12].Cells[5].Value.ToString();

                earnedValueAnalysis(dgvPlanning, daysSpent, daysAhead, daysBehind, budgetSpent, budgetAhead, budgetBehind, lblPlanningSchedule, lblPlanningBudget);

                string jsong = JsonConvert.SerializeObject(currentPlan);
                JsonHelper.saveDocument(jsong, Settings.Default.ProjectID, "PlanningDueDateModel");
            }
            else if (phase == 3)
            {
                //Start Date
                currentExecute.TimeMangementSD = dgvExecution.Rows[0].Cells[1].Value.ToString();
                currentExecute.TimeSheetSD = dgvExecution.Rows[1].Cells[1].Value.ToString();
                currentExecute.TimeSheetRegisterSD = dgvExecution.Rows[2].Cells[1].Value.ToString();
                currentExecute.CostManagementProcessSD = dgvExecution.Rows[3].Cells[1].Value.ToString();
                currentExecute.ExpenseFormSD = dgvExecution.Rows[4].Cells[1].Value.ToString();
                currentExecute.ExpenseRegisterSD = dgvExecution.Rows[5].Cells[1].Value.ToString();
                currentExecute.QualityManagementSD = dgvExecution.Rows[6].Cells[1].Value.ToString();
                currentExecute.QualityReviewPlanSD = dgvExecution.Rows[7].Cells[1].Value.ToString();
                currentExecute.QualityReviewFormSD = dgvExecution.Rows[8].Cells[1].Value.ToString();
                currentExecute.ChangeManagementProcessSD = dgvExecution.Rows[9].Cells[1].Value.ToString();
                currentExecute.ChangeRequestFormSD = dgvExecution.Rows[10].Cells[1].Value.ToString();
                currentExecute.ChangeRequestRegisterSD = dgvExecution.Rows[11].Cells[1].Value.ToString();
                currentExecute.RiskManagamentProcessSD = dgvExecution.Rows[12].Cells[1].Value.ToString();
                currentExecute.RiskFormSD = dgvExecution.Rows[13].Cells[1].Value.ToString();
                currentExecute.RiskRegisterSD = dgvExecution.Rows[14].Cells[1].Value.ToString();
                currentExecute.IssueManagementProcessSD = dgvExecution.Rows[15].Cells[1].Value.ToString();
                currentExecute.IssueFormSD = dgvExecution.Rows[16].Cells[1].Value.ToString();
                currentExecute.IssueRegisterSD = dgvExecution.Rows[17].Cells[1].Value.ToString();
                currentExecute.PurchaseOrderSD = dgvExecution.Rows[18].Cells[1].Value.ToString();
                currentExecute.ProcurementRegisterSD = dgvExecution.Rows[19].Cells[1].Value.ToString();
                currentExecute.AcceptanceManagementProcessSD = dgvExecution.Rows[20].Cells[1].Value.ToString();
                currentExecute.AcceptanceFormSD = dgvExecution.Rows[21].Cells[1].Value.ToString();
                currentExecute.AcceptanceRegisterSD = dgvExecution.Rows[22].Cells[1].Value.ToString();
                currentExecute.CommunicationsManagementProcessSD = dgvExecution.Rows[23].Cells[1].Value.ToString();
                currentExecute.ProjectStatusReportSD = dgvExecution.Rows[24].Cells[1].Value.ToString();
                currentExecute.CommunicationsRegisterSD = dgvExecution.Rows[25].Cells[1].Value.ToString();
                currentExecute.PhaseReviewExeSD = dgvExecution.Rows[26].Cells[1].Value.ToString();

                //Completed Date
                currentExecute.TimeMangementCD = dgvExecution.Rows[0].Cells[2].Value.ToString();
                currentExecute.TimeSheetCD = dgvExecution.Rows[1].Cells[2].Value.ToString();
                currentExecute.TimeSheetRegisterCD = dgvExecution.Rows[2].Cells[2].Value.ToString();
                currentExecute.CostManagementProcessCD = dgvExecution.Rows[3].Cells[2].Value.ToString();
                currentExecute.ExpenseFormCD = dgvExecution.Rows[4].Cells[2].Value.ToString();
                currentExecute.ExpenseRegisterCD = dgvExecution.Rows[5].Cells[2].Value.ToString();
                currentExecute.QualityManagementCD = dgvExecution.Rows[6].Cells[2].Value.ToString();
                currentExecute.QualityReviewPlanCD = dgvExecution.Rows[7].Cells[2].Value.ToString();
                currentExecute.QualityReviewFormCD = dgvExecution.Rows[8].Cells[2].Value.ToString();
                currentExecute.ChangeManagementProcessCD = dgvExecution.Rows[9].Cells[2].Value.ToString();
                currentExecute.ChangeRequestFormCD = dgvExecution.Rows[10].Cells[2].Value.ToString();
                currentExecute.ChangeRequestRegisterCD = dgvExecution.Rows[11].Cells[2].Value.ToString();
                currentExecute.RiskManagamentProcessCD = dgvExecution.Rows[12].Cells[2].Value.ToString();
                currentExecute.RiskFormCD = dgvExecution.Rows[13].Cells[2].Value.ToString();
                currentExecute.RiskRegisterCD = dgvExecution.Rows[14].Cells[2].Value.ToString();
                currentExecute.IssueManagementProcessCD = dgvExecution.Rows[15].Cells[2].Value.ToString();
                currentExecute.IssueFormCD = dgvExecution.Rows[16].Cells[2].Value.ToString();
                currentExecute.IssueRegisterCD = dgvExecution.Rows[17].Cells[2].Value.ToString();
                currentExecute.PurchaseOrderCD = dgvExecution.Rows[18].Cells[2].Value.ToString();
                currentExecute.ProcurementRegisterCD = dgvExecution.Rows[19].Cells[2].Value.ToString();
                currentExecute.AcceptanceManagementProcessCD = dgvExecution.Rows[20].Cells[2].Value.ToString();
                currentExecute.AcceptanceFormCD = dgvExecution.Rows[21].Cells[2].Value.ToString();
                currentExecute.AcceptanceRegisterCD = dgvExecution.Rows[22].Cells[2].Value.ToString();
                currentExecute.CommunicationsManagementProcessCD = dgvExecution.Rows[23].Cells[2].Value.ToString();
                currentExecute.ProjectStatusReportCD = dgvExecution.Rows[24].Cells[2].Value.ToString();
                currentExecute.CommunicationsRegisterCD = dgvExecution.Rows[25].Cells[2].Value.ToString();
                currentExecute.PhaseReviewExeCD = dgvExecution.Rows[26].Cells[2].Value.ToString();

                //Due Date
                currentExecute.TimeMangementDD = dgvExecution.Rows[0].Cells[3].Value.ToString();
                currentExecute.TimeSheetDD = dgvExecution.Rows[1].Cells[3].Value.ToString();
                currentExecute.TimeSheetRegisterDD = dgvExecution.Rows[2].Cells[3].Value.ToString();
                currentExecute.CostManagementProcessDD = dgvExecution.Rows[3].Cells[3].Value.ToString();
                currentExecute.ExpenseFormDD = dgvExecution.Rows[4].Cells[3].Value.ToString();
                currentExecute.ExpenseRegisterDD = dgvExecution.Rows[5].Cells[3].Value.ToString();
                currentExecute.QualityManagementDD = dgvExecution.Rows[6].Cells[3].Value.ToString();
                currentExecute.QualityReviewPlanDD = dgvExecution.Rows[7].Cells[3].Value.ToString();
                currentExecute.QualityReviewFormDD = dgvExecution.Rows[8].Cells[3].Value.ToString();
                currentExecute.ChangeManagementProcessDD = dgvExecution.Rows[9].Cells[3].Value.ToString();
                currentExecute.ChangeRequestFormDD = dgvExecution.Rows[10].Cells[3].Value.ToString();
                currentExecute.ChangeRequestRegisterDD = dgvExecution.Rows[11].Cells[3].Value.ToString();
                currentExecute.RiskManagamentProcessDD = dgvExecution.Rows[12].Cells[3].Value.ToString();
                currentExecute.RiskFormDD = dgvExecution.Rows[13].Cells[3].Value.ToString();
                currentExecute.RiskRegisterDD = dgvExecution.Rows[14].Cells[3].Value.ToString();
                currentExecute.IssueManagementProcessDD = dgvExecution.Rows[15].Cells[3].Value.ToString();
                currentExecute.IssueFormDD = dgvExecution.Rows[16].Cells[3].Value.ToString();
                currentExecute.IssueRegisterDD = dgvExecution.Rows[17].Cells[3].Value.ToString();
                currentExecute.PurchaseOrderDD = dgvExecution.Rows[18].Cells[3].Value.ToString();
                currentExecute.ProcurementRegisterDD = dgvExecution.Rows[19].Cells[3].Value.ToString();
                currentExecute.AcceptanceManagementProcessDD = dgvExecution.Rows[20].Cells[3].Value.ToString();
                currentExecute.AcceptanceFormDD = dgvExecution.Rows[21].Cells[3].Value.ToString();
                currentExecute.AcceptanceRegisterDD = dgvExecution.Rows[22].Cells[3].Value.ToString();
                currentExecute.CommunicationsManagementProcessDD = dgvExecution.Rows[23].Cells[3].Value.ToString();
                currentExecute.ProjectStatusReportDD = dgvExecution.Rows[24].Cells[3].Value.ToString();
                currentExecute.CommunicationsRegisterDD = dgvExecution.Rows[25].Cells[3].Value.ToString();
                currentExecute.PhaseReviewExeDD = dgvExecution.Rows[26].Cells[3].Value.ToString();

                //Planned Budget
                currentExecute.TimeMangementPlannedBudget = dgvExecution.Rows[0].Cells[4].Value.ToString();
                currentExecute.TimeSheetPlannedBudget = dgvExecution.Rows[1].Cells[4].Value.ToString();
                currentExecute.TimeSheetRegisterPlannedBudget = dgvExecution.Rows[2].Cells[4].Value.ToString();
                currentExecute.CostManagementProcessPlannedBudget = dgvExecution.Rows[3].Cells[4].Value.ToString();
                currentExecute.ExpenseFormPlannedBudget = dgvExecution.Rows[4].Cells[4].Value.ToString();
                currentExecute.ExpenseRegisterPlannedBudget = dgvExecution.Rows[5].Cells[4].Value.ToString();
                currentExecute.QualityManagementPlannedBudget = dgvExecution.Rows[6].Cells[4].Value.ToString();
                currentExecute.QualityReviewPlanPlannedBudget = dgvExecution.Rows[7].Cells[4].Value.ToString();
                currentExecute.QualityReviewFormPlannedBudget = dgvExecution.Rows[8].Cells[4].Value.ToString();
                currentExecute.ChangeManagementProcessPlannedBudget = dgvExecution.Rows[9].Cells[4].Value.ToString();
                currentExecute.ChangeRequestFormPlannedBudget = dgvExecution.Rows[10].Cells[4].Value.ToString();
                currentExecute.ChangeRequestRegisterPlannedBudget = dgvExecution.Rows[11].Cells[4].Value.ToString();
                currentExecute.RiskManagamentProcessPlannedBudget = dgvExecution.Rows[12].Cells[4].Value.ToString();
                currentExecute.RiskFormPlannedBudget = dgvExecution.Rows[13].Cells[4].Value.ToString();
                currentExecute.RiskRegisterPlannedBudget = dgvExecution.Rows[14].Cells[4].Value.ToString();
                currentExecute.IssueManagementProcessPlannedBudget = dgvExecution.Rows[15].Cells[4].Value.ToString();
                currentExecute.IssueFormPlannedBudget = dgvExecution.Rows[16].Cells[4].Value.ToString();
                currentExecute.IssueRegisterPlannedBudget = dgvExecution.Rows[17].Cells[4].Value.ToString();
                currentExecute.PurchaseOrderPlannedBudget = dgvExecution.Rows[18].Cells[4].Value.ToString();
                currentExecute.ProcurementRegisterPlannedBudget = dgvExecution.Rows[19].Cells[4].Value.ToString();
                currentExecute.AcceptanceManagementProcessPlannedBudget = dgvExecution.Rows[20].Cells[4].Value.ToString();
                currentExecute.AcceptanceFormPlannedBudget = dgvExecution.Rows[21].Cells[4].Value.ToString();
                currentExecute.AcceptanceRegisterPlannedBudget = dgvExecution.Rows[22].Cells[4].Value.ToString();
                currentExecute.CommunicationsManagementProcessPlannedBudget = dgvExecution.Rows[23].Cells[4].Value.ToString();
                currentExecute.ProjectStatusReportPlannedBudget = dgvExecution.Rows[24].Cells[4].Value.ToString();
                currentExecute.CommunicationsRegisterPlannedBudget = dgvExecution.Rows[25].Cells[4].Value.ToString();
                currentExecute.PhaseReviewExePlannedBudget = dgvExecution.Rows[26].Cells[4].Value.ToString();

                //Actual Budget Used
                currentExecute.TimeMangementBudget = dgvExecution.Rows[0].Cells[5].Value.ToString();
                currentExecute.TimeSheetBudget = dgvExecution.Rows[1].Cells[5].Value.ToString();
                currentExecute.TimeSheetRegisterBudget = dgvExecution.Rows[2].Cells[5].Value.ToString();
                currentExecute.CostManagementProcessBudget = dgvExecution.Rows[3].Cells[5].Value.ToString();
                currentExecute.ExpenseFormBudget = dgvExecution.Rows[4].Cells[5].Value.ToString();
                currentExecute.ExpenseRegisterBudget = dgvExecution.Rows[5].Cells[5].Value.ToString();
                currentExecute.QualityManagementBudget = dgvExecution.Rows[6].Cells[5].Value.ToString();
                currentExecute.QualityReviewPlanBudget = dgvExecution.Rows[7].Cells[5].Value.ToString();
                currentExecute.QualityReviewFormBudget = dgvExecution.Rows[8].Cells[5].Value.ToString();
                currentExecute.ChangeManagementProcessBudget = dgvExecution.Rows[9].Cells[5].Value.ToString();
                currentExecute.ChangeRequestFormBudget = dgvExecution.Rows[10].Cells[5].Value.ToString();
                currentExecute.ChangeRequestRegisterBudget = dgvExecution.Rows[11].Cells[5].Value.ToString();
                currentExecute.RiskManagamentProcessBudget = dgvExecution.Rows[12].Cells[5].Value.ToString();
                currentExecute.RiskFormBudget = dgvExecution.Rows[13].Cells[5].Value.ToString();
                currentExecute.RiskRegisterBudget = dgvExecution.Rows[14].Cells[5].Value.ToString();
                currentExecute.IssueManagementProcessBudget = dgvExecution.Rows[15].Cells[5].Value.ToString();
                currentExecute.IssueFormBudget = dgvExecution.Rows[16].Cells[5].Value.ToString();
                currentExecute.IssueRegisterBudget = dgvExecution.Rows[17].Cells[5].Value.ToString();
                currentExecute.PurchaseOrderBudget = dgvExecution.Rows[18].Cells[5].Value.ToString();
                currentExecute.ProcurementRegisterBudget = dgvExecution.Rows[19].Cells[5].Value.ToString();
                currentExecute.AcceptanceManagementProcessBudget = dgvExecution.Rows[20].Cells[5].Value.ToString();
                currentExecute.AcceptanceFormBudget = dgvExecution.Rows[21].Cells[5].Value.ToString();
                currentExecute.AcceptanceRegisterBudget = dgvExecution.Rows[22].Cells[5].Value.ToString();
                currentExecute.CommunicationsManagementProcessBudget = dgvExecution.Rows[23].Cells[5].Value.ToString();
                currentExecute.ProjectStatusReportBudget = dgvExecution.Rows[24].Cells[5].Value.ToString();
                currentExecute.CommunicationsRegisterBudget = dgvExecution.Rows[25].Cells[5].Value.ToString();
                currentExecute.PhaseReviewExeBudget = dgvExecution.Rows[26].Cells[5].Value.ToString();

                earnedValueAnalysis(dgvExecution, daysSpent, daysAhead, daysBehind, budgetSpent, budgetAhead, budgetBehind, lblExecutionSchedule, lblExecutionBudget);

                string jsong = JsonConvert.SerializeObject(currentExecute);
                JsonHelper.saveDocument(jsong, Settings.Default.ProjectID, "ExecutionDueDateModel");
            }
            else if (phase == 4)
            {

                //Start Date
                currentClose.ProjectClosureReportSD = dgvClosing.Rows[0].Cells[1].Value.ToString();
                currentClose.PostImplementationReviewSD = dgvClosing.Rows[1].Cells[1].Value.ToString();

                //Completed Date
                currentClose.ProjectClosureReportCD = dgvClosing.Rows[0].Cells[2].Value.ToString();
                currentClose.PostImplementationReviewCD = dgvClosing.Rows[1].Cells[2].Value.ToString();

                //Due Date
                currentClose.ProjectClosureReportDD = dgvClosing.Rows[0].Cells[3].Value.ToString();
                currentClose.PostImplementationReviewDD = dgvClosing.Rows[1].Cells[3].Value.ToString();

                //Planned Budget
                currentClose.ProjectClosureReportPlannedBudget = dgvClosing.Rows[0].Cells[4].Value.ToString();
                currentClose.PostImplementationReviewPlannedBudget = dgvClosing.Rows[1].Cells[4].Value.ToString();

                //Used Budget
                currentClose.ProjectClosureReportBudget = dgvClosing.Rows[0].Cells[5].Value.ToString();
                currentClose.PostImplementationReviewBudget = dgvClosing.Rows[1].Cells[5].Value.ToString();

                earnedValueAnalysis(dgvClosing, daysSpent, daysAhead, daysBehind, budgetSpent, budgetAhead, budgetBehind, lblClosingDays, lblClosingBudget);

                string jsong = JsonConvert.SerializeObject(currentClose);
                JsonHelper.saveDocument(jsong, Settings.Default.ProjectID, "ClosingDueDateModel");
            }

        }


        private void dgvInitiation_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            for(int i = 1; i < 4; i += 2)
            {
                // If any cell is clicked on the Second column which is our date Column  
                if (e.ColumnIndex == i && e.RowIndex != -1)
                {
                    //Initialized a new DateTimePicker Control  
                    InitDateTimePicker = new DateTimePicker();

                    //Adding DateTimePicker control into DataGridView   
                    dgvInitiation.Controls.Add(InitDateTimePicker);

                    // Setting the format (i.e. 2014-10-10)  
                    InitDateTimePicker.Format = DateTimePickerFormat.Short;

                    // It returns the retangular area that represents the Display area for a cell  
                    Rectangle oRectangle = dgvInitiation.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);

                    //Setting area for DateTimePicker Control  
                    InitDateTimePicker.Size = new Size(oRectangle.Width, oRectangle.Height);

                    // Setting Location  
                    InitDateTimePicker.Location = new Point(oRectangle.X, oRectangle.Y);

                    // An event attached to dateTimePicker Control which is fired when DateTimeControl is closed  
                    InitDateTimePicker.CloseUp += new EventHandler(InitDateTimePicker_CloseUp);

                    // An event attached to dateTimePicker Control which is fired when any date is selected  
                    InitDateTimePicker.TextChanged += new EventHandler(InitDateTimePicker_OnTextChange);

                    // Now make it visible  
                    InitDateTimePicker.Visible = true;
                }
            }
        }

        private void InitDateTimePicker_OnTextChange(object sender, EventArgs e)
        {
            // Initially set the date so that we can do the comparison
            dgvInitiation.CurrentCell.Value = InitDateTimePicker.Text.ToString();
            // This indicator will be used to determing whether the initiate date comes before or after the due date.
            // If indicator is < 0, then we know that the initiate date is before the due date, if it is > 0 then the initiate date is after the due date.
            int indicator = 0;
            // Getting the same row as the Initiate date cell we clicked on.
            int row = dgvInitiation.CurrentCell.RowIndex;
            if (dgvInitiation.CurrentCell.ColumnIndex == 1)
            {
                // Checking if there is a due date selected within the specific initiate date's row
                if (dgvInitiation.Rows[row].Cells[3].Value.ToString() != "")
                {
                    // Doing the indicator calculation
                    indicator = (Convert.ToDateTime(dgvInitiation.CurrentCell.Value) - Convert.ToDateTime(dgvInitiation.Rows[row].Cells[3].Value)).Days;
                    // If the initiate date is before the due date
                    if (indicator < 0)
                    {
                        dgvInitiation.CurrentCell.Value = InitDateTimePicker.Text.ToString();
                    }
                    // If the initiate date is after the due date we throw a message box prompting the user to select a date that comes before the due date
                    else if (indicator > 0)
                    {
                        MessageBox.Show("Please select a date before the specified due date.");
                        dgvInitiation.CurrentCell.Value = "";
                    }
                    else
                    {
                        dgvInitiation.CurrentCell.Value = InitDateTimePicker.Text.ToString();
                    }
                }
                else
                {
                    dgvInitiation.CurrentCell.Value = InitDateTimePicker.Text.ToString();
                }
            }
            else if (dgvInitiation.CurrentCell.ColumnIndex == 3)
            {
                // Checking if there is a due date selected within the specific initiate date's row
                if (dgvInitiation.Rows[row].Cells[1].Value.ToString() != "")
                {
                    // Doing the indicator calculation
                    indicator = (Convert.ToDateTime(dgvInitiation.CurrentCell.Value) - Convert.ToDateTime(dgvInitiation.Rows[row].Cells[1].Value)).Days;
                    // If the initiate date is before the due date
                    if (indicator > 0)
                    {
                        dgvInitiation.CurrentCell.Value = InitDateTimePicker.Text.ToString();
                    }
                    // If the initiate date is after the due date we throw a message box prompting the user to select a date that comes before the due date
                    else if (indicator < 0)
                    {
                        MessageBox.Show("Please select a date after the specified initiation date.");
                        dgvInitiation.CurrentCell.Value = "";
                    }
                    else
                    {
                        dgvInitiation.CurrentCell.Value = InitDateTimePicker.Text.ToString();
                    }
                }
                else
                {
                    dgvInitiation.CurrentCell.Value = InitDateTimePicker.Text.ToString();
                }
            }
            saveAllDueDate(1);
        }

        void InitDateTimePicker_CloseUp(object sender, EventArgs e)
        {
            // Hiding the control after use   
            InitDateTimePicker.Visible = false;
        }

        private void dgvPlanning_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            for(int i = 1; i < 4; i += 2)
            {
                if (e.ColumnIndex == i && e.RowIndex != -1)
                {
                    //Initialized a new DateTimePicker Control  
                    PlanDateTimePicker = new DateTimePicker();

                    //Adding DateTimePicker control into DataGridView   
                    dgvPlanning.Controls.Add(PlanDateTimePicker);

                    // Setting the format (i.e. 2014-10-10)  
                    PlanDateTimePicker.Format = DateTimePickerFormat.Short;

                    // It returns the retangular area that represents the Display area for a cell  
                    Rectangle oRectangle = dgvPlanning.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);

                    //Setting area for DateTimePicker Control  
                    PlanDateTimePicker.Size = new Size(oRectangle.Width, oRectangle.Height);

                    // Setting Location  
                    PlanDateTimePicker.Location = new Point(oRectangle.X, oRectangle.Y);

                    // An event attached to dateTimePicker Control which is fired when DateTimeControl is closed  
                    PlanDateTimePicker.CloseUp += new EventHandler(PlanDateTimePicker_CloseUp);

                    // An event attached to dateTimePicker Control which is fired when any date is selected  
                    PlanDateTimePicker.TextChanged += new EventHandler(PlanDateTimePicker_OnTextChange);

                    // Now make it visible  
                    PlanDateTimePicker.Visible = true;
                }
            }
        }

        private void PlanDateTimePicker_OnTextChange(object sender, EventArgs e)
        {
            // Initially set the date so that we can do the comparison
            dgvPlanning.CurrentCell.Value = PlanDateTimePicker.Text.ToString();
            // This indicator will be used to determing whether the initiate date comes before or after the due date.
            // If indicator is < 0, then we know that the initiate date is before the due date, if it is > 0 then the initiate date is after the due date.
            int indicator = 0;
            // Getting the same row as the Initiate date cell we clicked on.
            int row = dgvPlanning.CurrentCell.RowIndex;
            if (dgvPlanning.CurrentCell.ColumnIndex == 1)
            {
                // Checking if there is a due date selected within the specific initiate date's row
                if (dgvPlanning.Rows[row].Cells[3].Value.ToString() != "")
                {
                    // Doing the indicator calculation
                    indicator = (Convert.ToDateTime(dgvPlanning.CurrentCell.Value) - Convert.ToDateTime(dgvPlanning.Rows[row].Cells[3].Value)).Days;
                    // If the initiate date is before the due date
                    if (indicator < 0)
                    {
                        dgvPlanning.CurrentCell.Value = PlanDateTimePicker.Text.ToString();
                    }
                    // If the initiate date is after the due date we throw a message box prompting the user to select a date that comes before the due date
                    else if (indicator > 0)
                    {
                        MessageBox.Show("Please select a date before the specified due date.");
                        dgvPlanning.CurrentCell.Value = "";
                    }
                    else
                    {
                        dgvPlanning.CurrentCell.Value = PlanDateTimePicker.Text.ToString();
                    }
                }
                else
                {
                    dgvPlanning.CurrentCell.Value = PlanDateTimePicker.Text.ToString();
                }
            }
            else if (dgvPlanning.CurrentCell.ColumnIndex == 3)
            {
                // Checking if there is a due date selected within the specific initiate date's row
                if (dgvPlanning.Rows[row].Cells[1].Value.ToString() != "")
                {
                    // Doing the indicator calculation
                    indicator = (Convert.ToDateTime(dgvPlanning.CurrentCell.Value) - Convert.ToDateTime(dgvPlanning.Rows[row].Cells[1].Value)).Days;
                    // If the initiate date is before the due date
                    if (indicator > 0)
                    {
                        dgvPlanning.CurrentCell.Value = PlanDateTimePicker.Text.ToString();
                    }
                    // If the initiate date is after the due date we throw a message box prompting the user to select a date that comes before the due date
                    else if (indicator < 0)
                    {
                        MessageBox.Show("Please select a date after the specified initiation date.");
                        dgvPlanning.CurrentCell.Value = "";
                    }
                    else
                    {
                        dgvPlanning.CurrentCell.Value = PlanDateTimePicker.Text.ToString();
                    }
                }
                else
                {
                    dgvPlanning.CurrentCell.Value = PlanDateTimePicker.Text.ToString();
                }
            }

            saveAllDueDate(2);
        }


        void PlanDateTimePicker_CloseUp(object sender, EventArgs e)
        {
            // Hiding the control after use   
            PlanDateTimePicker.Visible = false;
        }

        private void dgvExecution_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            for(int i = 1; i < 4; i += 2)
            {
                if (e.ColumnIndex == i && e.RowIndex != -1)
                {
                    //Initialized a new DateTimePicker Control  
                    ExecuteDateTimePicker = new DateTimePicker();

                    //Adding DateTimePicker control into DataGridView   
                    dgvExecution.Controls.Add(ExecuteDateTimePicker);

                    // Setting the format (i.e. 2014-10-10)  
                    ExecuteDateTimePicker.Format = DateTimePickerFormat.Short;

                    // It returns the retangular area that represents the Display area for a cell  
                    Rectangle oRectangle = dgvExecution.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);

                    //Setting area for DateTimePicker Control  
                    ExecuteDateTimePicker.Size = new Size(oRectangle.Width, oRectangle.Height);

                    // Setting Location  
                    ExecuteDateTimePicker.Location = new Point(oRectangle.X, oRectangle.Y);

                    // An event attached to dateTimePicker Control which is fired when DateTimeControl is closed  
                    ExecuteDateTimePicker.CloseUp += new EventHandler(ExecuteDateTimePicker_CloseUp);

                    // An event attached to dateTimePicker Control which is fired when any date is selected  
                    ExecuteDateTimePicker.TextChanged += new EventHandler(ExecuteDateTimePicker_OnTextChange);

                    // Now make it visible  
                    ExecuteDateTimePicker.Visible = true;
                }
            }
        }


        private void ExecuteDateTimePicker_OnTextChange(object sender, EventArgs e)
        {
            // Initially set the date so that we can do the comparison
            dgvExecution.CurrentCell.Value = ExecuteDateTimePicker.Text.ToString();
            // This indicator will be used to determing whether the initiate date comes before or after the due date.
            // If indicator is < 0, then we know that the initiate date is before the due date, if it is > 0 then the initiate date is after the due date.
            int indicator = 0;
            // Getting the same row as the Initiate date cell we clicked on.
            int row = dgvExecution.CurrentCell.RowIndex;
            if (dgvExecution.CurrentCell.ColumnIndex == 1)
            {
                // Checking if there is a due date selected within the specific initiate date's row
                if (dgvExecution.Rows[row].Cells[3].Value.ToString() != "")
                {
                    // Doing the indicator calculation
                    indicator = (Convert.ToDateTime(dgvExecution.CurrentCell.Value) - Convert.ToDateTime(dgvExecution.Rows[row].Cells[3].Value)).Days;
                    // If the initiate date is before the due date
                    if (indicator < 0)
                    {
                        dgvExecution.CurrentCell.Value = ExecuteDateTimePicker.Text.ToString();
                    }
                    // If the initiate date is after the due date we throw a message box prompting the user to select a date that comes before the due date
                    else if (indicator > 0)
                    {
                        MessageBox.Show("Please select a date before the specified due date.");
                        dgvExecution.CurrentCell.Value = "";
                    }
                    else
                    {
                        dgvExecution.CurrentCell.Value = ExecuteDateTimePicker.Text.ToString();
                    }
                }
                else
                {
                    dgvExecution.CurrentCell.Value = ExecuteDateTimePicker.Text.ToString();
                }
            }
            else if (dgvExecution.CurrentCell.ColumnIndex == 3)
            {
                // Checking if there is a due date selected within the specific initiate date's row
                if (dgvExecution.Rows[row].Cells[1].Value.ToString() != "")
                {
                    // Doing the indicator calculation
                    indicator = (Convert.ToDateTime(dgvExecution.CurrentCell.Value) - Convert.ToDateTime(dgvExecution.Rows[row].Cells[1].Value)).Days;
                    // If the initiate date is before the due date
                    if (indicator > 0)
                    {
                        dgvExecution.CurrentCell.Value = ExecuteDateTimePicker.Text.ToString();
                    }
                    // If the initiate date is after the due date we throw a message box prompting the user to select a date that comes before the due date
                    else if (indicator < 0)
                    {
                        MessageBox.Show("Please select a date after the specified initiation date.");
                        dgvExecution.CurrentCell.Value = "";
                    }
                    else
                    {
                        dgvExecution.CurrentCell.Value = ExecuteDateTimePicker.Text.ToString();
                    }
                }
                else
                {
                    dgvExecution.CurrentCell.Value = ExecuteDateTimePicker.Text.ToString();
                }
            }

            saveAllDueDate(3);
        }


        void ExecuteDateTimePicker_CloseUp(object sender, EventArgs e)
        {
            // Hiding the control after use   
            ExecuteDateTimePicker.Visible = false;
        }

        

        private void CloseDateTimePicker_OnTextChange(object sender, EventArgs e)
        {
            // Initially set the date so that we can do the comparison
            dgvClosing.CurrentCell.Value = CloseDateTimePicker.Text.ToString();
            // This indicator will be used to determing whether the initiate date comes before or after the due date.
            // If indicator is < 0, then we know that the initiate date is before the due date, if it is > 0 then the initiate date is after the due date.
            int indicator = 0;
            // Getting the same row as the Initiate date cell we clicked on.
            int row = dgvClosing.CurrentCell.RowIndex;
            if (dgvClosing.CurrentCell.ColumnIndex == 1)
            {
                // Checking if there is a due date selected within the specific initiate date's row
                if (dgvClosing.Rows[row].Cells[3].Value.ToString() != "")
                {
                    // Doing the indicator calculation
                    indicator = (Convert.ToDateTime(dgvClosing.CurrentCell.Value) - Convert.ToDateTime(dgvClosing.Rows[row].Cells[3].Value)).Days;
                    // If the initiate date is before the due date
                    if (indicator < 0)
                    {
                        dgvClosing.CurrentCell.Value = CloseDateTimePicker.Text.ToString();
                    }
                    // If the initiate date is after the due date we throw a message box prompting the user to select a date that comes before the due date
                    else if (indicator > 0)
                    {
                        MessageBox.Show("Please select a date before the specified due date.");
                        dgvClosing.CurrentCell.Value = "";
                    }
                    else
                    {
                        dgvClosing.CurrentCell.Value = CloseDateTimePicker.Text.ToString();
                    }
                }
                else
                {
                    dgvClosing.CurrentCell.Value = CloseDateTimePicker.Text.ToString();
                }
            }
            else if (dgvClosing.CurrentCell.ColumnIndex == 3)
            {
                // Checking if there is a due date selected within the specific initiate date's row
                if (dgvClosing.Rows[row].Cells[1].Value.ToString() != "")
                {
                    // Doing the indicator calculation
                    indicator = (Convert.ToDateTime(dgvClosing.CurrentCell.Value) - Convert.ToDateTime(dgvClosing.Rows[row].Cells[1].Value)).Days;
                    // If the initiate date is before the due date
                    if (indicator > 0)
                    {
                        dgvClosing.CurrentCell.Value = CloseDateTimePicker.Text.ToString();
                    }
                    // If the initiate date is after the due date we throw a message box prompting the user to select a date that comes before the due date
                    else if (indicator < 0)
                    {
                        MessageBox.Show("Please select a date after the specified initiation date.");
                        dgvClosing.CurrentCell.Value = "";
                    }
                    else
                    {
                        dgvClosing.CurrentCell.Value = CloseDateTimePicker.Text.ToString();
                    }
                }
                else
                {
                    dgvClosing.CurrentCell.Value = CloseDateTimePicker.Text.ToString();
                }
            }

            saveAllDueDate(4);
        }


        void CloseDateTimePicker_CloseUp(object sender, EventArgs e)
        {
            // Hiding the control after use   
            CloseDateTimePicker.Visible = false;
        }

        private void dgvClosing_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            for(int i = 1; i < 4; i += 2)
            {
                if (e.ColumnIndex == i && e.RowIndex != -1)
                {
                    //Initialized a new DateTimePicker Control  
                    CloseDateTimePicker = new DateTimePicker();

                    //Adding DateTimePicker control into DataGridView   
                    dgvClosing.Controls.Add(CloseDateTimePicker);

                    // Setting the format (i.e. 2014-10-10)  
                    CloseDateTimePicker.Format = DateTimePickerFormat.Short;

                    // It returns the retangular area that represents the Display area for a cell  
                    Rectangle oRectangle = dgvClosing.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);

                    //Setting area for DateTimePicker Control  
                    CloseDateTimePicker.Size = new Size(oRectangle.Width, oRectangle.Height);

                    // Setting Location  
                    CloseDateTimePicker.Location = new Point(oRectangle.X, oRectangle.Y);

                    // An event attached to dateTimePicker Control which is fired when DateTimeControl is closed  
                    CloseDateTimePicker.CloseUp += new EventHandler(CloseDateTimePicker_CloseUp);

                    // An event attached to dateTimePicker Control which is fired when any date is selected  
                    CloseDateTimePicker.TextChanged += new EventHandler(CloseDateTimePicker_OnTextChange);

                    // Now make it visible  
                    CloseDateTimePicker.Visible = true;
                }
            }
        }

        bool canChange = false;
        private void dgvInitiation_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if(canChange)
                saveAllDueDate(1);
        }

        private void dgvPlanning_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (canChange)
                saveAllDueDate(2);
        }

        private void dgvExecution_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (canChange)
                saveAllDueDate(3);
        }

        private void dgvClosing_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (canChange)
                saveAllDueDate(4);
        }

        private bool validateDate(DateTime date1, DateTime date2)
        {
            if (date1 > date2)
                return false;
            else
                return true;
        }

        // METHOD


        public void earnedValueAnalysis(DataGridView dgvInitiation, int daysSpent, int daysAhead, int daysBehind, double budgetSpent, double budgetAhead, double budgetBehind, Label totalDaysInitlbl, Label lblTotalInitialBudget)
        {
            daysSpent = 0;
            budgetSpent = 0;

            //Logic for the intitiantion phase
            for (int i = 0; i < dgvInitiation.RowCount; i++)
            {
                daysAhead = 0;
                daysBehind = 0;

                budgetBehind = 0;
                budgetAhead = 0;

                if (dgvInitiation.Rows[i].Cells[4].Value.ToString() != "" && dgvInitiation.Rows[i].Cells[5].Value.ToString() != "")
                {
                    // Under Budget
                    if (Convert.ToDouble(dgvInitiation.Rows[i].Cells[4].Value.ToString()) > Convert.ToDouble(dgvInitiation.Rows[i].Cells[5].Value.ToString()))
                    {
                        budgetAhead = Convert.ToDouble(dgvInitiation.Rows[i].Cells[4].Value) - Convert.ToDouble(dgvInitiation.Rows[i].Cells[5].Value);
                        dgvInitiation.Rows[i].Cells[5].Style.ForeColor = Color.LimeGreen;
                    }
                    // Over Budget
                    else if (Convert.ToDouble(dgvInitiation.Rows[i].Cells[4].Value.ToString()) < Convert.ToDouble(dgvInitiation.Rows[i].Cells[5].Value.ToString()))
                    {
                        budgetBehind = Convert.ToDouble(dgvInitiation.Rows[i].Cells[4].Value) - Convert.ToDouble(dgvInitiation.Rows[i].Cells[5].Value);
                        dgvInitiation.Rows[i].Cells[5].Style.ForeColor = Color.Red;
                    }
                    else
                    {
                        dgvInitiation.Rows[i].Cells[5].Style.ForeColor = Color.Black;
                    }
                }
                else
                {
                    continue;
                }
                if (dgvInitiation.Rows[i].Cells[2].Value.ToString() != "" && dgvInitiation.Rows[i].Cells[3].Value.ToString() != "")
                {
                    // Behind Schedule
                    if (Convert.ToDateTime(dgvInitiation.Rows[i].Cells[2].Value) > Convert.ToDateTime(dgvInitiation.Rows[i].Cells[3].Value))
                    {
                        daysBehind = Convert.ToInt32((Convert.ToDateTime(dgvInitiation.Rows[i].Cells[2].Value) - Convert.ToDateTime(dgvInitiation.Rows[i].Cells[3].Value)).Days);
                        dgvInitiation.Rows[i].Cells[2].Style.ForeColor = Color.Red;
                    }
                    // Completed in time 
                    else if (Convert.ToDateTime(dgvInitiation.Rows[i].Cells[2].Value) < Convert.ToDateTime(dgvInitiation.Rows[i].Cells[3].Value))
                    {
                        daysAhead = Convert.ToInt32((Convert.ToDateTime(dgvInitiation.Rows[i].Cells[2].Value) - Convert.ToDateTime(dgvInitiation.Rows[i].Cells[3].Value)).Days);
                        dgvInitiation.Rows[i].Cells[2].Style.ForeColor = Color.LimeGreen;
                    }
                    else
                    {
                        dgvInitiation.Rows[i].Cells[2].Style.ForeColor = Color.Black;
                    }

                }
                else
                {
                    continue;
                }

                daysSpent = (daysSpent + (daysAhead + daysBehind));
                // Behind Schedule
                if (daysSpent > 0)
                {
                    totalDaysInitlbl.Text = daysSpent.ToString() + " day(s) behind schedule";
                    totalDaysInitlbl.ForeColor = Color.Red;
                }
                else if (daysSpent < 0)
                {
                    daysSpent *= -1;
                    totalDaysInitlbl.Text = daysSpent.ToString() + " day(s) ahead of schedule";
                    totalDaysInitlbl.ForeColor = Color.LimeGreen;
                    daysSpent *= -1;
                }
                else
                {
                    totalDaysInitlbl.Text = "On schedule";
                    totalDaysInitlbl.ForeColor = Color.Black;
                }

                budgetSpent = (budgetSpent + (budgetAhead + budgetBehind));

                if (budgetSpent > 0.0)
                {
                    lblTotalInitialBudget.Text = budgetSpent.ToString("C") + " under budget";
                    lblTotalInitialBudget.ForeColor = Color.LimeGreen;
                }
                else if (budgetSpent < 0.0)
                {
                    budgetSpent *= -1;
                    lblTotalInitialBudget.Text = budgetSpent.ToString("C") + " over budget";
                    lblTotalInitialBudget.ForeColor = Color.Red;
                    budgetSpent *= -1;
                }
                else
                {
                    lblTotalInitialBudget.Text = "On Budget";
                    lblTotalInitialBudget.ForeColor = Color.Black;
                }


            }
        }
    }

    // Methods

    
}


