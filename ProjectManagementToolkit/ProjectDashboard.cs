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
        List<string> initDocsListDueDate = new List<string>();
        List<string> planningDocsListDueDate = new List<string>();
        List<string> executeDocsListDueDate = new List<string>();
        List<string> closingListDueDate = new List<string>();

        private void ProjectDashboard_Load(object sender, EventArgs e)
        {
            pbarOverall.Hide();
            lblOverallProgress.Hide();


            List<string> planningDocsListStatus = new List<string>();
            List<string> executionDocsListStatus = new List<string>();
            List<string> closingDocsListStatus = new List<string>();

            string json = JsonHelper.loadProjectInfo(Settings.Default.Username);
            List<ProjectModel> projectListModel = JsonConvert.DeserializeObject<List<ProjectModel>>(json);
            projectModel = projectModel.getProjectModel(Settings.Default.ProjectID, projectListModel);


            /////////////////////////////////////////////////////////////////INITIATION PHASE/////////////////////////////////////////////////////////////////////////////////////////////////////////


            string jsonInitDue = JsonHelper.loadDocument(Settings.Default.ProjectID, "InitDueDateModel");
            InitDueDateModel tInit = JsonConvert.DeserializeObject<InitDueDateModel>(jsonInitDue);
            
            if(tInit != null)
            {
                initDocsListDueDate.Add(tInit.BusinessCaseDD);
                initDocsListDueDate.Add(tInit.FeasibilityStudyDD);
                initDocsListDueDate.Add(tInit.ProjectCharterDD);
                initDocsListDueDate.Add(tInit.JobDescriptionDD);
                initDocsListDueDate.Add(tInit.ProjectOfficeCheckListDD);
                initDocsListDueDate.Add(tInit.PhaseRevieFormInitiationDD);
                initDocsListDueDate.Add(tInit.TermOfReferenceDocument);
            }

            

            string jsonPlanningDue = JsonHelper.loadDocument(Settings.Default.ProjectID, "PlanningDueDateModel");
            PlanningDueDateModel tPlanning = JsonConvert.DeserializeObject<PlanningDueDateModel>(jsonPlanningDue);

            if(tPlanning != null)
            {
                planningDocsListDueDate.Add(tPlanning.ProjectPlan);
                planningDocsListDueDate.Add(tPlanning.ResourcePlan);
                planningDocsListDueDate.Add(tPlanning.FinancialPlan);
                planningDocsListDueDate.Add(tPlanning.QualityPlan);
                planningDocsListDueDate.Add(tPlanning.RiskPlan);
                planningDocsListDueDate.Add(tPlanning.AcceptancePlan);
                planningDocsListDueDate.Add(tPlanning.CommunicationPlan);
                planningDocsListDueDate.Add(tPlanning.ProcurementPlan);
                planningDocsListDueDate.Add(tPlanning.StatementOfWork);
                planningDocsListDueDate.Add(tPlanning.RequestForInformation);
                planningDocsListDueDate.Add(tPlanning.SupplierContract);
                planningDocsListDueDate.Add(tPlanning.RequestForProposal);
                planningDocsListDueDate.Add(tPlanning.PhaseReviewPlanning);
            }
            

            string jsonExecuteDue = JsonHelper.loadDocument(Settings.Default.ProjectID, "ExecutionDueDateModel");
            ExecutionDueDateModel tExecute = JsonConvert.DeserializeObject<ExecutionDueDateModel>(jsonExecuteDue);

            if(tExecute != null)
            {
                executeDocsListDueDate.Add(tExecute.TimeMangement);
                executeDocsListDueDate.Add(tExecute.TimeSheet);
                executeDocsListDueDate.Add(tExecute.TimeSheetRegister);
                executeDocsListDueDate.Add(tExecute.CostManagementProcess);
                executeDocsListDueDate.Add(tExecute.ExpenseForm);
                executeDocsListDueDate.Add(tExecute.ExpenseRegister);
                executeDocsListDueDate.Add(tExecute.QualityManagement);
                executeDocsListDueDate.Add(tExecute.QualityReviewPlan);
                executeDocsListDueDate.Add(tExecute.QualityReviewForm);
                executeDocsListDueDate.Add(tExecute.ChangeManagementProcess);
                executeDocsListDueDate.Add(tExecute.ChangeRequestForm);
                executeDocsListDueDate.Add(tExecute.ChangeRequestRegister);
                executeDocsListDueDate.Add(tExecute.RiskManagamentProcess);
                executeDocsListDueDate.Add(tExecute.RiskForm);
                executeDocsListDueDate.Add(tExecute.RiskRegister);
                executeDocsListDueDate.Add(tExecute.IssueManagementProcess);
                executeDocsListDueDate.Add(tExecute.IssueForm);
                executeDocsListDueDate.Add(tExecute.IssueRegister);
                executeDocsListDueDate.Add(tExecute.PurchaseOrder);
                executeDocsListDueDate.Add(tExecute.ProcurementRegister);
                executeDocsListDueDate.Add(tExecute.AcceptanceManagementProcess);
                executeDocsListDueDate.Add(tExecute.AcceptanceForm);
                executeDocsListDueDate.Add(tExecute.AcceptanceRegister);
                executeDocsListDueDate.Add(tExecute.CommunicationsManagementProcess);
                executeDocsListDueDate.Add(tExecute.ProjectStatusReport);
                executeDocsListDueDate.Add(tExecute.CommunicationsRegister);
                executeDocsListDueDate.Add(tExecute.PhaseReviewExe);
            }

           

            string jsonClosingDue = JsonHelper.loadDocument(Settings.Default.ProjectID, "ClosingDueDateModel");
            ClosingDueDateModel tClosing = JsonConvert.DeserializeObject<ClosingDueDateModel>(jsonClosingDue);

            if(tClosing != null)
            {
                closingListDueDate.Add(tClosing.ProjectClosureReport);
                closingListDueDate.Add(tClosing.PostImplementationReview);
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
            }
            else
                //IsBusinessCaseModelDone = "";
                initDocsListStatus.Add("");



            //////FEASIBILITY STUDY/////////
            string json2 = JsonHelper.loadDocument(Settings.Default.ProjectID, "FeasibilityStudy");
            versionControl1 = JsonConvert.DeserializeObject<VersionControl<FeasibilityStudyModel>>(json2);

            if (versionControl1 != null)
            {
                currentFeasibilityStudyModel = JsonConvert.DeserializeObject<FeasibilityStudyModel>(versionControl1.getLatest(versionControl1.DocumentModels));
                initDocsListStatus.Add(currentFeasibilityStudyModel.FeasibilityStudyProgress);
            }
            else
                initDocsListStatus.Add("");




            //////PROJECT CHARTER/////////
            string json3 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectCharter");
            versionControl2 = JsonConvert.DeserializeObject<VersionControl<ProjectCharterModel>>(json3);


            if (versionControl2 != null)
            {
                currentProjectCharter = JsonConvert.DeserializeObject<ProjectCharterModel>(versionControl2.getLatest(versionControl2.DocumentModels));
                initDocsListStatus.Add(currentProjectCharter.ProjectCharterProgress);
            }
            else

                initDocsListStatus.Add("");




            //////JOB DESCRIPTION/////////
            string json4 = JsonHelper.loadDocument(Settings.Default.ProjectID, "JobDescription");
            versionControl3 = JsonConvert.DeserializeObject<VersionControl<JobDescriptionModel>>(json4);

            if (versionControl3 != null)
            {
                currentJobDescription = JsonConvert.DeserializeObject<JobDescriptionModel>(versionControl3.getLatest(versionControl3.DocumentModels));
                initDocsListStatus.Add(currentJobDescription.JobDescriptionProgress);

            }
            else

                initDocsListStatus.Add("");


            //////PROJECT OFFICE CHECKLIST/////////
            string json5 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectOfficeCheckList");
            versionControl4 = JsonConvert.DeserializeObject<VersionControl<ProjectOfficeChecklistModel>>(json5);


            if (versionControl4 != null)
            {
                currentProjectOfficeChecklist = JsonConvert.DeserializeObject<ProjectOfficeChecklistModel>(versionControl4.getLatest(versionControl4.DocumentModels));
                initDocsListStatus.Add(currentProjectOfficeChecklist.ProjectOfficeCheckListProgress);

            }
            else
                initDocsListStatus.Add("");


            //////PHASE REVIEW FORM INITIATION/////////
            string json6 = JsonHelper.loadDocument(Settings.Default.ProjectID, "PhaseReviewFormInitiation");
            versionControl5 = JsonConvert.DeserializeObject<VersionControl<PhaseReviewFormInitiationModel>>(json6);


            if (versionControl5 != null)
            {
                currentPhaseReviewFormInitiation = JsonConvert.DeserializeObject<PhaseReviewFormInitiationModel>(versionControl5.getLatest(versionControl5.DocumentModels));
                initDocsListStatus.Add(currentPhaseReviewFormInitiation.PhaseReviewFormInitiationProgress);

            }
            else
                initDocsListStatus.Add("");


            //////TERMS OF REFERENCE/////////
            string json7 = JsonHelper.loadDocument(Settings.Default.ProjectID, "TermOfReferenceDocument");
            versionControl6 = JsonConvert.DeserializeObject<VersionControl<TermsOfReferenceModel>>(json7);


            if (versionControl6 != null)
            {
                currentTermOfReference = JsonConvert.DeserializeObject<TermsOfReferenceModel>(versionControl6.getLatest(versionControl6.DocumentModels));
                initDocsListStatus.Add(currentTermOfReference.TermOfReferenceProgress);

            }
            else
                initDocsListStatus.Add("");

            /////////////////////////////////////////////////////////////////PLANNING PHASE/////////////////////////////////////////////////////////////////////////////////////////////////////////

            //////ProjectPlan/////////
            string json8 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectPlan");
            VersionControl<ProjectPlanModel> versionControl7 = JsonConvert.DeserializeObject<VersionControl<ProjectPlanModel>>(json8);
            ProjectPlanModel currentProjectPlan;

            if (versionControl7 != null)
            {
                currentProjectPlan = JsonConvert.DeserializeObject<ProjectPlanModel>(versionControl7.getLatest(versionControl7.DocumentModels));
                planningDocsListStatus.Add(currentProjectPlan.projectPlanProgress);

            }
            else
                //  IsProjectPlanDone = "";
                planningDocsListStatus.Add("");

            //////ResourcePlan/////////
            string json9 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ResourcePlan");
            VersionControl<ResourcePlanModel> versionControl8 = JsonConvert.DeserializeObject<VersionControl<ResourcePlanModel>>(json9);
            ResourcePlanModel currentResourcePlan;

            if (versionControl8 != null)
            {
                currentResourcePlan = JsonConvert.DeserializeObject<ResourcePlanModel>(versionControl8.getLatest(versionControl8.DocumentModels));
                planningDocsListStatus.Add(currentResourcePlan.ResourcePlanProgress);

            }
            else
                //  IsResourcePlanDone = "";
                planningDocsListStatus.Add("");

            //////FinancialPlan/////////
            string json10 = JsonHelper.loadDocument(Settings.Default.ProjectID, "FinancialPlan");
            VersionControl<FinancialPlanModel> versionControl9 = JsonConvert.DeserializeObject<VersionControl<FinancialPlanModel>>(json10);
            FinancialPlanModel currentFinancialPlan;

            if (versionControl9 != null)
            {
                currentFinancialPlan = JsonConvert.DeserializeObject<FinancialPlanModel>(versionControl9.getLatest(versionControl9.DocumentModels));
                planningDocsListStatus.Add(currentFinancialPlan.FinancialPlanProgress);

            }
            else
                //  IsFinancialPlanDone = "";
                planningDocsListStatus.Add("");

            //////QualityPlan/////////
            string json11 = JsonHelper.loadDocument(Settings.Default.ProjectID, "QualityPlan");
            VersionControl<QualityPlanModel> versionControl10 = JsonConvert.DeserializeObject<VersionControl<QualityPlanModel>>(json11);
            QualityPlanModel currentQualityPlan;

            if (versionControl10 != null)
            {
                currentQualityPlan = JsonConvert.DeserializeObject<QualityPlanModel>(versionControl10.getLatest(versionControl10.DocumentModels));
                planningDocsListStatus.Add(currentQualityPlan.QualityPlanProgress);

            }
            else
                //  IsQualityPlanDone = "";
                planningDocsListStatus.Add("");

            //////RiskPlan/////////
            string json12 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RiskPlan");
            VersionControl<RiskPlanModel> versionControl11 = JsonConvert.DeserializeObject<VersionControl<RiskPlanModel>>(json12);
            RiskPlanModel currentRiskPlan;

            if (versionControl11 != null)
            {
                currentRiskPlan = JsonConvert.DeserializeObject<RiskPlanModel>(versionControl11.getLatest(versionControl11.DocumentModels));
                planningDocsListStatus.Add(currentRiskPlan.RiskPlanProgress);
            }
            else
                //  IsRiskPlanDone = "";
                planningDocsListStatus.Add("");

            //////AcceptancePlan/////////
            string json13 = JsonHelper.loadDocument(Settings.Default.ProjectID, "AcceptancePlan");
            VersionControl<AcceptancePlanModel> versionControl12 = JsonConvert.DeserializeObject<VersionControl<AcceptancePlanModel>>(json13);
            AcceptancePlanModel currentAcceptancePlan;

            if (versionControl12 != null)
            {
                currentAcceptancePlan = JsonConvert.DeserializeObject<AcceptancePlanModel>(versionControl12.getLatest(versionControl12.DocumentModels));
                planningDocsListStatus.Add(currentAcceptancePlan.AcceptancePlanProgress);

            }
            else
                //  IsAcceptancePlanDone = "";
                planningDocsListStatus.Add("");

            //////CommunicationPlan/////////
            string json14 = JsonHelper.loadDocument(Settings.Default.ProjectID, "CommunicationPlan");
            VersionControl<CommunicationsPlanModel> versionControl13 = JsonConvert.DeserializeObject<VersionControl<CommunicationsPlanModel>>(json14);
            CommunicationsPlanModel currentCommunicationPlan;

            if (versionControl13 != null)
            {
                currentCommunicationPlan = JsonConvert.DeserializeObject<CommunicationsPlanModel>(versionControl13.getLatest(versionControl13.DocumentModels));
                planningDocsListStatus.Add(currentCommunicationPlan.CommunicationPlanProgress);
            }
            else
                //  IsCommunicationPlanDone = "";
                planningDocsListStatus.Add("");

            //////ProcurementPlan/////////
            string json15 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProcurementPlan");
            VersionControl<ProcurementPlanModel> versionControl14 = JsonConvert.DeserializeObject<VersionControl<ProcurementPlanModel>>(json15);
            ProcurementPlanModel currentProcurementPlan;

            if (versionControl14 != null)
            {
                currentProcurementPlan = JsonConvert.DeserializeObject<ProcurementPlanModel>(versionControl14.getLatest(versionControl14.DocumentModels));
                planningDocsListStatus.Add(currentProcurementPlan.ProcurementPlanProgress);

            }
            else
                //  IsProcurementPlanDone = "";
                planningDocsListStatus.Add("");

            //////StatementOfWork/////////
            string json16 = JsonHelper.loadDocument(Settings.Default.ProjectID, "StatementOfWork");
            VersionControl<StatementOfWorkModel> versionControl15 = JsonConvert.DeserializeObject<VersionControl<StatementOfWorkModel>>(json16);
            StatementOfWorkModel currentStatementOfWork;

            if (versionControl15 != null)
            {
                currentStatementOfWork = JsonConvert.DeserializeObject<StatementOfWorkModel>(versionControl15.getLatest(versionControl15.DocumentModels));
                planningDocsListStatus.Add(currentStatementOfWork.StatementOfWorkProgress);

            }
            else
                //  IsStatementOfWorkDone = "";
                planningDocsListStatus.Add("");

            //////RequestForInformation/////////
            string json17 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RequestForInformation");
            VersionControl<RequestForInformationModel> versionControl16 = JsonConvert.DeserializeObject<VersionControl<RequestForInformationModel>>(json17);
            RequestForInformationModel currentRequestForInformation;

            if (versionControl16 != null)
            {
                currentRequestForInformation = JsonConvert.DeserializeObject<RequestForInformationModel>(versionControl16.getLatest(versionControl16.DocumentModels));
                planningDocsListStatus.Add(currentRequestForInformation.RequestForInformationProgress);

            }
            else
                //  IsRequestForInformationDone = "";
                planningDocsListStatus.Add("");

            //////SupplierContract/////////
            string json18 = JsonHelper.loadDocument(Settings.Default.ProjectID, "SupplierContract");
            VersionControl<SupplierContractModel> versionControl17 = JsonConvert.DeserializeObject<VersionControl<SupplierContractModel>>(json18);
            SupplierContractModel currentSupplierContract;

            if (versionControl17 != null)
            {
                currentSupplierContract = JsonConvert.DeserializeObject<SupplierContractModel>(versionControl17.getLatest(versionControl17.DocumentModels));
                planningDocsListStatus.Add(currentSupplierContract.SupplierContractProgress);
            }
            else
                //  IsSupplierContractDone = "";
                planningDocsListStatus.Add("");

            //////RequestForProposal/////////
            string json19 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RequestForProposal");
            VersionControl<RequestForProposalModel> versionControl18 = JsonConvert.DeserializeObject<VersionControl<RequestForProposalModel>>(json19);
            RequestForProposalModel currentRequestForProposal;

            if (versionControl18 != null)
            {
                currentRequestForProposal = JsonConvert.DeserializeObject<RequestForProposalModel>(versionControl18.getLatest(versionControl18.DocumentModels));
                planningDocsListStatus.Add(currentRequestForProposal.RequestForProposalProgress);

            }
            else
                //  IsRequestForProposalDone = "";
                planningDocsListStatus.Add("");

            //////PhaseReviewPlanning/////////
            string json20 = JsonHelper.loadDocument(Settings.Default.ProjectID, "PhaseReviewPlanning");
            VersionControl<PhaseReviewPlanningModel> versionControl19 = JsonConvert.DeserializeObject<VersionControl<PhaseReviewPlanningModel>>(json20);
            PhaseReviewPlanningModel currentPhaseReviewPlanning;

            if (versionControl19 != null)
            {
                currentPhaseReviewPlanning = JsonConvert.DeserializeObject<PhaseReviewPlanningModel>(versionControl19.getLatest(versionControl19.DocumentModels));
                planningDocsListStatus.Add(currentPhaseReviewPlanning.PhaseReviewPlanningProgress);

            }
            else
                //  IsPhaseReviewPlanningDone = "";
                planningDocsListStatus.Add("");


            /////////////////////////////////////////////////////////////////EXECUTION PHASE/////////////////////////////////////////////////////////////////////////////////////////////////////////

            //////TimeManagement Process/////////
            string json21 = JsonHelper.loadDocument(Settings.Default.ProjectID, "TimeMangement");
            VersionControl<TimeMangementProcessModel> versionControl20 = JsonConvert.DeserializeObject<VersionControl<TimeMangementProcessModel>>(json21);
            TimeMangementProcessModel currentTimeManagementProcess;

            if (versionControl20 != null)
            {
                currentTimeManagementProcess = JsonConvert.DeserializeObject<TimeMangementProcessModel>(versionControl20.getLatest(versionControl20.DocumentModels));
                executionDocsListStatus.Add(currentTimeManagementProcess.TimeManagementProcessProgress);
            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////TimeSheet Process/////////
            string json22 = JsonHelper.loadDocument(Settings.Default.ProjectID, "TimeSheet");
            VersionControl<TimeSheetModel> versionControl21 = JsonConvert.DeserializeObject<VersionControl<TimeSheetModel>>(json22);
            TimeSheetModel currentTimeSheet;

            if (versionControl21 != null)
            {
                currentTimeSheet = JsonConvert.DeserializeObject<TimeSheetModel>(versionControl21.getLatest(versionControl21.DocumentModels));
                executionDocsListStatus.Add(currentTimeSheet.TimeSheetProgress);
            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////TimeSheet Register/////////
            string json23 = JsonHelper.loadDocument(Settings.Default.ProjectID, "TimeSheetRegister");
            VersionControl<TimesheetRegisterModel.TimesheetEntry> versionControl22 = JsonConvert.DeserializeObject<VersionControl<TimesheetRegisterModel.TimesheetEntry>>(json23);
            TimesheetRegisterModel.TimesheetEntry currentTimeSheetRegister;

            if (versionControl22 != null)
            {
                currentTimeSheetRegister = JsonConvert.DeserializeObject<TimesheetRegisterModel.TimesheetEntry>(versionControl22.getLatest(versionControl22.DocumentModels));
                executionDocsListStatus.Add(currentTimeSheetRegister.TimeSheetRegisterProgress);
            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////CostManagement Process/////////
            string json24 = JsonHelper.loadDocument(Settings.Default.ProjectID, "CostManagementProcess");
            VersionControl<CostManagementProcessModel> versionControl23 = JsonConvert.DeserializeObject<VersionControl<CostManagementProcessModel>>(json24);
            CostManagementProcessModel currentCostManagementProcess;

            if (versionControl23 != null)
            {
                currentCostManagementProcess = JsonConvert.DeserializeObject<CostManagementProcessModel>(versionControl23.getLatest(versionControl23.DocumentModels));
                executionDocsListStatus.Add(currentCostManagementProcess.CostManagementProcessProgress);
            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////ExpenseForm Process/////////
            string json25 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ExpenseForm");
            VersionControl<ExpenseFormModel> versionControl24 = JsonConvert.DeserializeObject<VersionControl<ExpenseFormModel>>(json25);
            ExpenseFormModel currentExpenseFormProcess;

            if (versionControl24 != null)
            {
                currentExpenseFormProcess = JsonConvert.DeserializeObject<ExpenseFormModel>(versionControl24.getLatest(versionControl24.DocumentModels));
                executionDocsListStatus.Add(currentExpenseFormProcess.ExpenseFormProgress);

            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////ExpenseRegister Process/////////
            string json26 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ExpenseRegister");
            VersionControl<ProjectManagementToolkit.MPMM.MPMM_Document_Models.ExpenseRegister.ExpenseEntry> versionControl25 = JsonConvert.DeserializeObject<VersionControl<ProjectManagementToolkit.MPMM.MPMM_Document_Models.ExpenseRegister.ExpenseEntry>>(json26);
            ProjectManagementToolkit.MPMM.MPMM_Document_Models.ExpenseRegister.ExpenseEntry currentExpenseRegister;

            if (versionControl25 != null)
            {
                currentExpenseRegister = JsonConvert.DeserializeObject<ProjectManagementToolkit.MPMM.MPMM_Document_Models.ExpenseRegister.ExpenseEntry>(versionControl25.getLatest(versionControl25.DocumentModels));
                executionDocsListStatus.Add(currentExpenseRegister.ExpenseRegisterProgress);

            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////QualityManagement Process/////////
            string json27 = JsonHelper.loadDocument(Settings.Default.ProjectID, "QualityManagement");
            VersionControl<QualityManagementProcessModel> versionControl26 = JsonConvert.DeserializeObject<VersionControl<QualityManagementProcessModel>>(json27);
            QualityManagementProcessModel currentQualityMnagementProcess;

            if (versionControl26 != null)
            {
                currentQualityMnagementProcess = JsonConvert.DeserializeObject<QualityManagementProcessModel>(versionControl26.getLatest(versionControl26.DocumentModels));
                executionDocsListStatus.Add(currentQualityMnagementProcess.QualityManagementProcessProgress);

            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");



            //////QualityReviewPlan Process/////////
            string json28 = JsonHelper.loadDocument(Settings.Default.ProjectID, "QualityReviewPlan");
            VersionControl<QualityReviewPlanModel> versionControl27 = JsonConvert.DeserializeObject<VersionControl<QualityReviewPlanModel>>(json28);
            QualityReviewPlanModel currentQualityReviewPlan;

            if (versionControl27 != null)
            {
                currentQualityReviewPlan = JsonConvert.DeserializeObject<QualityReviewPlanModel>(versionControl27.getLatest(versionControl27.DocumentModels));
                executionDocsListStatus.Add(currentQualityReviewPlan.QualityReviewPlanProgress);

            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////QualityReviewForm Process/////////
            string json29 = JsonHelper.loadDocument(Settings.Default.ProjectID, "QualityReviewForm");
            VersionControl<QualityRegisterModel.ConformanceOfProcess> versionControl28 = JsonConvert.DeserializeObject<VersionControl<QualityRegisterModel.ConformanceOfProcess>>(json29);
            QualityRegisterModel.ConformanceOfProcess currentQualityReviewForm;

            if (versionControl28 != null)
            {
                currentQualityReviewForm = JsonConvert.DeserializeObject<QualityRegisterModel.ConformanceOfProcess>(versionControl28.getLatest(versionControl28.DocumentModels));
                executionDocsListStatus.Add(currentQualityReviewForm.QualityRegisterProgress);

            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////ChangeManagementProcess Process/////////
            string json30 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ChangeManagementProcess");
            VersionControl<ChangeManagementProcessModel> versionControl29 = JsonConvert.DeserializeObject<VersionControl<ChangeManagementProcessModel>>(json30);
            ChangeManagementProcessModel currentChangeManagementProcess;

            if (versionControl29 != null)
            {
                currentChangeManagementProcess = JsonConvert.DeserializeObject<ChangeManagementProcessModel>(versionControl29.getLatest(versionControl29.DocumentModels));
                executionDocsListStatus.Add(currentChangeManagementProcess.ChangeManagementProcessProgress);

            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////ChangeRequestForm Process/////////
            string json31 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ChangeRequestForm");
            VersionControl<ChangeRequestModel> versionControl30 = JsonConvert.DeserializeObject<VersionControl<ChangeRequestModel>>(json31);
            ChangeRequestModel currentChangeRequestForm;

            if (versionControl30 != null)
            {
                currentChangeRequestForm = JsonConvert.DeserializeObject<ChangeRequestModel>(versionControl30.getLatest(versionControl30.DocumentModels));
                executionDocsListStatus.Add(currentChangeRequestForm.ChangeRequestProgress);

            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////ChangeRequestRegister Process/////////
            string json32 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ChangeRequestRegister");
            VersionControl<ChangeRegisterModel> versionControl31 = JsonConvert.DeserializeObject<VersionControl<ChangeRegisterModel>>(json32);
            ChangeRegisterModel currentChangeRegister;

            if (versionControl31 != null)
            {
                currentChangeRegister = JsonConvert.DeserializeObject<ChangeRegisterModel>(versionControl31.getLatest(versionControl31.DocumentModels));
                executionDocsListStatus.Add(currentChangeRegister.ChangeRegisterProgress);

            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////RiskManagamentProcess Process/////////
            string json33 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RiskManagamentProcess");
            VersionControl<RiskManagmentProcessModel> versionControl32 = JsonConvert.DeserializeObject<VersionControl<RiskManagmentProcessModel>>(json33);
            RiskManagmentProcessModel currentRiskManagementProcess;

            if (versionControl32 != null)
            {
                currentRiskManagementProcess = JsonConvert.DeserializeObject<RiskManagmentProcessModel>(versionControl32.getLatest(versionControl32.DocumentModels));
                executionDocsListStatus.Add(currentRiskManagementProcess.RiskManagementProcessProgress);

            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");


            //////RiskForm/////////
            string json34 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RiskForm");
            VersionControl<RiskFormModel> versionControl33 = JsonConvert.DeserializeObject<VersionControl<RiskFormModel>>(json34);
            RiskFormModel currentRiskForm;

            if (versionControl33 != null)
            {
                currentRiskForm = JsonConvert.DeserializeObject<RiskFormModel>(versionControl33.getLatest(versionControl33.DocumentModels));
                executionDocsListStatus.Add(currentRiskForm.RiskFormProgress);

            }
            else
                //  IsPhaseReviewPlanningDone = "";
                executionDocsListStatus.Add("");

            //////RiskRegister/////////
            string json35 = JsonHelper.loadDocument(Settings.Default.ProjectID, "RiskRegister");
            VersionControl<RiskRegisterModel> versionControl34 = JsonConvert.DeserializeObject<VersionControl<RiskRegisterModel>>(json35);
            RiskRegisterModel currentRiskRegister;

            if (versionControl34 != null)
            {
                currentRiskRegister = JsonConvert.DeserializeObject<RiskRegisterModel>(versionControl34.getLatest(versionControl34.DocumentModels));
                executionDocsListStatus.Add(currentRiskRegister.RiskRegisterProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////IssueManagementProcess/////////
            string json36 = JsonHelper.loadDocument(Settings.Default.ProjectID, "IssueManagementProcess");
            VersionControl<IssueManagementProcessModel> versionControl35 = JsonConvert.DeserializeObject<VersionControl<IssueManagementProcessModel>>(json36);
            IssueManagementProcessModel currentIssueManagementProcess;

            if (versionControl35 != null)
            {
                currentIssueManagementProcess = JsonConvert.DeserializeObject<IssueManagementProcessModel>(versionControl35.getLatest(versionControl35.DocumentModels));
                executionDocsListStatus.Add(currentIssueManagementProcess.IssueManagementProcessProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////IssueForm/////////
            string json37 = JsonHelper.loadDocument(Settings.Default.ProjectID, "IssueForm");
            VersionControl<IssueFormModel> versionControl36 = JsonConvert.DeserializeObject<VersionControl<IssueFormModel>>(json37);
            IssueFormModel currentIssueForm;

            if (versionControl36 != null)
            {
                currentIssueForm = JsonConvert.DeserializeObject<IssueFormModel>(versionControl36.getLatest(versionControl36.DocumentModels));
                executionDocsListStatus.Add(currentIssueForm.IssueFormProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////IssueRegister/////////
            string json38 = JsonHelper.loadDocument(Settings.Default.ProjectID, "IssueRegister");
            VersionControl<IssueRegisterModel> versionControl37 = JsonConvert.DeserializeObject<VersionControl<IssueRegisterModel>>(json38);
            IssueRegisterModel currentIssueRegister;

            if (versionControl37 != null)
            {
                currentIssueRegister = JsonConvert.DeserializeObject<IssueRegisterModel>(versionControl37.getLatest(versionControl37.DocumentModels));
                executionDocsListStatus.Add(currentIssueRegister.IssueRegisterProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////PurchaseOrder/////////
            string json39 = JsonHelper.loadDocument(Settings.Default.ProjectID, "PurchaseOrder");
            VersionControl<PurchaseOrderModel> versionControl38 = JsonConvert.DeserializeObject<VersionControl<PurchaseOrderModel>>(json39);
            PurchaseOrderModel currentPurchaseOrder;

            if (versionControl38 != null)
            {
                currentPurchaseOrder = JsonConvert.DeserializeObject<PurchaseOrderModel>(versionControl38.getLatest(versionControl38.DocumentModels));
                executionDocsListStatus.Add(currentPurchaseOrder.PurchaseOrderProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////ProcurementRegister/////////
            string json40 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProcurementRegister");
            VersionControl<ProcurementRegisterModel> versionControl39 = JsonConvert.DeserializeObject<VersionControl<ProcurementRegisterModel>>(json40);
            ProcurementRegisterModel currentProcurementRegister;

            if (versionControl39 != null)
            {
                currentProcurementRegister = JsonConvert.DeserializeObject<ProcurementRegisterModel>(versionControl39.getLatest(versionControl39.DocumentModels));
                executionDocsListStatus.Add(currentProcurementRegister.ProcurementRegisterProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////AcceptanceManagementProcess/////////
            string json41 = JsonHelper.loadDocument(Settings.Default.ProjectID, "AcceptanceManagementProcess");
            VersionControl<AcceptanceManagementProcessModel> versionControl40 = JsonConvert.DeserializeObject<VersionControl<AcceptanceManagementProcessModel>>(json41);
            AcceptanceManagementProcessModel currentAcceptanceManagementProcess;

            if (versionControl40 != null)
            {
                currentAcceptanceManagementProcess = JsonConvert.DeserializeObject<AcceptanceManagementProcessModel>(versionControl40.getLatest(versionControl40.DocumentModels));
                executionDocsListStatus.Add(currentAcceptanceManagementProcess.AcceptanceManagementProcessProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////AcceptanceForm/////////
            string json42 = JsonHelper.loadDocument(Settings.Default.ProjectID, "AcceptanceForm");
            VersionControl<AcceptanceFormModel> versionControl41 = JsonConvert.DeserializeObject<VersionControl<AcceptanceFormModel>>(json42);
            AcceptanceFormModel currentAcceptanceForm;

            if (versionControl41 != null)
            {
                currentAcceptanceForm = JsonConvert.DeserializeObject<AcceptanceFormModel>(versionControl41.getLatest(versionControl41.DocumentModels));
                executionDocsListStatus.Add(currentAcceptanceForm.AcceptanceFormProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////AcceptanceRegister/////////
            string json43 = JsonHelper.loadDocument(Settings.Default.ProjectID, "AcceptanceRegister");
            VersionControl<AcceptanceRegisterModel> versionControl42 = JsonConvert.DeserializeObject<VersionControl<AcceptanceRegisterModel>>(json43);
            AcceptanceRegisterModel currentAcceptanceRegister;

            if (versionControl42 != null)
            {
                currentAcceptanceRegister = JsonConvert.DeserializeObject<AcceptanceRegisterModel>(versionControl42.getLatest(versionControl42.DocumentModels));
                executionDocsListStatus.Add(currentAcceptanceRegister.AcceptanceRegisterProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////CommunicationsManagementProcess/////////
            string json44 = JsonHelper.loadDocument(Settings.Default.ProjectID, "CommunicationsManagementProcess");
            VersionControl<CommunicationsManagementProcessModel> versionControl43 = JsonConvert.DeserializeObject<VersionControl<CommunicationsManagementProcessModel>>(json44);
            CommunicationsManagementProcessModel currentCommunicationsManagementProcess;

            if (versionControl43 != null)
            {
                currentCommunicationsManagementProcess = JsonConvert.DeserializeObject<CommunicationsManagementProcessModel>(versionControl43.getLatest(versionControl43.DocumentModels));
                executionDocsListStatus.Add(currentCommunicationsManagementProcess.CommunicationsManagementProcessProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////ProjectStatusReport/////////
            string json45 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectStatusReport");
            VersionControl<ProjectStatusReportModel> versionControl44 = JsonConvert.DeserializeObject<VersionControl<ProjectStatusReportModel>>(json45);
            ProjectStatusReportModel currentProjectStatusReport;

            if (versionControl44 != null)
            {
                currentProjectStatusReport = JsonConvert.DeserializeObject<ProjectStatusReportModel>(versionControl44.getLatest(versionControl44.DocumentModels));
                executionDocsListStatus.Add(currentProjectStatusReport.ProjectStatusReportProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////CommunicationsRegister/////////
            string json46 = JsonHelper.loadDocument(Settings.Default.ProjectID, "CommunicationsRegister");
            VersionControl<CommunicationRegisterModel> versionControl45 = JsonConvert.DeserializeObject<VersionControl<CommunicationRegisterModel>>(json46);
            CommunicationRegisterModel currentCommunicationsRegister;

            if (versionControl45 != null)
            {
                currentCommunicationsRegister = JsonConvert.DeserializeObject<CommunicationRegisterModel>(versionControl45.getLatest(versionControl45.DocumentModels));
                executionDocsListStatus.Add(currentCommunicationsRegister.CommunicationsRegisterProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////PhaseReviewExe/////////
            string json47 = JsonHelper.loadDocument(Settings.Default.ProjectID, "PhaseReviewExe");
            VersionControl<PhaseReviewFormExecutionModel> versionControl46 = JsonConvert.DeserializeObject<VersionControl<PhaseReviewFormExecutionModel>>(json47);
            PhaseReviewFormExecutionModel currentPhaseReviewExe;

            if (versionControl46 != null)
            {
                currentPhaseReviewExe = JsonConvert.DeserializeObject<PhaseReviewFormExecutionModel>(versionControl46.getLatest(versionControl46.DocumentModels));
                executionDocsListStatus.Add(currentPhaseReviewExe.PhaseReviewExeProgress);

            }
            else
                executionDocsListStatus.Add("");

            //////////////////////////////////////////////////////CLOSING PHASE///////////////////////////////////////////////////////////////////////////////////
            //////ProjectClosureReport/////////
            string json48 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectClosureReport");
            VersionControl<ProjectClosureReportModel> versionControl47 = JsonConvert.DeserializeObject<VersionControl<ProjectClosureReportModel>>(json48);
            ProjectClosureReportModel currentProjectClosureReport;

            if (versionControl47 != null)
            {
                currentProjectClosureReport = JsonConvert.DeserializeObject<ProjectClosureReportModel>(versionControl47.getLatest(versionControl47.DocumentModels));
                closingDocsListStatus.Add(currentProjectClosureReport.ProjectClosureReportProgress);

            }
            else
                closingDocsListStatus.Add("");

            //////PostImplementationReview/////////
            string json49 = JsonHelper.loadDocument(Settings.Default.ProjectID, "PostImplementationReview");
            VersionControl<PostImplementationReviewModel> versionControl48 = JsonConvert.DeserializeObject<VersionControl<PostImplementationReviewModel>>(json49);
            PostImplementationReviewModel currentPostImplementationReview;

            if (versionControl48 != null)
            {
                currentPostImplementationReview = JsonConvert.DeserializeObject<PostImplementationReviewModel>(versionControl48.getLatest(versionControl48.DocumentModels));
                closingDocsListStatus.Add(currentPostImplementationReview.PostImplementationReviewProgress);

            }
            else
                closingDocsListStatus.Add("");

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
                MessageBox.Show("No documents have been added yet.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                tabInitiation.Visible = false;
                chart1.Visible = false;
                chart2.Visible = false;
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
                    dgvInitiation.Rows[i].Cells[1].Value = "";

                    if (initDocsListDueDate.Count > 0)
                        dgvInitiation.Rows[i].Cells[1].Value = initDocsListDueDate[i];//initDocsListStatus[i].dueDate.ToString();

                    if (initDocsListStatus[i] == "UNDONE")
                    {
                        dgvInitiation.Rows[i].Cells[2].Style.BackColor = Color.Orange;
                        inprog++;
                    }
                    else if (initDocsListStatus[i] == "DONE")
                    {
                        initationProgressVal++;
                        ///////////////////////////AL DIE CODE OM TE CHECK OF IETS VOOR IETS ANDERS GEDOEN IS/////////////////////
                        k = i;

                        comp++;
                        dgvInitiation.Rows[i].Cells[2].Style.BackColor = Color.LimeGreen;
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
                        dgvInitiation.Rows[i].Cells[2].Style.BackColor = Color.Gray;
                    }


                }

                for (int j = 0; j < k; j++)
                {
                    if (initDocsListStatus[j] == "") //Check if the previous tasks are not done or in progress, because then they are behind schedule
                    {
                        //Increment the behind schedule tasks
                        behind++;
                        //Set all the tasks that are behind schedule to display red
                        dgvInitiation.Rows[j].Cells[2].Style.BackColor = Color.Red;
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
                    dgvPlanning.Rows[i].Cells[1].Value = "";

                    if (planningDocsListDueDate.Count > 0)
                        dgvPlanning.Rows[i].Cells[1].Value = planningDocsListDueDate[i];

                    if (planningDocsListStatus[i] == "UNDONE")
                    {
                        dgvPlanning.Rows[i].Cells[2].Style.BackColor = Color.Orange;
                        inprogPlanning++;
                    }
                    else if (planningDocsListStatus[i] == "DONE")
                    {
                        planningProgressVal++;

                        k = i;

                        compPlanning++;
                        dgvPlanning.Rows[i].Cells[2].Style.BackColor = Color.LimeGreen;
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
                        dgvPlanning.Rows[i].Cells[2].Style.BackColor = Color.Gray;
                    }
                }

                for (int j = 0; j < k; j++)
                {
                    if (planningDocsListStatus[j] == "") //Check if the previous tasks are not done or in progress, because then they are behind schedule
                    {
                        //Increment the behind schedule tasks
                        behindPlanning++;
                        //Set all the tasks that are behind schedule to display red
                        dgvPlanning.Rows[j].Cells[2].Style.BackColor = Color.Red;
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
                    dgvExecution.Rows[i].Cells[1].Value = "";

                    if (executeDocsListDueDate.Count > 0)
                        dgvExecution.Rows[i].Cells[1].Value = executeDocsListDueDate[i];


                    if (executionDocsListStatus[i] == "UNDONE")
                    {
                        dgvExecution.Rows[i].Cells[2].Style.BackColor = Color.Orange;
                        inprogExecution++;
                    }
                    else if (executionDocsListStatus[i] == "DONE")
                    {
                        executionProgressVal++;

                        k = i;

                        compExecution++;
                        dgvExecution.Rows[i].Cells[2].Style.BackColor = Color.LimeGreen;
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
                        dgvExecution.Rows[i].Cells[2].Style.BackColor = Color.Gray;
                    }
                }

                for (int j = 0; j < k; j++)
                {
                    if (executionDocsListStatus[j] == "") //Check if the previous tasks are not done or in progress, because then they are behind schedule
                    {
                        //Increment the behind schedule tasks
                        behindExecution++;
                        //Set all the tasks that are behind schedule to display red
                        dgvExecution.Rows[j].Cells[2].Style.BackColor = Color.Red;
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
                    dgvClosing.Rows[i].Cells[1].Value = "";

                    if (closingListDueDate.Count > 0)
                        dgvClosing.Rows[i].Cells[1].Value = closingListDueDate[i];

                    if (closingDocsListStatus[i] == "UNDONE")
                    {
                        dgvClosing.Rows[i].Cells[2].Style.BackColor = Color.Orange;
                        inprogClosing++;
                    }
                    else if (closingDocsListStatus[i] == "DONE")
                    {
                        closingProgressVal++;

                        k = i;

                        compClosing++;
                        dgvClosing.Rows[i].Cells[2].Style.BackColor = Color.LimeGreen;
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
                        dgvClosing.Rows[i].Cells[2].Style.BackColor = Color.Gray;
                    }
                }

                for (int j = 0; j < k; j++)
                {
                    if (closingDocsListStatus[j] == "") //Check if the previous tasks are not done or in progress, because then they are behind schedule
                    {
                        //Increment the behind schedule tasks
                        behindClosing++;
                        //Set all the tasks that are behind schedule to display red
                        dgvClosing.Rows[j].Cells[2].Style.BackColor = Color.Red;
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


        private void saveAllDueDate(int phase)
        {

            if (phase == 1)
            {
                currentInit.BusinessCaseDD = dgvInitiation.Rows[0].Cells[1].Value.ToString();
                currentInit.FeasibilityStudyDD = dgvInitiation.Rows[1].Cells[1].Value.ToString();
                currentInit.ProjectCharterDD = dgvInitiation.Rows[2].Cells[1].Value.ToString();
                currentInit.JobDescriptionDD = dgvInitiation.Rows[3].Cells[1].Value.ToString();
                currentInit.ProjectOfficeCheckListDD = dgvInitiation.Rows[4].Cells[1].Value.ToString();
                currentInit.PhaseRevieFormInitiationDD = dgvInitiation.Rows[5].Cells[1].Value.ToString();
                currentInit.TermOfReferenceDocument = dgvInitiation.Rows[6].Cells[1].Value.ToString();

                string jsong = JsonConvert.SerializeObject(currentInit);
                JsonHelper.saveDocument(jsong, Settings.Default.ProjectID, "InitDueDateModel");
            }
            else if (phase == 2)
            {
                MessageBox.Show(dgvPlanning.Rows[0].Cells[1].Value.ToString());


                currentPlan.ProjectPlan = dgvPlanning.Rows[0].Cells[1].Value.ToString();
                currentPlan.ResourcePlan = dgvPlanning.Rows[1].Cells[1].Value.ToString();
                currentPlan.FinancialPlan = dgvPlanning.Rows[2].Cells[1].Value.ToString();
                currentPlan.QualityPlan = dgvPlanning.Rows[3].Cells[1].Value.ToString();
                currentPlan.RiskPlan = dgvPlanning.Rows[4].Cells[1].Value.ToString();
                currentPlan.AcceptancePlan = dgvPlanning.Rows[5].Cells[1].Value.ToString();
                currentPlan.CommunicationPlan = dgvPlanning.Rows[6].Cells[1].Value.ToString();
                currentPlan.ProcurementPlan = dgvPlanning.Rows[7].Cells[1].Value.ToString();
                currentPlan.StatementOfWork = dgvPlanning.Rows[8].Cells[1].Value.ToString();
                currentPlan.RequestForInformation = dgvPlanning.Rows[9].Cells[1].Value.ToString();
                currentPlan.SupplierContract = dgvPlanning.Rows[10].Cells[1].Value.ToString();
                currentPlan.RequestForProposal = dgvPlanning.Rows[11].Cells[1].Value.ToString();
                currentPlan.PhaseReviewPlanning = dgvPlanning.Rows[12].Cells[1].Value.ToString();

                string jsong = JsonConvert.SerializeObject(currentPlan);
                JsonHelper.saveDocument(jsong, Settings.Default.ProjectID, "PlanningDueDateModel");
            }
            else if (phase == 3)
            {
                currentExecute.TimeMangement = dgvExecution.Rows[0].Cells[1].Value.ToString();
                currentExecute.TimeSheet = dgvExecution.Rows[1].Cells[1].Value.ToString();
                currentExecute.TimeSheetRegister = dgvExecution.Rows[2].Cells[1].Value.ToString();
                currentExecute.CostManagementProcess = dgvExecution.Rows[3].Cells[1].Value.ToString();
                currentExecute.ExpenseForm = dgvExecution.Rows[4].Cells[1].Value.ToString();
                currentExecute.ExpenseRegister = dgvExecution.Rows[5].Cells[1].Value.ToString();
                currentExecute.QualityManagement = dgvExecution.Rows[6].Cells[1].Value.ToString();
                currentExecute.QualityReviewPlan = dgvExecution.Rows[7].Cells[1].Value.ToString();
                currentExecute.QualityReviewForm = dgvExecution.Rows[8].Cells[1].Value.ToString();
                currentExecute.ChangeManagementProcess = dgvExecution.Rows[9].Cells[1].Value.ToString();
                currentExecute.ChangeRequestForm = dgvExecution.Rows[10].Cells[1].Value.ToString();
                currentExecute.ChangeRequestRegister = dgvExecution.Rows[11].Cells[1].Value.ToString();
                currentExecute.RiskManagamentProcess = dgvExecution.Rows[12].Cells[1].Value.ToString();
                currentExecute.RiskForm = dgvExecution.Rows[13].Cells[1].Value.ToString();
                currentExecute.RiskRegister = dgvExecution.Rows[14].Cells[1].Value.ToString();
                currentExecute.IssueManagementProcess = dgvExecution.Rows[15].Cells[1].Value.ToString();
                currentExecute.IssueForm = dgvExecution.Rows[16].Cells[1].Value.ToString();
                currentExecute.IssueRegister = dgvExecution.Rows[17].Cells[1].Value.ToString();
                currentExecute.PurchaseOrder = dgvExecution.Rows[18].Cells[1].Value.ToString();
                currentExecute.ProcurementRegister = dgvExecution.Rows[19].Cells[1].Value.ToString();
                currentExecute.AcceptanceManagementProcess = dgvExecution.Rows[20].Cells[1].Value.ToString();
                currentExecute.AcceptanceForm = dgvExecution.Rows[21].Cells[1].Value.ToString();
                currentExecute.AcceptanceRegister = dgvExecution.Rows[22].Cells[1].Value.ToString();
                currentExecute.CommunicationsManagementProcess = dgvExecution.Rows[23].Cells[1].Value.ToString();
                currentExecute.ProjectStatusReport = dgvExecution.Rows[24].Cells[1].Value.ToString();
                currentExecute.CommunicationsRegister = dgvExecution.Rows[25].Cells[1].Value.ToString();
                currentExecute.PhaseReviewExe = dgvExecution.Rows[26].Cells[1].Value.ToString();

                string jsong = JsonConvert.SerializeObject(currentExecute);
                JsonHelper.saveDocument(jsong, Settings.Default.ProjectID, "ExecutionDueDateModel");
            }
            else if (phase == 4)
            {
                currentClose.ProjectClosureReport = dgvClosing.Rows[0].Cells[1].Value.ToString();
                currentClose.PostImplementationReview = dgvClosing.Rows[1].Cells[1].Value.ToString();

                string jsong = JsonConvert.SerializeObject(currentClose);
                JsonHelper.saveDocument(jsong, Settings.Default.ProjectID, "ClosingDueDateModel");
            }

        }


        private void dgvInitiation_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // If any cell is clicked on the Second column which is our date Column  
            if (e.ColumnIndex == 1 && e.RowIndex != -1)
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

        private void InitDateTimePicker_OnTextChange(object sender, EventArgs e)
        {
            dgvInitiation.CurrentCell.Value = InitDateTimePicker.Text.ToString();
            saveAllDueDate(1);
        }

        void InitDateTimePicker_CloseUp(object sender, EventArgs e)
        {
            // Hiding the control after use   
            InitDateTimePicker.Visible = false;
        }

        private void dgvPlanning_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1 && e.RowIndex != -1)
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

        private void PlanDateTimePicker_OnTextChange(object sender, EventArgs e)
        {
            dgvPlanning.CurrentCell.Value = PlanDateTimePicker.Text.ToString();
            saveAllDueDate(2);
        }


        void PlanDateTimePicker_CloseUp(object sender, EventArgs e)
        {
            // Hiding the control after use   
            PlanDateTimePicker.Visible = false;
        }

        private void dgvExecution_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1 && e.RowIndex != -1)
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


        private void ExecuteDateTimePicker_OnTextChange(object sender, EventArgs e)
        {
            dgvExecution.CurrentCell.Value = ExecuteDateTimePicker.Text.ToString();
            saveAllDueDate(3);
        }


        void ExecuteDateTimePicker_CloseUp(object sender, EventArgs e)
        {
            // Hiding the control after use   
            ExecuteDateTimePicker.Visible = false;
        }

        private void CloseDateTimePicker_OnTextChange(object sender, EventArgs e)
        {
            dgvClosing.CurrentCell.Value = CloseDateTimePicker.Text.ToString();
            saveAllDueDate(4);
        }


        void CloseDateTimePicker_CloseUp(object sender, EventArgs e)
        {
            // Hiding the control after use   
            CloseDateTimePicker.Visible = false;
        }

        private void dgvClosing_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1 && e.RowIndex != -1)
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
}


