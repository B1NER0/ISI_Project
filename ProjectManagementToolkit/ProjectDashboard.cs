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

        private void ProjectDashboard_Load(object sender, EventArgs e)
        {

            pbarOverall.Hide();
            lblOverallProgress.Hide();

            List<string> initDocsListStatus = new List<string>();
            List<string> planningDocsListStatus = new List<string>();
            List<string> executionDocsListStatus = new List<string>();

            string json = JsonHelper.loadProjectInfo(Settings.Default.Username);
            List<ProjectModel> projectListModel = JsonConvert.DeserializeObject<List<ProjectModel>>(json);
            projectModel = projectModel.getProjectModel(Settings.Default.ProjectID, projectListModel);


            /////////////////////////////////////////////////////////////////INITIATION PHASE/////////////////////////////////////////////////////////////////////////////////////////////////////////

            ////////BUSINESSCASE////////
            //Verander Json
            string json1 = JsonHelper.loadDocument(Settings.Default.ProjectID, "BusinessCase");

            //Check versions
            VersionControl<BusinessCaseModel> versionControl = JsonConvert.DeserializeObject<VersionControl<BusinessCaseModel>>(json1);
            //Get current businesscaseModel
            BusinessCaseModel currentBusinessCaseModel;

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
            VersionControl<FeasibilityStudyModel> versionControl1 = JsonConvert.DeserializeObject<VersionControl<FeasibilityStudyModel>>(json2);
            FeasibilityStudyModel currentFeasibilityStudyModel; //= JsonConvert.DeserializeObject<FeasibilityStudyModel>(versionControl1.getLatest(versionControl1.DocumentModels));


            if (versionControl1 != null)
            {
                currentFeasibilityStudyModel = JsonConvert.DeserializeObject<FeasibilityStudyModel>(versionControl1.getLatest(versionControl1.DocumentModels));
                initDocsListStatus.Add(currentFeasibilityStudyModel.FeasibilityStudyProgress);
            }
            else
                // IsFeasibilityStudyDone = "";
                initDocsListStatus.Add("");




            //////PROJECT CHARTER/////////
            string json3 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectCharter");
            VersionControl<ProjectCharterModel> versionControl2 = JsonConvert.DeserializeObject<VersionControl<ProjectCharterModel>>(json3);
            ProjectCharterModel currentProjectCharter;

            if (versionControl2 != null)
            {
                currentProjectCharter = JsonConvert.DeserializeObject<ProjectCharterModel>(versionControl2.getLatest(versionControl2.DocumentModels));
                initDocsListStatus.Add(currentProjectCharter.ProjectCharterProgress);
            }
            else
                // IsProjectCharterDone = "";
                initDocsListStatus.Add("");




            //////JOB DESCRIPTION/////////
            string json4 = JsonHelper.loadDocument(Settings.Default.ProjectID, "JobDescription");
            VersionControl<JobDescriptionModel> versionControl3 = JsonConvert.DeserializeObject<VersionControl<JobDescriptionModel>>(json4);
            JobDescriptionModel currentJobDescription;

            if (versionControl3 != null)
            {
                currentJobDescription = JsonConvert.DeserializeObject<JobDescriptionModel>(versionControl3.getLatest(versionControl3.DocumentModels));
                initDocsListStatus.Add(currentJobDescription.JobDescriptionProgress);

            }
            else
                //  IsJobDescriptionDone = "";
                initDocsListStatus.Add("");


            //////PROJECT OFFICE CHECKLIST/////////
            string json5 = JsonHelper.loadDocument(Settings.Default.ProjectID, "ProjectOfficeCheckList");
            VersionControl<ProjectOfficeChecklistModel> versionControl4 = JsonConvert.DeserializeObject<VersionControl<ProjectOfficeChecklistModel>>(json5);
            ProjectOfficeChecklistModel currentProjectOfficeChecklist;

            if (versionControl4 != null)
            {
                currentProjectOfficeChecklist = JsonConvert.DeserializeObject<ProjectOfficeChecklistModel>(versionControl4.getLatest(versionControl4.DocumentModels));
                initDocsListStatus.Add(currentProjectOfficeChecklist.ProjectOfficeCheckListProgress);

            }
            else
                //    IsProjectOfficeChecklistDone = "";
                initDocsListStatus.Add("");


            //////PHASE REVIEW FORM INITIATION/////////
            string json6 = JsonHelper.loadDocument(Settings.Default.ProjectID, "PhaseReviewFormInitiation");
            VersionControl<PhaseReviewFormInitiationModel> versionControl5 = JsonConvert.DeserializeObject<VersionControl<PhaseReviewFormInitiationModel>>(json6);
            PhaseReviewFormInitiationModel currentPhaseReviewFormInitiation;

            if (versionControl5 != null)
            {
                currentPhaseReviewFormInitiation = JsonConvert.DeserializeObject<PhaseReviewFormInitiationModel>(versionControl5.getLatest(versionControl5.DocumentModels));
                initDocsListStatus.Add(currentPhaseReviewFormInitiation.PhaseReviewFormInitiationProgress);

            }
            else
                // IsPhaseReviewInitiationDone = "";
                initDocsListStatus.Add("");


            //////TERMS OF REFERENCE/////////
            string json7 = JsonHelper.loadDocument(Settings.Default.ProjectID, "TermOfReferenceDocument");
            VersionControl<TermsOfReferenceModel> versionControl6 = JsonConvert.DeserializeObject<VersionControl<TermsOfReferenceModel>>(json7);
            TermsOfReferenceModel currentTermOfReference;

            if (versionControl6 != null)
            {
                currentTermOfReference = JsonConvert.DeserializeObject<TermsOfReferenceModel>(versionControl6.getLatest(versionControl6.DocumentModels));
                initDocsListStatus.Add(currentTermOfReference.TermOfReferenceProgress);

            }
            else
                //  IsTermOfReferenceDone = "";
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

            //Get localdocs
            List<string> localDocuments = getLocalDocuments();

            lblProjectName.Text = projectModel.ProjectName;

            chart1.ChartAreas[0].BackColor = Color.Transparent;
            chart1.Legends[0].BackColor = Color.Transparent;
            chart2.Legends[0].BackColor = Color.Transparent;



            //Counters for completed, uncompleted, and in progress tasks
            int comp = 0, uncomp = 0, inprog = 0;
            int compPlanning = 0, uncompPlanning = 0, inprogPlanning = 0;
            int compExecution = 0, uncompExecution = 0, inprogExecution = 0;
            int compClosing = 0, uncompClosing = 0, inprogClosing = 0;

            if (localDocuments == null)
            {
                MessageBox.Show("No documents have been added yet.", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                ///////////////INITIATION PHASE/////////////////
                initiationDocuments.Add("BusinessCase");
                initiationDocuments.Add("FeasibilityStudy");
                initiationDocuments.Add("ProjectCharter");
                initiationDocuments.Add("JobDescription");
                initiationDocuments.Add("ProjectOfficeCheckList");
                initiationDocuments.Add("PhaseReviewFormInitiation");
                initiationDocuments.Add("TermOfReferenceDocument");

                initDocsListStatus.Add("BusinessCase");


                //lblInitiationProgress.Text = "Progress: 0%";
                //pbarInitiation.Value = 0;
                //pbarInitiation.Maximum = initiationDocuments.Count;


                for (int i = 0; i < initiationDocuments.Count; i++)
                {
                    initationProgressVal++;
                    dgvInitiation.Rows.Add();
                    dgvInitiation.Rows[i].Cells[0].Value = initiationDocuments[i];


                    if (initDocsListStatus[i] == "UNDONE")
                    {
                        dgvInitiation.Rows[i].Cells[1].Style.BackColor = Color.Orange;
                        inprog++;
                    }
                    else if (initDocsListStatus[i] == "DONE")
                    {
                        comp++;
                        dgvInitiation.Rows[i].Cells[1].Style.BackColor = Color.LimeGreen;
                        // pbarInitiation.Value = (int)initationProgressVal;
                        initationPercentage = ((initationProgressVal) / initiationDocuments.Count)*100;
                        //lblInitiationProgress.Text = "Progress: " + Math.Round(initationPercentage, 2) + "%";

                        xValues1[0] = "Initiation";
                        yValues1[0] = initationPercentage;

                        yValues2[0] = 100 - initationPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        uncomp++;
                        dgvInitiation.Rows[i].Cells[1].Style.BackColor = Color.Gray;
                    }


                }

                lblInitNumTasks.Text = (initationPercentage /100).ToString("p");


                chartInit.ChartAreas[0].BackColor = Color.Transparent;
                chartInit.Legends[0].BackColor = Color.Transparent;
                chartInit.Legends[0].BackColor = Color.Transparent;
                string[] xInit = { "Completed Tasks  " + comp, "Not started Tasks  " + uncomp, "In Progress Tasks " + inprog };

                double[] yInit = { comp, uncomp, inprog };

                chartInit.Series["Series1"].Points.DataBindXY(xInit, yInit);
                chartInit.Series["Series1"].ChartType = SeriesChartType.Doughnut;

                chartInit.Series["Series1"]["PieLabelStyle"] = "Disabled";
                chartInit.Legends[0].Enabled = true;

                chartInit.Text = "Test";

                chartInit.Series["Series1"].Points[0].Color = Color.LimeGreen;
                chartInit.Series["Series1"].Points[1].Color = Color.Gray;
                chartInit.Series["Series1"].Points[2].Color = Color.Orange;
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
                    planningProgressVal++;
                    dgvPlanning.Rows.Add();
                    dgvPlanning.Rows[i].Cells[0].Value = planningDocuments[i];

                    if (planningDocsListStatus[i] == "UNDONE")
                    {
                        dgvPlanning.Rows[i].Cells[1].Style.BackColor = Color.Orange;
                        inprogPlanning++;
                    }
                    else if (planningDocsListStatus[i] == "DONE")
                    {
                        compPlanning++;
                        dgvPlanning.Rows[i].Cells[1].Style.BackColor = Color.LimeGreen;
                        planningPercentage = ((planningProgressVal) / planningDocuments.Count) * 100;

                        xValues1[1] = "Planning";
                        yValues1[1] = planningPercentage;

                        yValues2[1] = 100 - planningPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        uncompPlanning++;
                        dgvPlanning.Rows[i].Cells[1].Style.BackColor = Color.Gray;
                    }
                }

                lblPlanNumTasks.Text = (planningPercentage / 100).ToString("p");


                chartPlanning.ChartAreas[0].BackColor = Color.Transparent;
                chartPlanning.Legends[0].BackColor = Color.Transparent;
                chartPlanning.Legends[0].BackColor = Color.Transparent;
                string[] xPlan = { "Completed Tasks  " + compPlanning, "Not started Tasks  " + uncompPlanning, "In Progress Tasks " + inprogPlanning };

                double[] yPlan = { compPlanning, uncompPlanning, inprogPlanning };

                chartPlanning.Series["Series1"].Points.DataBindXY(xPlan, yPlan);
                chartPlanning.Series["Series1"].ChartType = SeriesChartType.Doughnut;

                chartPlanning.Series["Series1"]["PieLabelStyle"] = "Disabled";
                chartPlanning.Legends[0].Enabled = true;

                chartPlanning.Text = "Test";

                chartPlanning.Series["Series1"].Points[0].Color = Color.LimeGreen;
                chartPlanning.Series["Series1"].Points[1].Color = Color.Gray;
                chartPlanning.Series["Series1"].Points[2].Color = Color.Orange;
                ///////////////////////////////////////////////////////////////////////////////////////////////////

                ///////////////EXECUTION PHASE/////////////////
                //executionDocuments.Add("BuildDeliverables");
                //executionDocuments.Add("MonitorAndControl");
                executionDocuments.Add("TimeMangement");
                executionDocuments.Add("TimeSheet");
                executionDocuments.Add("TimeSheetRegister");
                executionDocuments.Add("CostManagementProcess");
                executionDocuments.Add("ExpenseForm");
                executionDocuments.Add("ExpenseRegister");


                ///////////////////////////////////////////////////////////////////////
                ///Sal jy dalk net ook kyk na die Quality goed, daar is 4 goed hier maar net 3 goed in die mainform onder quality

                executionDocuments.Add("QualityManagement"); 
                executionDocuments.Add("QualityReviewPlan"); 
                executionDocuments.Add("QualityReviewForm"); 
                //executionDocuments.Add("QualityReviewRegister"); 
                ////////////////////////////////////////////////////////////////////////


                executionDocuments.Add("ChangeManagementProcess");
                executionDocuments.Add("ChangeRequestForm");
                executionDocuments.Add("ChangeRequestRegister");
                executionDocuments.Add("RiskManagamentProcess");
                executionDocuments.Add("RiskForm");
                //executionDocuments.Add("RiskRegister"); Al die moet nog in kom Rickus :)
                //executionDocuments.Add("IssueManagementProcess");
                //executionDocuments.Add("IssueForm");
                //executionDocuments.Add("IssueRegister");
                //executionDocuments.Add("PurchaseOrder");
                //executionDocuments.Add("ProcurementRegister");
                //executionDocuments.Add("AcceptanceManagementProcess");
                //executionDocuments.Add("AcceptanceForm");
                //executionDocuments.Add("AcceptanceRegister");
                //executionDocuments.Add("CommunicationsManagementProcess");
                //executionDocuments.Add("ProjectStatusReport");
                //executionDocuments.Add("CommunicationsRegister");
                //executionDocuments.Add("PhaseReviewExe");

                // executionDocsListStatus.Add("TimeMangement");

                //lblExecutionProgress.Text = "Progress: 0%";
                //pbarExecution.Value = 0;
                //pbarExecution.Maximum = executionDocuments.Count;

                //for (int i = 0; i < executionDocuments.Count; i++)
                //{
                //    dgvExecution.Rows.Add();
                //    dgvExecution.Rows[i].Cells[0].Value = executionDocuments[i];
                //    if (localDocuments.Contains(executionDocuments[i]))
                //    {
                //        executionProgressVal++;
                //        dgvExecution.Rows[i].Cells[1].Value = true;
                //        //pbarExecution.Value = (int)executionProgressVal;
                //        executionPercentage = ((executionProgressVal) / executionDocuments.Count) * 100;
                //        //lblExecutionProgress.Text = "Progress: " + Math.Round(executionPercentage, 2) + "%";

                //        xValues1[2] = "Execution";
                //        yValues1[2] = executionPercentage;

                //        yValues2[2] = 100 - executionPercentage;

                //        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                //        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                //    }
                //    else
                //    {
                //        dgvExecution.Rows[i].Cells[1].Value = false;


                //        xValues1[3] = "Execution";
                //        yValues1[3] = 0;

                //        yValues2[3] = 100;

                //        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                //        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                //    }
                //}

                for (int i = 0; i < executionDocuments.Count; i++)
                {
                    
                    dgvExecution.Rows.Add();
                    dgvExecution.Rows[i].Cells[0].Value = executionDocuments[i];

                    if (executionDocsListStatus[i] == "UNDONE")
                    {
                        dgvExecution.Rows[i].Cells[1].Style.BackColor = Color.Orange;
                        inprogExecution++;
                    }
                    else if (executionDocsListStatus[i] == "DONE")
                    {
                        executionProgressVal++;
                        compExecution++;
                        dgvExecution.Rows[i].Cells[1].Style.BackColor = Color.LimeGreen;
                        executionPercentage = ((executionProgressVal) / executionDocuments.Count) * 100;

                        xValues1[1] = "Execution";
                        yValues1[1] = executionPercentage;

                        yValues2[1] = 100 - executionPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        uncompExecution++;
                        dgvExecution.Rows[i].Cells[1].Style.BackColor = Color.Gray;
                    }
                }

                lblExecNumTasks.Text = (executionPercentage / 100).ToString("p");


                chartExecution.ChartAreas[0].BackColor = Color.Transparent;
                chartExecution.Legends[0].BackColor = Color.Transparent;
                chartExecution.Legends[0].BackColor = Color.Transparent;
                string[] xExec = { "Completed Tasks  " + compExecution, "Not started Tasks  " + uncompExecution, "In Progress Tasks " + inprogExecution };

                double[] yExec = { compExecution, uncompExecution, inprogExecution };

                chartExecution.Series["Series1"].Points.DataBindXY(xExec, yExec);
                chartExecution.Series["Series1"].ChartType = SeriesChartType.Doughnut;

                chartExecution.Series["Series1"]["PieLabelStyle"] = "Disabled";
                chartExecution.Legends[0].Enabled = true;

                chartExecution.Text = "Test";

                chartExecution.Series["Series1"].Points[0].Color = Color.LimeGreen;
                chartExecution.Series["Series1"].Points[1].Color = Color.Gray;
                chartExecution.Series["Series1"].Points[2].Color = Color.Orange;


                closingDocuments.Add("ProjectClosureReport");
                closingDocuments.Add("PostImplementationReview");

                lblClosingProgress.Text = "Progress: 0%";
                pbarClosing.Value = 0;
                pbarClosing.Maximum = closingDocuments.Count;
                for (int i = 0; i < closingDocuments.Count; i++)
                {
                    dgvClosing.Rows.Add();
                    dgvClosing.Rows[i].Cells[0].Value = closingDocuments[i];
                    if (localDocuments.Contains(closingDocuments[i]))
                    {
                        closingProgressVal++;
                        dgvClosing.Rows[i].Cells[1].Value = true;
                        pbarClosing.Value = (int)closingProgressVal;
                        closingPercentage = ((closingProgressVal) / closingDocuments.Count) * 100;
                        lblClosingProgress.Text = "Progress: " + Math.Round(closingPercentage, 2) + "%";

                        xValues1[3] = "Closing";
                        yValues1[3] = closingPercentage;

                        yValues2[3] = 100 - closingPercentage;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
                    }
                    else
                    {
                        dgvClosing.Rows[i].Cells[1].Value = false;

                        xValues1[3] = "Closing";
                        yValues1[3] = 0;

                        yValues2[3] = 100;

                        chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
                        chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);

                    }
                }

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

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    initationProgressVal = 0;
        //    for (int i = 0; i < pbarInitiation.Maximum; i++) // Go through documents
        //    {
        //        initationProgressVal++; //Increase progress
        //        dgvInitiation.Rows[i].Cells[1].Value = true; // check each unchecked checkbox

        //    }
        //    pbarInitiation.Value = pbarInitiation.Maximum;
        //    lblInitiationProgress.Text = "Progress: 100%";

        //    //////////////////// Calculate overall progressbar percentage //////////////
        //    double overallProgressVal = (initationProgressVal + planningProgressVal + executionProgressVal + closingProgressVal);
        //    pbarOverall.Value = (int)overallProgressVal;
        //    pbarOverall.Maximum = initiationDocuments.Count + planningDocuments.Count + executionDocuments.Count + closingDocuments.Count;
        //    double overallPercentage = ((overallProgressVal) / pbarOverall.Maximum) * 100;
        //    lblOverallProgress.Text = "Overall Progress: " + Math.Round(overallPercentage, 2) + "%";
        //    ////////////////////////////////////////////////////////////////////////////
        //    string[] xValues = { "Completed Tasks", "Not Started Tasks" };
        //    double[] yValues = { overallPercentage, 100 - overallPercentage };

        //    chart1.Series["Series1"].Points.DataBindXY(xValues, yValues);
        //    chart1.Series["Series1"].ChartType = SeriesChartType.Doughnut;

        //    chart1.Series["Series1"]["PieLabelStyle"] = "Disabled";
        //    chart1.Legends[0].Enabled = true;

        //    chart1.Series["Series1"].Points[0].Color = Color.LimeGreen;
        //    chart1.Series["Series1"].Points[1].Color = Color.Gray;

        //    foreach (DataPoint p in chart1.Series["Series1"].Points)
        //    {
        //        p.Label = "#PERCENT\n#VALX";
        //    }

        //    initationPercentage = ((initationProgressVal) / initiationDocuments.Count) * 100;

        //    xValues1[0] = "Initiation";
        //    //yValues1[0] = 0;

        //    yValues2[0] = 100 - initationPercentage;

        //    chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
        //    chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
        //}

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    planningProgressVal = 0;
        //    for (int i = 0; i < pbarPlanning.Maximum; i++) // Go through documents
        //    {
        //        planningProgressVal++; //Increase progress
        //        dgvPlanning.Rows[i].Cells[1].Value = true; // check each unchecked checkbox
        //    }

        //    pbarPlanning.Value = pbarPlanning.Maximum;
        //    lblPlanningProgress.Text = "Progress: 100%";

        //    //////////////////// Calculate overall progressbar percentage //////////////
        //    double overallProgressVal = (initationProgressVal + planningProgressVal + executionProgressVal + closingProgressVal);
        //    pbarOverall.Value = (int)overallProgressVal;
        //    pbarOverall.Maximum = initiationDocuments.Count + planningDocuments.Count + executionDocuments.Count + closingDocuments.Count;
        //    double overallPercentage = ((overallProgressVal) / pbarOverall.Maximum) * 100;
        //    lblOverallProgress.Text = "Overall Progress: " + Math.Round(overallPercentage, 2) + "%";
        //    ////////////////////////////////////////////////////////////////////////////
        //    string[] xValues = { "Completed Tasks", "Not Started Tasks" };
        //    double[] yValues = { overallPercentage, 100 - overallPercentage };

        //    chart1.Series["Series1"].Points.DataBindXY(xValues, yValues);
        //    chart1.Series["Series1"].ChartType = SeriesChartType.Doughnut;

        //    chart1.Series["Series1"]["PieLabelStyle"] = "Disabled";
        //    chart1.Legends[0].Enabled = true;

        //    chart1.Series["Series1"].Points[0].Color = Color.LimeGreen;
        //    chart1.Series["Series1"].Points[1].Color = Color.Gray;


        //    foreach (DataPoint p in chart1.Series["Series1"].Points)
        //    {
        //        p.Label = "#PERCENT\n#VALX";
        //    }

        //    planningPercentage = ((planningProgressVal) / planningDocuments.Count) * 100;

        //    yValues1[1] = 100;
        //    yValues2[1] = 100 - planningPercentage;

        //    chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
        //    chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
        //}

        //private void button3_Click(object sender, EventArgs e)
        //{
        //    executionProgressVal = 0;
        //    for (int i = 0; i < pbarExecution.Maximum; i++) // Go through documents
        //    {
        //        executionProgressVal++; //Increase progress
        //        dgvExecution.Rows[i].Cells[1].Value = true; // check each unchecked checkbox
        //    }
        //    pbarExecution.Value = pbarExecution.Maximum;
        //    lblExecutionProgress.Text = "Progress: 100%";

        //    //////////////////// Calculate overall progressbar percentage //////////////
        //    double overallProgressVal = (initationProgressVal + planningProgressVal + executionProgressVal + closingProgressVal);
        //    pbarOverall.Value = (int)overallProgressVal;
        //    pbarOverall.Maximum = initiationDocuments.Count + planningDocuments.Count + executionDocuments.Count + closingDocuments.Count;
        //    double overallPercentage = ((overallProgressVal) / pbarOverall.Maximum) * 100;
        //    lblOverallProgress.Text = "Overall Progress: " + Math.Round(overallPercentage, 2) + "%";
        //    ////////////////////////////////////////////////////////////////////////////
        //    ///
        //    string[] xValues = { "Completed Tasks", "Not Started Tasks" };
        //    double[] yValues = { overallPercentage, 100 - overallPercentage };

        //    chart1.Series["Series1"].Points.DataBindXY(xValues, yValues);
        //    chart1.Series["Series1"].ChartType = SeriesChartType.Doughnut;

        //    chart1.Series["Series1"]["PieLabelStyle"] = "Disabled";
        //    chart1.Legends[0].Enabled = true;

        //    chart1.Series["Series1"].Points[0].Color = Color.LimeGreen;
        //    chart1.Series["Series1"].Points[1].Color = Color.Gray;


        //    foreach (DataPoint p in chart1.Series["Series1"].Points)
        //    {
        //        p.Label = "#PERCENT\n#VALX";
        //    }

        //    executionPercentage = ((executionProgressVal) / executionDocuments.Count) * 100;

        //    yValues1[2] = 100;
        //    yValues2[2] = 100 - executionPercentage;

        //    chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
        //    chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
        //}

        private void button4_Click(object sender, EventArgs e)
        {
            closingProgressVal = 0;
            for (int i = 0; i < pbarClosing.Maximum; i++) // Go through documents
            {
                closingProgressVal++; //Increase progress
                dgvClosing.Rows[i].Cells[1].Value = true; // check each unchecked checkbox
            }
            pbarClosing.Value = pbarClosing.Maximum;
            lblClosingProgress.Text = "Progress: 100%";

            //////////////////// Calculate overall progressbar percentage //////////////
            double overallProgressVal = (initationProgressVal + planningProgressVal + executionProgressVal + closingProgressVal);
            pbarOverall.Value = (int)overallProgressVal;
            pbarOverall.Maximum = initiationDocuments.Count + planningDocuments.Count + executionDocuments.Count + closingDocuments.Count;
            double overallPercentage = ((overallProgressVal) / pbarOverall.Maximum) * 100;
            lblOverallProgress.Text = "Overall Progress: " + Math.Round(overallPercentage, 2) + "%";
            ////////////////////////////////////////////////////////////////////////////
            ///
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

            closingPercentage = ((closingProgressVal) / closingDocuments.Count) * 100;

            yValues1[3] = 100;
            yValues2[3] = 100 - closingPercentage;

            chart2.Series["Completed"].Points.DataBindXY(xValues1, yValues1);
            chart2.Series["Not Started"].Points.DataBindXY(xValues1, yValues2);
        }

        private void chart3_Click(object sender, EventArgs e)
        {

        }
    }
}


