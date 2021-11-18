using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.OleDb;
using Microsoft.Office.Interop;
using ProjectManagementToolkit.MPMM.MPMM_Document_Forms;

namespace ProjectManagementToolkit.MPMM.MPMM_Forms.Project_Management
{
    public partial class frmPLSM_Workpackages : Form
    {
        public frmPLSM_Workpackages()
        {
            InitializeComponent();
        }

        private void btnBackToPLMSFrontEnd_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void btnWorkpackageRisk_Click(object sender, EventArgs e)
        {
            
        }

        private void frmPLSM_Workpackages_Load(object sender, EventArgs e)
        {
             
        }

        private void btnOpportunityScrnIdentifyOppForDev_Click(object sender, EventArgs e)
        {
            frmOpportunityScreeningWorkpackages frmOppScrnWork = new frmOpportunityScreeningWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen Opportunities for strategic fit'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmOppScrnWork.OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmOppScrnWork.Show();
        }

        private void btnOppScrnIdentifyNeedsOfStakeholders_Click(object sender, EventArgs e)
        {
            frmOpportunityScreeningWorkpackages frmOppScrnWork = new frmOpportunityScreeningWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Identify Needs of Stakeholders'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmOppScrnWork.OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmOppScrnWork.Show();
        }

        private void btnOppScrnScreenOppForStratFit_Click(object sender, EventArgs e)
        {
            frmOpportunityScreeningWorkpackages frmOppScrnWork = new frmOpportunityScreeningWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen Opportunities for strategic fit'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmOppScrnWork. OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmOppScrnWork.Show();
        }

        private void btnOppScrnReviewCollectSuppData_Click(object sender, EventArgs e)
        {
            frmOpportunityScreeningWorkpackages frmOppScrnWork = new frmOpportunityScreeningWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Review and collect supporting Data'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmOppScrnWork.OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmOppScrnWork.Show();
        }

        private void btnOppScrnDevPreFeasabilityPlan_Click(object sender, EventArgs e)
        {
            frmOpportunityScreeningWorkpackages frmOppScrnWork = new frmOpportunityScreeningWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop Pre-Feasibility Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmOppScrnWork.OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmOppScrnWork.Show();
        }

        private void btnOppScrnPriorOppApplyingScrTool_Click(object sender, EventArgs e)
        {
            frmOpportunityScreeningWorkpackages frmOppScrnWork = new frmOpportunityScreeningWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Prioritise Opportunities applying scoring tool'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmOppScrnWork.OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmOppScrnWork.Show();
        }

        private void btnOppScrnSubmitOppToInvestPortfolio_Click(object sender, EventArgs e)
        {
            frmOpportunityScreeningWorkpackages frmOppScrnWork = new frmOpportunityScreeningWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Submit Opportunities to Investment Portfolio Process'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmOppScrnWork.OpportunityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmOppScrnWork.Show();
        }

        private void btnConcludeStage_Click(object sender, EventArgs e)
        {
            frmOpportunityScreeningWorkpackages frmOppScrnWork = new frmOpportunityScreeningWorkpackages();

            frmOppScrnWork.Show();
        }

        private void WorkPackagesBuildTestTabPage_Click(object sender, EventArgs e)
        {

        }

        private void btnPreFeasProjectDef_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Definition : Project Brief'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnPreFeasEstTheCoreProjectTeam_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish the Core Project Team.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnPreFeasInitiateHighLevelProjectDev_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Initiate high-level project development planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnPreFeasAppointProjectTeam_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Appoint the project team:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnPreFeasPerformConceptDesign_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform Concept Design'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnPreFeasPerformBaseInfra_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform base infrastructure assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnPreFeasDetermineLegalAndRegulatoryReq_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Determine Legal and Regulatory requirements'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnPreFeasPerformStakeholderNeeds_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform stakeholder needs impact assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnPreFeasCommunicationAndStakeholder_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();

            frmPreFeasWork.Show();
        }

        private void btnPreFeasPartnerSelection_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Partner Selection'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnIdFinAndCommercialStruct_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Identify financial and commercial structures'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnPreFeasConstructPreliminaryBussinessCase_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Construct a preliminary Business Case and Risk Assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnPreFeasScreenPriorAndDiscardOpp_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen, prioritise and discard opportunities'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPreFeasWork.PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPreFeasWork.Show();
        }

        private void btnPreFeasConludeStage_Click(object sender, EventArgs e)
        {
            frmPreFeasabilityWorkpackages frmPreFeasWork = new frmPreFeasabilityWorkpackages();

            frmPreFeasWork.Show();
        }

        private void btnFeasProjectStartUpDef_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Project Start Up Definition%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasDevelopProjectCharter_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop Project Charter:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasStructureDCO_PMOContr_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Structure DCO & PMO Contract'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasEstProjectManagementSupp_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish Project Management Support'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasEstPMOPartyStructures_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish PMO & 3rd Party Structures'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasEstClientOfficeManagementSupp_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish Client Office Management Support'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasAssessingLegalReq_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();

            frmFeasWork.Show();
        }

        private void btnFeasCompileBusinessCase_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasPerformBasicDesign_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform Basic Design'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasPreliminarySiteSelection_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Preliminary Site Selection'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasEnvironmentalImpactAssessment_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();

            frmFeasWork.Show();
        }

        private void btnFeasFinaliseSiteSelection_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();

            frmFeasWork.Show();
        }

        private void btnFeasDevelopContractingStrat_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop Contracting Strategy'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasConcludeStage_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();

            frmFeasWork.Show();
        }

        private void btnFeasExecuteEIA_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Execute EIA, Regulatory and Legal'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasProduceURS_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Produce (URS) User Requirement Specs.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasUndertakeFeasStudy_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Undertake a feasabillity study:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnFeasDealStructuringAndBankableReport_Click(object sender, EventArgs e)
        {
            frmFeasabilityWorkpackages frmFeasWork = new frmFeasabilityWorkpackages();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Deal Structuring and Bankable Report'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmFeasWork.FeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmFeasWork.Show();
        }

        private void btnBCInitiateStructProject_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Initiate and Structure Project Set-Up'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void btnBCFinaliseRegulatoryLegalAppr_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Finalise Regulatory and Legal Approvals'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void btnBCDevEngineeringSpecs_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop Engineering Specifications'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void btnBCScopeFreeze_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Scope Freeze'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void btnBCProjectFinManagement_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Financial management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void btnBCResolveCommercialFinancialIssues_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Resolve Commercial and Financial Issues'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void ProjectProcurementManagement_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Procurement management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void btnBCPrepareDistributeRFI_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Prepare and Distribute RFI%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void btnBCProjectExecutionStrat_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Execution Strategy'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void btnBCDetailBusinessPlanDeveloped_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Detailed Business Plan Developed'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void btnBCLifecycleOperationPlan_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Lifecycle Operation Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void btnBCDataHandOverTExecution_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Data Hand-Over to Execution'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBP.BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBP.Show();
        }

        private void btnBCConcludeStage_Click(object sender, EventArgs e)
        {
            frmBusinessPlan frmBP = new frmBusinessPlan();

            frmBP.Show();
        }

        private void btnProjectExeProjectStartUp_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Start-Up'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeEstProjectExe_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish Project Execution Org. Structure'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeCreateExePlans_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a procurement plan%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeCreateExePlansRiskPlan_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a risk plan%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeCreateExePlansScopeWBS_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans : Scope, WBS and resources Planning%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeCreateExePlansFinPlan_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a financial plan%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeCreateExePlansQualityPlan_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a quality plan%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeCreateExePlansScheduleProject_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans : Schedule Project Work'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeCreateExePlansCommunicationPlan_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a Communications plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeCreateExePlansChangeManagementPlan_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a change management plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeCompileBalanceEnquiries_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Balance of Enquiries'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeConfigManagement_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Configuration management Implement configuration Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeCreateExePlansAcceptancePlan_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create an acceptance plan%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeDesignFreeze_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Design Freeze'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeConceptDev_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Concept Development%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeUpdateProjectImplementationPlan_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Update Project Implementation Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void ProjectImplementationPlan_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Implementation Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPEP.ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPEP.Show();
        }

        private void btnProjectExeConcludeStage_Click(object sender, EventArgs e)
        {
            frmProjectExecutionPlanning frmPEP = new frmProjectExecutionPlanning();

            frmPEP.Show();
        }

        private void btnContractConcCompileBalanceOfEnquiries_Click(object sender, EventArgs e)
        {
            frmContractConclusion frmCC = new frmContractConclusion();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Balance of Enquiries'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCC.ContractConcDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCC.Show();
        }

        private void btnContractConcUpdateProjectImplementationPlan_Click(object sender, EventArgs e)
        {
            frmContractConclusion frmCC = new frmContractConclusion();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Update Project Implementation Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCC.ContractConcDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCC.Show();
        }

        private void btnContractConcConcludeStage_Click(object sender, EventArgs e)
        {
            frmContractConclusion frmCC = new frmContractConclusion();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Conclude Stage%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCC.ContractConcDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCC.Show();
        }

        private void btnDetailDesignDevDeployment_Click(object sender, EventArgs e)
        {
            frmDetailedDesign frmDD = new frmDetailedDesign();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Developing deployment and operational concepts'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmDD.DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmDD.Show();
        }

        private void btnDetailDesignConcludeStage_Click(object sender, EventArgs e)
        {
            frmDetailedDesign frmDD = new frmDetailedDesign();

            frmDD.Show();
        }

        private void btnDetailDesignBackUpData_Click(object sender, EventArgs e)
        {
            frmDetailedDesign frmDD = new frmDetailedDesign();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Back up data'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmDD.DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmDD.Show();
            frmDD.btnDetailDesignWBackUpData.Focus();
        }

        private void btnDetailDesignPrepareProcesses_Click(object sender, EventArgs e)
        {
            frmDetailedDesign frmDD = new frmDetailedDesign();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Prepare processes and organization'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmDD.DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmDD.Show();
        }

        private void btnDetailDesignActivateProcesses_Click(object sender, EventArgs e)
        {
            frmDetailedDesign frmDD = new frmDetailedDesign();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Activate processes and organization'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmDD.DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmDD.Show();
            
        }

        private void btnDetailDesignImplementProcessOrgConcept_Click(object sender, EventArgs e)
        {
            frmDetailedDesign frmDD = new frmDetailedDesign();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Implement process and organizational concept%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmDD.DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmDD.Show();
        }

        private void btnDetailDesignDevProcessOrganizational_Click(object sender, EventArgs e)
        {
            frmDetailedDesign frmDD = new frmDetailedDesign();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop process and organizational concept%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmDD.DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmDD.Show();
        }

        private void btnDetailDesignDeployPlanning_Click(object sender, EventArgs e)
        {
            frmDetailedDesign frmDD = new frmDetailedDesign();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Deployment Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmDD.DetailedDesignDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmDD.Show();
        }

        private void btnBuildTestBuildExecute_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Build  Execute'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestConductTest_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Conduct tests%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestAcceptingDeliverables_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Accepting Deliverables%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestDeploySystem_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Deploy system%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestDeployPlanning_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Deployment Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestDevelopTrainingConcept_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop training concept'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestDraftTrainingDoc_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Draft training documentation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestTrainUsers_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Train users'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestImplementProtectionMeasures_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Implement protection measures'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestStartUpAndCommissioning_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Start-Up and Commissioning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestHandOver_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Hand Over / Partial Hand Over'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestTransferSollution_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Transfer sollution%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestConsolidateDoc_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Consolidate Documentation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBandT.BuildAndTestDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBandT.Show();
        }

        private void btnBuildTestConcludeStage_Click(object sender, EventArgs e)
        {
            frmBuildAndTest frmBandT = new frmBuildAndTest();

            frmBandT.Show();
        }

        private void btnImpSitePrepAndAccess_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Site Preparation and access'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpEstSiteSuppInfra_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();

            frmImp.Show();
        }

        private void btnImpBuildConstructDeliverable_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Build Construct%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpManageDetailDesign_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Product delivery management%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpProductDeliveryManagement_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Product delivery management%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpConductTests_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Conduct tests%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpAcceptingDeliverables_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Accepting Deliverables%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpEstOperationalReadiness_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish Operational Readiness'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpDeploymentPlanning_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Deployment Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpHandOver_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Hand Over / Partial Hand Over'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpDeploySystem_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Deploy system%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpStartUpAndCommissioning_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Start-Up and Commissioning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpTransferSollution_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Transfer sollution%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmImp.ImplementationDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmImp.Show();
        }

        private void btnImpConcludeStage_Click(object sender, EventArgs e)
        {
            frmImplementation frmImp = new frmImplementation();

            frmImp.Show();
        }

        private void btnTransferHandOver_Click(object sender, EventArgs e)
        {
            frmTransferWorkpackage frmTFW = new frmTransferWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Hand Over / Partial Hand Over'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmTFW.TransferDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmTFW.Show();
        }

        private void btnTransferDeploymentPlanning_Click(object sender, EventArgs e)
        {
            frmTransferWorkpackage frmTFW = new frmTransferWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Deployment Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmTFW.TransferDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmTFW.Show();
        }

        private void btnTransferStartUpCommissioning_Click(object sender, EventArgs e)
        {
            frmTransferWorkpackage frmTFW = new frmTransferWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Start-Up and Commissioning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmTFW.TransferDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmTFW.Show();
        }

        private void btnTransferDeploySystem_Click(object sender, EventArgs e)
        {
            frmTransferWorkpackage frmTFW = new frmTransferWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Deploy system%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmTFW.TransferDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmTFW.Show();
        }

        private void btnTransferTransferSollution_Click(object sender, EventArgs e)
        {
            frmTransferWorkpackage frmTFW = new frmTransferWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Transfer sollution%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmTFW.TransferDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmTFW.Show();
        }

        private void btnTransferConductTests_Click(object sender, EventArgs e)
        {
            frmTransferWorkpackage frmTFW = new frmTransferWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Conduct tests%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmTFW.TransferDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmTFW.Show();
        }

        private void btnTransferConsolidateDoc_Click(object sender, EventArgs e)
        {
            frmTransferWorkpackage frmTFW = new frmTransferWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Consolidate Documentation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmTFW.TransferDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmTFW.Show();
        }

        private void btnCloseOutCloseOutProject_Click(object sender, EventArgs e)
        {
            frmCloseOutWorkpackage frmCO = new frmCloseOutWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Close-Out Project'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCO.CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCO.Show();
        }

        private void btnCloseOutInitiatePlantLifecycle_Click(object sender, EventArgs e)
        {
            frmCloseOutWorkpackage frmCO = new frmCloseOutWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Initiate Plant Lifecycle Plan - O&M'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCO.CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCO.Show();
        }

        private void btnCloseOutConcludeQualityAssurance_Click(object sender, EventArgs e)
        {
            frmCloseOutWorkpackage frmCO = new frmCloseOutWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude quality assurance'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCO.CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCO.Show();
        }

        private void btnCloseOutConcludeProjectMarketing_Click(object sender, EventArgs e)
        {
            frmCloseOutWorkpackage frmCO = new frmCloseOutWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude project marketing'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCO.CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCO.Show();
        }

        private void btnCloseOutConcludeBusinessProcessModeling_Click(object sender, EventArgs e)
        {
            frmCloseOutWorkpackage frmCO = new frmCloseOutWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude business process modeling'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCO.CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCO.Show();
        }

        private void btnCloseOutConcludeRiskManagement_Click(object sender, EventArgs e)
        {
            frmCloseOutWorkpackage frmCO = new frmCloseOutWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude risk management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCO.CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCO.Show();
        }

        private void btnCloseOutConcludeInfoSecurityAndDataProtection_Click(object sender, EventArgs e)
        {
            frmCloseOutWorkpackage frmCO = new frmCloseOutWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude information security and data protection'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCO.CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCO.Show();
        }

        private void btnCloseOutConcludeConfigManagement_Click(object sender, EventArgs e)
        {
            frmCloseOutWorkpackage frmCO = new frmCloseOutWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conclude configuration management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCO.CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCO.Show();
        }

        private void btnCloseOutDeCommissioningProject_Click(object sender, EventArgs e)
        {
            frmCloseOutWorkpackage frmCO = new frmCloseOutWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%De-Commissioning A Project%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmCO.CloseOutDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmCO.Show();
        }

        private void btnEvalConfirmProjectComp_Click(object sender, EventArgs e)
        {
            frmEvaluateWorkpackage frmE = new frmEvaluateWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Confirm Project Completion'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmE.EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmE.Show();
        }

        private void btnEvalDeCommissionProject_Click(object sender, EventArgs e)
        {
            frmEvaluateWorkpackage frmE = new frmEvaluateWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%De-Commissioning A Project%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmE.EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmE.Show();
        }

        private void btnEvalEvalProjectGov_Click(object sender, EventArgs e)
        {
            frmEvaluateWorkpackage frmE = new frmEvaluateWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Evaluate Project Governance and Operational Delivery.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmE.EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmE.Show();
        }

        private void btnEvalIdClosureAction_Click(object sender, EventArgs e)
        {
            frmEvaluateWorkpackage frmE = new frmEvaluateWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Identify Closure Actions%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmE.EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmE.Show();
        }

        private void btnEvalEvalTechDel_Click(object sender, EventArgs e)
        {
            frmEvaluateWorkpackage frmE = new frmEvaluateWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Evaluate Technical Delivery.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmE.EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmE.Show();
        }

        private void btnEvalUndertakeClosureActions_Click(object sender, EventArgs e)
        {
            frmEvaluateWorkpackage frmE = new frmEvaluateWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Underatke Closure Actions%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmE.EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmE.Show();
        }

        private void btnEvalAuditBusinessPlan_Click(object sender, EventArgs e)
        {
            frmEvaluateWorkpackage frmE = new frmEvaluateWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit Business Plan Benefits'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmE.EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmE.Show();
        }

        private void btnEvalIdFollowOnActions_Click(object sender, EventArgs e)
        {
            frmEvaluateWorkpackage frmE = new frmEvaluateWorkpackage();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Identifying Follow-On Actions'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmE.EvaluateDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmE.Show();
        }
    }
}
