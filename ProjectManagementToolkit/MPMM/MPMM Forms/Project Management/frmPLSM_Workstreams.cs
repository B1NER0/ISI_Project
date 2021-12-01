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

namespace ProjectManagementToolkit.MPMM.MPMM_Forms.Project_Management
{
    public partial class frmPLSM_Workstreams : Form
    {
        public frmPLSM_Workstreams()
        {
            InitializeComponent();
        }

        private void btnWorkstreamDirectingAProject_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Directing a project%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamScreenOppProblems_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Screen, prioritise and discard opportunities%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamFundsAcquisitionProcess_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Submit Opportunities to Investment Portfolio Process'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamIdBusinessRisks_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop a Business Case:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamPriorOpportunitiesProblems_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Prioritise Opportunities applying scoring tool'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamDefineHighLevelBusinessBenf_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop a Business Case:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
            
        }

        private void btnWorkstreamRegisterOpportunitiesProb_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen, prioritise and discard opportunities'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamConfirmBusinessBenefits_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit business plan benefit realisation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamConductBenefitRealisation_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();

            frmPMStream.Show();
        }

        private void btnWorkstreamMonitorBenefitRealisation_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();

            frmPMStream.Show();
        }

        private void btnWorkstreamMonitorGovernanceCompliance_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Evaluate Project Governance and Operational Delivery.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamAuditBenefitRealisation_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit Business Plan Benefits'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMStream.PortfolioManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMStream.Show();
        }

        private void btnWorkstreamProcessStageReport_Click(object sender, EventArgs e)
        {
            frmPortfolioManagementStream frmPMStream = new frmPortfolioManagementStream();

            frmPMStream.Show();
        }

        //ENGINEERING STREAM
        private void btnIdentifyOpp_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen Opportunities for strategic fit'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnWorkReq_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();

            frmEngineeringStream.Show();
        }

        private void btnStudyBusiness_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen Opportunities for strategic fit'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnAnalyseDifferent_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnEstimateCost_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop a Business Case:'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnHighlevel_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform Basic Design'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnConsiderAlt_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop a Business Case:'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnValidateTechnical_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop a Business Case:'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnAddressLegal_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Execute EIA, Regulatory and Legal'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnNextLevel_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop Engineering Specifications'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnTechinicalRecom_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Procurement management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnDesignComponent_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Product delivery management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnDesignInterface_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Configuration management Implement configuration Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnReviewDesignSpec_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Design Freeze'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnPrepareTest_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create an acceptance plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnPrepareBuild_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop prototype'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnDevelopConfig_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Configuration management Implement configuration Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnConfigInterfaces_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Configuration management Implement configuration Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnUnitInterface_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conduct tests'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnDevelopManuals_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop training concept'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnPrepareProduction_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();

            frmEngineeringStream.Show();
        }

        private void btnTrainOperators_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop training concept'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnDeployAsset_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Start-Up and Commissioning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnOwner_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Hand Over / Partial Hand Over'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnCommision_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Start-Up and Commissioning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnCheckGuarantee_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Initiate Plant Lifecycle Plan - O&M'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmEngineeringStream.EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmEngineeringStream.Show();
        }

        private void btnArchiveOutputs_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();

            frmEngineeringStream.Show();
        }

        private void btnBestPractices_Click(object sender, EventArgs e)
        {
            frmEngineeringStream frmEngineeringStream = new frmEngineeringStream();

            frmEngineeringStream.Show();
        }
        
        //ARCHITECTURE STREAM
        private void btnPreparePrefeasiblity_Click(object sender, EventArgs e)
        {
            frmArchitectureStream frmArchitectureStream = new frmArchitectureStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform base infrastructure assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmArchitectureStream.ArchitectureStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmArchitectureStream.Show();
        }

        private void btnPrepareConceptual_Click(object sender, EventArgs e)
        {
            frmArchitectureStream frmArchitectureStream = new frmArchitectureStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform base infrastructure assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmArchitectureStream.ArchitectureStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmArchitectureStream.Show();
        }

        private void btnPrepareDesign_Click(object sender, EventArgs e)
        {
            frmArchitectureStream frmArchitectureStream = new frmArchitectureStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform base infrastructure assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmArchitectureStream.ArchitectureStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmArchitectureStream.Show();
        }

        private void btnPreparePreContract_Click(object sender, EventArgs e)
        {
            frmArchitectureStream frmArchitectureStream = new frmArchitectureStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmArchitectureStream.ArchitectureStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmArchitectureStream.Show();
        }

        private void btnPrepareDetailedDesign_Click(object sender, EventArgs e)
        {
            frmArchitectureStream frmArchitectureStream = new frmArchitectureStream();

            frmArchitectureStream.Show();
        }

        private void btnPrepareTranferImplementation_Click(object sender, EventArgs e)
        {
            frmArchitectureStream frmArchitectureStream = new frmArchitectureStream();

            frmArchitectureStream.Show();
        }

        //PROCUREMENT STREAM
        private void btnRFPResources_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmProcurementStream.ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmProcurementStream.Show();
        }

        private void btnEvaluateRFP_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmProcurementStream.ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmProcurementStream.Show();
        }

        private void btnRFPPrototypes_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmProcurementStream.ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmProcurementStream.Show();
        }

        private void btnPrepareRFP_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmProcurementStream.ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmProcurementStream.Show();
        }

        private void btnGetTender_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();

            frmProcurementStream.Show();
        }

        private void btnIssueRFPs_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmProcurementStream.ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmProcurementStream.Show();
        }

        private void btnEvaluate_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmProcurementStream.ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmProcurementStream.Show();
        }

        private void btnAwardContracts_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmProcurementStream.ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmProcurementStream.Show();
        }

        private void btnFeedbackTender_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmProcurementStream.ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmProcurementStream.Show();
        }

        private void btnConclude_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create a procurement plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmProcurementStream.ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmProcurementStream.Show();
        }

        private void btnIdentifyContracts_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();

            frmProcurementStream.Show();
        }

        private void btnEvaluatePerformance_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit business plan benefit realisation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmProcurementStream.ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmProcurementStream.Show();
        }

        private void btnMonitorCompliance_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Evaluate Technical Delivery.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmProcurementStream.ProcurementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmProcurementStream.Show();            
        }

        private void btnProcessStage_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();

            frmProcurementStream.Show();
        }

        private void btnCloseContracts_Click(object sender, EventArgs e)
        {
            frmProcurementStream frmProcurementStream = new frmProcurementStream();

            frmProcurementStream.Show();
        }

        private void btnDefineBenfits_Click(object sender, EventArgs e)
        {
            frmBenefitsRealisationStream frmBenefitsRealisationStream = new frmBenefitsRealisationStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBenefitsRealisationStream.BenefitsRealisationStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBenefitsRealisationStream.Show();
        }

        private void btnIdentifyEarlyBenefits_Click(object sender, EventArgs e)
        {
            frmBenefitsRealisationStream frmBenefitsRealisationStream = new frmBenefitsRealisationStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBenefitsRealisationStream.BenefitsRealisationStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBenefitsRealisationStream.Show();
        }

        private void btnPrepareBenefits_Click(object sender, EventArgs e)
        {
            frmBenefitsRealisationStream frmBenefitsRealisationStream = new frmBenefitsRealisationStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop a Business Case: Identify'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBenefitsRealisationStream.BenefitsRealisationStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBenefitsRealisationStream.Show();
        }

        private void btnIdentifyBenefits_Click(object sender, EventArgs e)
        {
            frmBenefitsRealisationStream frmBenefitsRealisationStream = new frmBenefitsRealisationStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop a Business Case: Identify'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBenefitsRealisationStream.BenefitsRealisationStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBenefitsRealisationStream.Show();
        }

        private void btnUpdatePlan_Click(object sender, EventArgs e)
        {
            frmBenefitsRealisationStream frmBenefitsRealisationStream = new frmBenefitsRealisationStream();

            frmBenefitsRealisationStream.Show();
        }

        private void btnMonitorBenefit_Click(object sender, EventArgs e)
        {
            frmBenefitsRealisationStream frmBenefitsRealisationStream = new frmBenefitsRealisationStream();

            frmBenefitsRealisationStream.Show();
        }

        private void btnConductPlan_Click(object sender, EventArgs e)
        {
            frmBenefitsRealisationStream frmBenefitsRealisationStream = new frmBenefitsRealisationStream();

            frmBenefitsRealisationStream.Show();
        }

        private void btnConfirmBenefits_Click(object sender, EventArgs e)
        {
            frmBenefitsRealisationStream frmBenefitsRealisationStream = new frmBenefitsRealisationStream();

            frmBenefitsRealisationStream.Show();
        }

        private void btnReviewPlan_Click(object sender, EventArgs e)
        {
            frmBenefitsRealisationStream frmBenefitsRealisationStream = new frmBenefitsRealisationStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit business plan benefit realisation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmBenefitsRealisationStream.BenefitsRealisationStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmBenefitsRealisationStream.Show();
        }

        private void btnProcessStageReport_Click(object sender, EventArgs e)
        {
            frmBenefitsRealisationStream frmBenefitsRealisationStream = new frmBenefitsRealisationStream();

            frmBenefitsRealisationStream.Show();
        }

        private void ProgrammeManagementStreamTabPage_Click(object sender, EventArgs e)
        {

        }

        private void btnProgrammeWorkStreamPlanRemainingStage_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Planning%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamEstimateFunding_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Submit Opportunities to Investment Portfolio Process'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamAquireTask_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();

            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamInitProject_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Initiate high-level project development planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamDefineProject_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Construct a preliminary Business Case and Risk Assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamFundsAllocation_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Submit Opportunities to Investment Portfolio Process'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamConfirmProject_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Submit Opportunities to Investment Portfolio Process'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamDefineInvestment_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop a Business Case:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamRegiterProject_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();

            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamPrepareSupporting_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamReviewBusinessCase_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();

            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamAssessProject_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();

            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamRecordProject_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Close-Out Project'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamAuditBenefit_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Audit Business Plan Benefits'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamCloseProject_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE 'De-Commissioning A Project%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamManagingStageBoundaries_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Management:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamMonitorControlPerformance_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Management:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProgrammeWorkStreamMonitorProgramme_Click(object sender, EventArgs e)
        {
            frmProgrammeManagementStream frmPMS = new frmProgrammeManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Management:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMS.ProgrammeManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMS.Show();
        }

        private void btnProjectWorkStreamPlanNextSt_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Planning%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMSS.ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMSS.Show();
        }

        private void btnProjectWorkStreamAquireTask_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Planning%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMSS.ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMSS.Show();
        }

        private void btnProjectWorkStreamCreateInitDraft_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Implementation Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMSS.ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMSS.Show();
        }

        private void btnProjectWorkStreamCreateFinalPPlan_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Execution Strategy'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMSS.ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMSS.Show();
        }

        private void btnProjectWorkStreamStartingUp_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Project Start Up Definition:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMSS.ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMSS.Show();
        }

        private void btnProjectWorkStreamIdentifyProject_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Develop a Business Case:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMSS.ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMSS.Show();
        }

        private void btnProjectWorkStreamCreateVariousExe_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans : Schedule Project Work'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMSS.ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMSS.Show();
        }

        private void btnProjectWorkStreamArchiveProjectOutputs_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Consolidate Documentation'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMSS.ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMSS.Show();
        }

        private void btnProjectWorkStreamIntegratePlans_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();

            frmPMSS.Show();
        }

        private void btnProjectWorkStreamPlanNextStage_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Planning%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMSS.ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMSS.Show();
        }

        private void btnProjectWorkStreamPrepareStage_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();

            frmPMSS.Show();
        }

        private void btnProjectWorkStreamMonitorControlProject_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Stage Management%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            frmPMSS.ProjectManagementStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
            frmPMSS.Show();
        }

        private void btnProjectWorkStreamMonitorBenefitReal_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();

            frmPMSS.Show();
        }

        private void btnProjectWorkStreamProcessStageEndReport_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();

            frmPMSS.Show();
        }

        private void btnProjectWorkStreamConductBenefitRealisation_Click(object sender, EventArgs e)
        {
            frmProjectManagementStream frmPMSS = new frmProjectManagementStream();

            frmPMSS.Show();
        }


        //BENEFITS REALISATION STREAM
    }
}
