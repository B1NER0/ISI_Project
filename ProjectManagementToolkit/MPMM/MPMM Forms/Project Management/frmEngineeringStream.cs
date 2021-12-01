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
    public partial class frmEngineeringStream : Form
    {
        public frmEngineeringStream()
        {
            InitializeComponent();
        }

        private void btnBackToPortfolioManagementStream_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnIdentifyOpp_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen Opportunities for strategic fit'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnWorkReq_Click(object sender, EventArgs e)
        {

        }

        private void btnStudyBusiness_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen Opportunities for strategic fit'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnAnalyseDifferent_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Business Case'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnEstimateCost_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop a Business Case:'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnHighlevel_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform Basic Design'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnConsiderAlt_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop a Business Case:'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnValidateTechnical_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop a Business Case:'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnAddressLegal_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Execute EIA, Regulatory and Legal'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnNextLevel_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop Engineering Specifications'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnTechinicalRecom_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Procurement management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnDesignComponent_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Product delivery management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnDesignInterface_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Configuration management Implement configuration Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnReviewDesignSpec_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Design Freeze'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPrepareTest_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Create the various execution plans: Create an acceptance plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPrepareBuild_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop prototype'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnDevelopConfig_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Configuration management Implement configuration Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnConfigInterfaces_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Configuration management Implement configuration Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnUnitInterface_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Conduct tests'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnDevelopManuals_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop training concept'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPrepareProduction_Click(object sender, EventArgs e)
        {

        }

        private void btnTrainOperators_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop training concept'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnDeployAsset_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Start-Up and Commissioning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnOwner_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Hand Over / Partial Hand Over'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnCommision_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Start-Up and Commissioning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnCheckGuarantee_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Initiate Plant Lifecycle Plan - O&M'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            EngineeringStreamDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnArchiveOutputs_Click(object sender, EventArgs e)
        {

        }

        private void btnBestPractices_Click(object sender, EventArgs e)
        {

        }
    }
}
