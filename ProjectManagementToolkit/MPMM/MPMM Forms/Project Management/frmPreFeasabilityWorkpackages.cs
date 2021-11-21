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
    public partial class frmPreFeasabilityWorkpackages : Form
    {
        public frmPreFeasabilityWorkpackages()
        {
            InitializeComponent();
        }

        private void btnPreFeasReturnToPreFeas_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPreFeasWProjectBrief_Click(object sender, EventArgs e)
        {

            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Definition : Project Brief'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasEstablishCoreProjectTeam_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish the Core Project Team.'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasAppointProjectTeam_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Appoint the project team:%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasPerformInfrastructAssess_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform base infrastructure assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasDeterLeaglReq_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Determine Legal and Regulatory requirements'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasPerformStakeholderNeeds_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform stakeholder needs impact assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasInitiateHighLevelProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Initiate high-level project development planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasPerformConceptDesign_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Perform Concept Design'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasPartnerSelection_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Partner Selection'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasIdFinCommercialStruct_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Identify financial and commercial structures'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasConstrPriliBusinessCase_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Construct a preliminary Business Case and Risk Assessment'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasScreenPriorDisOpp_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, @"Data\", fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Screen, prioritise and discard opportunities'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            PreFeasabilityDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnPreFeasConcludeStage_Click(object sender, EventArgs e)
        {

        }
    }
}
