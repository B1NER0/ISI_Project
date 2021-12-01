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
    public partial class frmProjectExecutionPlanning : Form
    {
        public frmProjectExecutionPlanning()
        {
            InitializeComponent();
        }

        private void btnBackToProjectExecution_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnProjectExeWProjectStartUp_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Start-Up'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWEstProjectExe_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Establish Project Execution Org. Structure'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWCreateExePlans_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a procurement plan%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWCreateExePlansRiskPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a risk plan%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWCreateExePlansScopeWBS_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans : Scope, WBS and resources Planning%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWCreateExePlansFinPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a financial plan%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWCreateExePlansQualityPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a quality plan%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWCreateExePlansScheduleProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans : Schedule Project Work'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWCreateExePlansCommunicationPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a Communications plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWCreateExePlansChangeManagementPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create a change management plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWCompileBalanceEnquiries_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Compile Balance of Enquiries'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWConfigManagement_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Configuration management Implement configuration Planning'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWCreateExePlansAcceptancePlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Create the various execution plans: Create an acceptance plan%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWDesignFreeze_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Design Freeze'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWConceptDev_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Concept Development%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWUpdateProjectImplementationPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Update Project Implementation Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnProjectExeWProjectImplementationPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Implementation Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            ProjectExecutionDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }
    }
}
