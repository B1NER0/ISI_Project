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
    public partial class frmBusinessPlan : Form
    {
        public frmBusinessPlan()
        {
            InitializeComponent();
        }

        private void btnBCWInitiateStructProject_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Initiate and Structure Project Set-Up'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBCWFinaliseRegulatoryLegalAppr_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Finalise Regulatory and Legal Approvals'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBCWDevEngineeringSpecs_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Develop Engineering Specifications'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBackToBusinessPlan_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnBCWScopeFreeze_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Scope Freeze'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBCWLifecycleOperationPlan_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Lifecycle Operation Plan'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBCWProjectFinManagement_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Financial management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBCWResolveCommercialFinancialIssues_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Resolve Commercial and Financial Issues'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBCWDetailBusinessPlanDeveloped_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Detailed Business Plan Developed'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();

        }

        private void btnBCWDataHandOverTExecution_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Data Hand-Over to Execution'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBCWProjectProcurementManagement_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Procurement management'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBCWPrepareDistributeRFI_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package LIKE '%Prepare and Distribute RFI%'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }

        private void btnBCWProjectExecutionStrat_Click(object sender, EventArgs e)
        {
            string fileName = "PLMSWorkPackages.xlsx";
            string path = Path.Combine(Environment.CurrentDirectory, fileName);
            OleDbConnection myconnection = new OleDbConnection(@"Provider = Microsoft.ACE.OLEDB.12.0; Data Source='" + path + "';Extended Properties='Excel 12.0;HDR = YES';");
            OleDbDataAdapter oda = new OleDbDataAdapter("SELECT * FROM [sheet1$] WHERE Work_Package = 'Project Execution Strategy'", myconnection);
            DataSet ds = new DataSet();
            oda.Fill(ds);
            BusinessPlanDGV.DataSource = ds.Tables[0];
            myconnection.Close();
        }
    }
}
